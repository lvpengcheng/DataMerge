"""Isolated tests for scheme heartbeat routing and deferred binary downloads."""
import ast
import asyncio
import json
import logging
import re
import shutil
import tempfile
import unittest
import uuid
from pathlib import Path
from types import SimpleNamespace


class HTTPException(Exception):
    def __init__(self, status_code, detail):
        self.status_code, self.detail = status_code, detail


def response(path=None, **kwargs):
    return SimpleNamespace(path=path, **kwargs)


class ApplyHeartbeatTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.user = SimpleNamespace(id=1)
        self.req = SimpleNamespace(session_id='a' * 32)
        self.calls = []

        async def work(req, user, workflow):
            self.calls.append(workflow)
            return {'ok': True}

        async def execute(req, user, workflow):
            self.calls.append(workflow)
            path = self.root / ('split/result.zip' if workflow == 'split' else 'result.xlsx')
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b'binary-result')
            return response(path, media_type='application/zip' if workflow == 'split' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={'x-integrate-matched': '12', 'content-disposition': 'attachment; filename=result'})

        self.scope = dict(asyncio=asyncio, json=json, logging=logging, logger=logging.getLogger(__name__),
            Path=Path, uuid=uuid, re=re, shutil=shutil, HTTPException=HTTPException,
            StreamingResponse=lambda body, **kw: response(body_iterator=body, **kw), FileResponse=response,
            BackgroundTask=lambda fn: fn, _workflow_session_dir=lambda *args: self.root,
            _integrate_execute_response=execute, _integrate_reparse_headers_impl=work,
            _integrate_scheme_apply_validate_impl=work)
        names = {'integrate_execute', 'integrate_reparse_headers', 'integrate_scheme_apply_validate',
                 '_integration_heartbeat_response', '_stream_integrate_analysis',
                 '_integrate_prepare_download', 'integrate_download_result'}
        tree = ast.parse((Path(__file__).resolve().parents[1] / 'backend/api/tools.py').read_text(encoding='utf-8'))
        selected = []
        for node in tree.body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in names:
                node.decorator_list = []
                node.returns = None
                for arg in node.args.args:
                    arg.annotation = None
                node.args.defaults = [ast.Constant(None) for _ in node.args.defaults]
                selected.append(node)
        module = ast.fix_missing_locations(ast.Module(body=selected, type_ignores=[]))
        exec(compile(module, '<heartbeat routes>', 'exec'), self.scope)

    async def test_scheme_and_header_routes_stream_for_both_workflows(self):
        for workflow in ('integrate', 'split'):
            for name in ('integrate_reparse_headers', 'integrate_scheme_apply_validate'):
                result = await self.scope[name](self.req, self.user, workflow, True)
                self.assertEqual(result.headers['X-Accel-Buffering'], 'no')
                chunks = [part async for part in result.body_iterator]
                self.assertTrue(chunks[0].isspace())
                self.assertEqual(json.loads(''.join(chunks)), {'ok': True})
                self.assertEqual(await self.scope[name](self.req, self.user, workflow, False), {'ok': True})

    async def test_execute_stream_then_binary_download_and_cleanup(self):
        for workflow in ('integrate', 'split'):
            source = self.root / 'source.xlsx'
            source.write_bytes(b'keep-upload')
            result = await self.scope['integrate_execute'](self.req, self.user, workflow, True)
            data = json.loads(''.join([part async for part in result.body_iterator]))
            prefix = 'table-split' if workflow == 'split' else 'integrate'
            self.assertTrue(data['download_url'].startswith(f'/api/tools/{prefix}/result?'))
            token = data['download_url'].split('token=')[1]
            download = self.scope['integrate_download_result']
            with self.assertRaises(HTTPException) as error:
                await download(self.req.session_id, token, SimpleNamespace(id=2), workflow)
            self.assertEqual(error.exception.status_code, 403)
            file = await download(self.req.session_id, token, self.user, workflow)
            self.assertEqual(file.path.read_bytes(), b'binary-result')
            self.assertEqual(file.headers['x-integrate-matched'], '12')
            self.assertNotIn(b' ', file.path.read_bytes())
            file.background()
            self.assertFalse(file.path.exists())
            self.assertEqual(source.read_bytes(), b'keep-upload')
            with self.assertRaises(HTTPException) as error:
                await download(self.req.session_id, token, self.user, workflow)
            self.assertEqual(error.exception.status_code, 404)

    async def test_invalid_download_credentials(self):
        with self.assertRaises(HTTPException) as error:
            await self.scope['integrate_download_result']('../escape', 'b' * 32, self.user, 'integrate')
        self.assertEqual(error.exception.status_code, 400)


if __name__ == '__main__':
    unittest.main()
