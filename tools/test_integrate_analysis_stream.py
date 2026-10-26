"""Exercise the gateway heartbeat without loading database or Excel runtimes."""
import ast
import asyncio
import json
import logging
from pathlib import Path
from types import SimpleNamespace


class HTTPException(Exception):
    def __init__(self, detail, status_code=400):
        self.detail = detail
        self.status_code = status_code


source = Path(__file__).resolve().parents[1] / 'backend/api/tools.py'
tree = ast.parse(source.read_text(encoding='utf-8'))
endpoint = next(node for node in tree.body if isinstance(node, ast.AsyncFunctionDef)
                and node.name == 'integrate_analyze')
stream_default = dict(zip(
    [arg.arg for arg in endpoint.args.args[-len(endpoint.args.defaults):]],
    endpoint.args.defaults))['stream']
assert isinstance(stream_default, ast.Call) and stream_default.args[0].value is True, (
    'Cached clients that omit stream must receive gateway heartbeats')
template = source.parents[2] / 'frontend/templates/tools.html'
assert 'tools.js?v=20261023-apply-heartbeat' in template.read_text(encoding='utf-8')
function = next(node for node in tree.body if isinstance(node, ast.AsyncFunctionDef)
                and node.name == '_stream_integrate_analysis')


async def fast_wait(tasks, timeout):
    return await asyncio.wait(tasks, timeout=0.001)


scope = dict(asyncio=SimpleNamespace(create_task=asyncio.create_task, wait=fast_wait,
                                    gather=asyncio.gather),
             json=json, HTTPException=HTTPException, logger=logging.getLogger(__name__))
exec(compile(ast.Module(body=[function], type_ignores=[]), str(source), 'exec'), scope)
stream = scope['_stream_integrate_analysis']


async def main():
    async def success():
        await asyncio.sleep(0.02)
        return {'session_id': 'test', 'files': []}

    chunks = [chunk async for chunk in stream(success())]
    assert len(chunks) > 2, 'Expected periodic heartbeats before result'
    assert json.loads(''.join(chunks))['session_id'] == 'test'

    async def failure():
        raise HTTPException('解析超时')

    assert json.loads(''.join([part async for part in stream(failure())])) == {'error': '解析超时', 'status_code': 400}
    cancelled = asyncio.Event()

    async def pending():
        try:
            await asyncio.Event().wait()
        finally:
            cancelled.set()

    response = stream(pending())
    await anext(response)
    await asyncio.sleep(0)
    await response.aclose()
    assert cancelled.is_set(), 'Disconnect must cancel parsing task'
    print('PASS: heartbeat, JSON result, parse error, disconnect cancellation')


if __name__ == '__main__':
    asyncio.run(main())
