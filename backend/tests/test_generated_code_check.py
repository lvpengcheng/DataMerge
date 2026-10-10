from backend.utils.generated_code_check import check_generated_code


def test_missing_closure_record_and_helper_are_found():
    issues = check_generated_code('def fill():\n    def put():\n        return prev_rec.get("含 空格(字段)")\n    missing_helper()\n')
    assert any('prev_rec' in x['message'] for x in issues)
    assert any('missing_helper' in x['message'] for x in issues)


def test_forward_function_and_injected_names_are_valid():
    assert not check_generated_code('def main():\n    return helper(source_data)\ndef helper(x):\n    return x["含 空格(字段)"]\n', {'source_data'})


def test_unbound_local_detected():
    assert check_generated_code('value = 1\ndef main():\n    print(value)\n    value = 2\n')


def test_precise_edit_dependency_group_is_atomic(monkeypatch):
    from backend.ai_engine import precise_edit
    monkeypatch.setattr(precise_edit, '_call_provider', lambda *a: '{"edits":[{"find":"return 1","replace":"return helper()"},{"find":"missing anchor","replace":"def helper(): return 2"}]}')
    reasons = []
    result = precise_edit.run_precise_edit(object(), 'def main():\n    return 1\n', '修改', reason_sink=reasons)
    assert result is None and reasons


def test_precise_edit_rejects_new_undefined_name(monkeypatch):
    from backend.ai_engine import precise_edit
    monkeypatch.setattr(precise_edit, '_call_provider', lambda *a: '{"edits":[{"find":"return 1","replace":"return missing_helper()"}]}')
    reasons = []
    assert precise_edit.run_precise_edit(object(), 'def main():\n    return 1\n', '修改', reason_sink=reasons) is None
    assert 'missing_helper' in reasons[0]


def test_validation_feedback_repairs_edit_once(monkeypatch):
    from backend.ai_engine import precise_edit
    prompts = []
    def reply(provider, prompt, *args):
        prompts.append(prompt)
        replacement = 'return missing_helper()' if len(prompts) == 1 else 'return 2'
        import json
        return json.dumps({'edits': [{'find': 'return 1', 'replace': replacement}]})
    monkeypatch.setattr(precise_edit, '_call_provider', reply)
    result = precise_edit.run_precise_edit(object(), 'def main():\n    return 1\n', '修改')
    assert result == 'def main():\n    return 2\n'
    assert len(prompts) == 2 and 'undefined name' in prompts[1]
