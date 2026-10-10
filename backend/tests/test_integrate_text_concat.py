import pytest
from backend.utils.integrate_engine import eval_source_expr_cross, eval_source_expr, validate_formula_remainder


def sources():
    return {
        'A.xlsx': {'cols': ['姓名', '工号', '金额'], 'rows': {'k': [
            {'姓名': '张三', '工号': '0012', '金额': 10},
            {'姓名': '李四', '工号': '0099', '金额': 20}]}},
        'B.xlsx': {'cols': ['部门', '补贴'], 'rows': {'k': [{'部门': '研发', '补贴': 5}]}},
    }


@pytest.mark.parametrize('expr,expected', [
    ('姓名&B.xlsx.部门', '张三研发'),
    ('=姓名&" / "&B.xlsx.部门', '张三 / 研发'),
    ('工号&"-"&B.xlsx.部门', '0012-研发'),
    ('姓名&"="&(金额+B.xlsx.补贴)', '张三=35'),
    ('IF(金额>0,姓名&" / "&B.xlsx.部门,"")', '张三 / 研发'),
    ('姓名&"姓名<>IF("&B.xlsx.部门', '张三姓名<>IF(研发'),
    ('姓名&"说""你好"""&B.xlsx.部门', '张三说"你好"研发'),
    ('ROUND(金额+B.xlsx.补贴,2)', 35),
    ('姓名&不存在', None),
    ('姓名&&B.xlsx.部门', None),
    ('姓名&__import__("os")', None),
])
def test_cross_concat(expr, expected):
    assert eval_source_expr_cross(expr, 'A.xlsx', sources(), 'k') == expected


def test_missing_and_single_table():
    data = sources()
    data['B.xlsx']['rows'] = {}
    assert eval_source_expr_cross('姓名&B.xlsx.部门', 'A.xlsx', data, 'k') == '张三'
    assert eval_source_expr_cross('姓名&B.xlsx.部门', 'A.xlsx', data, 'missing') is None
    assert eval_source_expr('姓名&工号', data['A.xlsx']['rows']['k'], data['A.xlsx']['cols']) == '张三0012'
    assert validate_formula_remainder(' & " / 中文" & ')
