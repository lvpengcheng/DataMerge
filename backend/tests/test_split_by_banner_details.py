"""人员账单 Sheet 拆分回归测试。"""

import openpyxl

from split_by_banner import _scan_person_detail_blocks


def test_scan_normal_and_retro_person_detail_blocks():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "账单"

    ws.append(["账单标题"])
    ws.append(["公司名称", None, "账单号"])
    ws.append(["序号", "姓名", "月份"])
    ws.append(["S/N", "Name", "YYYYMM"])
    ws.append([1, "张三", "202608"])
    ws.append([2, "李四", "202608"])
    ws.append(["合计", None, None])
    ws.append(["历月补收明细"])
    ws.append(["序号", "姓名", "月份"])
    ws.append(["S/N", "Name", "YYYYMM"])
    ws.append([1, "王五", "202607"])
    ws.append(["合计", None, None])
    ws.append(["服务费明细"])
    ws.append(["S/N", "YYYYMM", "Item"])
    ws.append([1, "202608", "服务费"])

    assert _scan_person_detail_blocks(ws) == [
        {
            "title": "正常数据",
            "header_start": 3,
            "header_end": 4,
            "data_start": 5,
            "data_end": 6,
        },
        {
            "title": "历月补收明细",
            "header_start": 9,
            "header_end": 10,
            "data_start": 11,
            "data_end": 11,
        },
    ]
