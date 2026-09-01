#!/usr/bin/env python3
"""Generate a game-test-design testcase workbook.

Usage:
  python gen_testcase_xlsx.py -o cases.xlsx
  python gen_testcase_xlsx.py -i cases.json -o cases.xlsx

The JSON input is an array of objects. Unknown keys are ignored; missing
columns are emitted as blank cells so the file remains a usable skeleton.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


COLUMNS = [
    "用例ID", "模块/功能", "状态节点", "前置条件", "操作步骤", "预期结果",
    "测试类型", "打断层级", "影响面", "风险级别", "优先级", "备注",
]
LEVELS = ["无", "L1", "L2", "L3", "L4", "L5", "L6-a", "L6-b", "L7"]
IMPACTS = ["自身", "队友", "敌对玩家", "中立单位", "环境场景", "观战/旁观者", "服务器全局", "AI/召唤物"]


def load_rows(path: Path | None):
    if path is None:
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list) or any(not isinstance(row, dict) for row in data):
        raise ValueError("输入 JSON 必须是对象数组")
    return data


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "TestCases"
    ws.append(COLUMNS)
    for idx, row in enumerate(rows, 1):
        values = [row.get(col, "") for col in COLUMNS]
        if not values[0]:
            values[0] = f"GT-{idx:03d}"
        ws.append(values)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = [12, 18, 18, 28, 36, 42, 16, 14, 20, 12, 10, 30]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Keep the framework vocabulary visible and reusable in the workbook.
    dictionary = wb.create_sheet("Dictionaries")
    dictionary.append(["打断层级", "影响面"])
    max_len = max(len(LEVELS), len(IMPACTS))
    for i in range(max_len):
        dictionary.append([
            LEVELS[i] if i < len(LEVELS) else "",
            IMPACTS[i] if i < len(IMPACTS) else "",
        ])
    for cell in dictionary[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    dictionary.column_dimensions["A"].width = 14
    dictionary.column_dimensions["B"].width = 18

    wb.defined_names.add(
        DefinedName("InterruptionLevels", attr_text="'Dictionaries'!$A$2:$A$10")
    )
    level_validation = DataValidation(
        type="list", formula1="=InterruptionLevels", allow_blank=True
    )
    ws.add_data_validation(level_validation)
    level_validation.add("H2:H1048576")
    return wb


def main():
    parser = argparse.ArgumentParser(description="生成游戏测试用例 Excel 骨架")
    parser.add_argument("-i", "--input", type=Path, help="用例 JSON 文件（对象数组）")
    parser.add_argument("-o", "--output", type=Path, required=True, help="输出 .xlsx 路径")
    args = parser.parse_args()
    rows = load_rows(args.input)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(rows).save(args.output)
    print(f"已生成: {args.output}（{len(rows)} 条用例）")


if __name__ == "__main__":
    main()
