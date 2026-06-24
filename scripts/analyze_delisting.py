from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


TARGET_REVENUE_SHEET = "收入及直接成本明细表"
PRODUCT_LIST_SHEETS = ("自研类", "引入类")
ACTIVE_STATUSES_FOR_RULES_1_TO_3 = {"已上市", "已入库"}
AS_OF_DATE = date(2026, 6, 23)
OLDER_THAN_TWO_YEARS_BEFORE = date(2024, 6, 23)
REVENUE_START_MONTH = "2024-06"
REVENUE_END_MONTH = "2026-05"
ERROR_LITERALS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}
OUTPUT_HEADERS = [
    "产品编码",
    "产品名称",
    "产品状态",
    "创建时间",
    "部门",
    "产品收入",
    "产品毛利",
    "退市类型",
    "理由",
]
RULE_TEXT = {
    "1": "已上市/已入库但未出现在收入及直接成本明细表",
    "2": "创建时间超过2年或为空，且2年内产品收入为0",
    "3": "创建时间超过2年或为空，且2年内产品毛利≤0",
    "4": "产品状态为退市中（暂不判断超过1年）",
}


def normalize(value: Any) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return "" if text in ERROR_LITERALS else text


def parse_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = normalize(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def month_key(value: Any) -> str:
    parsed = parse_date(value)
    if parsed:
        return parsed.strftime("%Y-%m")
    return normalize(value).replace("/", "-")[:7]


def to_decimal(value: Any) -> Decimal:
    if value is None or value == "" or isinstance(value, bool):
        return Decimal("0")
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return Decimal("0")


def decimal_to_number(value: Decimal) -> int | float:
    rounded = value.quantize(Decimal("0.01"))
    return int(rounded) if rounded == rounded.to_integral() else float(rounded)


def first_existing(headers: list[str], names: list[str]) -> int | None:
    for name in names:
        if name in headers:
            return headers.index(name)
    return None


def require_index(header_map: dict[str, int], name: str, sheet_name: str) -> int:
    if name not in header_map:
        raise ValueError(f"{sheet_name} 缺少必要字段：{name}")
    return header_map[name]


def read_product_list(product_list_path: Path) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    workbook = load_workbook(product_list_path, read_only=False, data_only=True)
    products: dict[str, dict[str, Any]] = {}
    status_counter: Counter[str] = Counter()
    sheet_counts: dict[str, int] = {}

    for sheet_name in PRODUCT_LIST_SHEETS:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"产品全量列表缺少 sheet：{sheet_name}")
        sheet = workbook[sheet_name]
        headers = [normalize(sheet.cell(1, col).value) for col in range(1, sheet.max_column + 1)]
        header_map = {header: idx for idx, header in enumerate(headers) if header}
        code_idx = require_index(header_map, "产品编码", sheet.title)
        name_idx = require_index(header_map, "产品名称", sheet.title)
        status_idx = require_index(header_map, "产品状态", sheet.title)
        created_idx = require_index(header_map, "产品创建时间", sheet.title)
        department_idx = require_index(header_map, "产品所属部门", sheet.title)
        start_row = 4 if sheet.title == "自研类" else 2
        sheet_count = 0

        for row_index in range(start_row, sheet.max_row + 1):
            row = [sheet.cell(row_index, col).value for col in range(1, sheet.max_column + 1)]
            code = normalize(row[code_idx])
            name = normalize(row[name_idx])
            if not code and not name:
                continue
            if not code:
                raise ValueError(f"{sheet.title} 第 {row_index} 行缺少产品编码，无法按产品编码匹配")
            if code in products:
                raise ValueError(f"全量产品清单存在重复产品编码：{code}")
            status = normalize(row[status_idx])
            products[code] = {
                "产品编码": code,
                "产品名称": name,
                "产品状态": status,
                "创建时间": normalize(row[created_idx]),
                "创建日期": parse_date(row[created_idx]),
                "部门": normalize(row[department_idx]),
                "来源sheet": sheet.title,
                "来源行号": row_index,
            }
            status_counter[status] += 1
            sheet_count += 1
        sheet_counts[sheet.title] = sheet_count

    return products, {
        "product_count": len(products),
        "status_counts": dict(status_counter),
        "sheet_counts": sheet_counts,
    }


def read_revenue(revenue_paths: list[Path]) -> tuple[dict[str, dict[str, Any]], set[str], list[dict[str, Any]]]:
    aggregates = defaultdict(lambda: {"产品收入": Decimal("0"), "产品毛利": Decimal("0"), "all_rows": 0, "window_rows": 0})
    all_revenue_codes: set[str] = set()
    workbook_summaries: list[dict[str, Any]] = []

    for workbook_path in revenue_paths:
        workbook = load_workbook(workbook_path, read_only=True, data_only=True)
        if TARGET_REVENUE_SHEET not in workbook.sheetnames:
            raise ValueError(f"{workbook_path.name} 缺少 sheet：{TARGET_REVENUE_SHEET}")
        sheet = workbook[TARGET_REVENUE_SHEET]
        headers = [normalize(value) for value in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
        code_idx = first_existing(headers, ["产品编码"])
        revenue_idx = first_existing(headers, ["产品收入", "金额"])
        gross_profit_idx = first_existing(headers, ["产品毛利"])
        month_idx = first_existing(headers, ["收入月份"])
        missing = [
            label
            for label, idx in {
                "产品编码": code_idx,
                "产品收入/金额": revenue_idx,
                "产品毛利": gross_profit_idx,
                "收入月份": month_idx,
            }.items()
            if idx is None
        ]
        if missing:
            raise ValueError(f"{workbook_path.name} 缺少必要字段：{', '.join(missing)}")

        source_rows = 0
        window_rows = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            source_rows += 1
            code = normalize(row[code_idx])
            if not code:
                continue
            all_revenue_codes.add(code)
            aggregates[code]["all_rows"] += 1
            current_month = month_key(row[month_idx])
            if REVENUE_START_MONTH <= current_month <= REVENUE_END_MONTH:
                window_rows += 1
                aggregates[code]["window_rows"] += 1
                aggregates[code]["产品收入"] += to_decimal(row[revenue_idx])
                aggregates[code]["产品毛利"] += to_decimal(row[gross_profit_idx])

        workbook_summaries.append(
            {
                "file": workbook_path.name,
                "source_rows": source_rows,
                "window_rows": window_rows,
                "revenue_column": headers[revenue_idx],
            }
        )
    return aggregates, all_revenue_codes, workbook_summaries


def build_candidates(products: dict[str, dict[str, Any]], revenue: dict[str, dict[str, Any]], all_revenue_codes: set[str]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for code, product in products.items():
        status = product["产品状态"]
        revenue_values = revenue[code]
        product_revenue = revenue_values["产品收入"]
        product_gross_profit = revenue_values["产品毛利"]
        old_or_blank = product["创建日期"] is None or product["创建日期"] < OLDER_THAN_TWO_YEARS_BEFORE
        active = status in ACTIVE_STATUSES_FOR_RULES_1_TO_3
        rules: list[str] = []

        if active and code not in all_revenue_codes:
            rules.append("1")
        if active and old_or_blank and product_revenue == 0:
            rules.append("2")
        if active and old_or_blank and product_gross_profit <= 0:
            rules.append("3")
        if status == "退市中":
            rules.append("4")
        if not rules:
            continue

        delisting_type = "强制退市" if {"1", "2"} & set(rules) else "建议退市"
        candidates.append(
            {
                "产品编码": product["产品编码"],
                "产品名称": product["产品名称"],
                "产品状态": status,
                "创建时间": product["创建时间"],
                "部门": product["部门"],
                "产品收入": decimal_to_number(product_revenue),
                "产品毛利": decimal_to_number(product_gross_profit),
                "退市类型": delisting_type,
                "理由": "；".join(RULE_TEXT[rule] for rule in rules),
                "命中规则": ",".join(rules),
                "来源sheet": product["来源sheet"],
                "来源行号": product["来源行号"],
            }
        )

    type_rank = {"强制退市": 0, "建议退市": 1}
    candidates.sort(key=lambda row: (type_rank[row["退市类型"]], row["命中规则"], row["产品编码"]))
    return candidates


def make_payload(product_list_path: Path, revenue_paths: list[Path]) -> dict[str, Any]:
    products, product_summary = read_product_list(product_list_path)
    revenue, all_revenue_codes, workbook_summaries = read_revenue(revenue_paths)
    audit_rows = build_candidates(products, revenue, all_revenue_codes)
    rule_counts: Counter[str] = Counter()
    type_counts: Counter[str] = Counter()
    candidate_status_counts: Counter[str] = Counter()
    for row in audit_rows:
        type_counts[row["退市类型"]] += 1
        candidate_status_counts[row["产品状态"]] += 1
        for rule in row["命中规则"].split(","):
            rule_counts[rule] += 1

    return {
        "headers": OUTPUT_HEADERS,
        "metadata": {
            "as_of_date": AS_OF_DATE.isoformat(),
            "older_than_two_years_before": OLDER_THAN_TWO_YEARS_BEFORE.isoformat(),
            "revenue_window": {"start_month": REVENUE_START_MONTH, "end_month": REVENUE_END_MONTH},
            "active_statuses_for_rules_1_to_3": sorted(ACTIVE_STATUSES_FOR_RULES_1_TO_3),
            "product_summary": product_summary,
            "revenue_workbooks": workbook_summaries,
            "all_revenue_code_count": len(all_revenue_codes),
            "candidate_count": len(audit_rows),
            "delisting_type_counts": dict(type_counts),
            "candidate_status_counts": dict(candidate_status_counts),
            "rule_counts": dict(rule_counts),
        },
        "rows": [{header: row[header] for header in OUTPUT_HEADERS} for row in audit_rows],
        "audit_rows": audit_rows,
    }


def add_table(sheet, ref: str, name: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    sheet.add_table(table)


def write_output_workbook(payload: dict[str, Any], output_path: Path) -> None:
    workbook = Workbook()
    result = workbook.active
    result.title = "退市筛选结果"
    result.append(OUTPUT_HEADERS)
    for row in payload["rows"]:
        result.append([row.get(header, "") for header in OUTPUT_HEADERS])
    widths = [18, 34, 12, 22, 22, 16, 16, 14, 82]
    for index, width in enumerate(widths, start=1):
        result.column_dimensions[result.cell(1, index).column_letter].width = width
    for cell in result[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in result.iter_rows(min_row=2, min_col=6, max_col=7):
        for cell in row:
            cell.number_format = "#,##0.00"
    for row in result.iter_rows(min_row=2, min_col=8, max_col=9):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    result.freeze_panes = "A2"
    result.auto_filter.ref = result.dimensions
    add_table(result, result.dimensions, "DelistingScreeningTable")

    summary = workbook.create_sheet("口径说明")
    metadata = payload["metadata"]
    summary_rows = [
        ["项目", "说明"],
        ["基准日期", metadata["as_of_date"]],
        ["2年创建时间阈值", f"早于 {metadata['older_than_two_years_before']}，或创建时间为空"],
        ["2年经营数据窗口", f"{metadata['revenue_window']['start_month']} 至 {metadata['revenue_window']['end_month']}"],
        ["规则1/2/3状态范围", "、".join(metadata["active_statuses_for_rules_1_to_3"])],
        ["规则4状态范围", "退市中；暂不判断超过1年"],
        ["候选总数", metadata["candidate_count"]],
        ["强制退市数量", metadata["delisting_type_counts"].get("强制退市", 0)],
        ["建议退市数量", metadata["delisting_type_counts"].get("建议退市", 0)],
        ["规则1命中数", metadata["rule_counts"].get("1", 0)],
        ["规则2命中数", metadata["rule_counts"].get("2", 0)],
        ["规则3命中数", metadata["rule_counts"].get("3", 0)],
        ["规则4命中数", metadata["rule_counts"].get("4", 0)],
        ["全量产品数量", metadata["product_summary"]["product_count"]],
        ["全量产品状态分布", "；".join(f"{k}:{v}" for k, v in metadata["product_summary"]["status_counts"].items())],
        ["收入明细文件", "；".join(f"{item['file']}（{item['revenue_column']}，窗口内{item['window_rows']}行）" for item in metadata["revenue_workbooks"])],
    ]
    for row in summary_rows:
        summary.append(row)
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 120
    for cell in summary[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    for row in summary.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
    summary.freeze_panes = "A2"
    add_table(summary, summary.dimensions, "ScreeningBasisTable")

    rules = workbook.create_sheet("规则汇总")
    rule_rows = [
        ["规则", "命中数量", "退市类型影响", "说明"],
        ["规则1", metadata["rule_counts"].get("1", 0), "强制退市", RULE_TEXT["1"]],
        ["规则2", metadata["rule_counts"].get("2", 0), "强制退市", RULE_TEXT["2"]],
        ["规则3", metadata["rule_counts"].get("3", 0), "建议退市", RULE_TEXT["3"]],
        ["规则4", metadata["rule_counts"].get("4", 0), "建议退市", RULE_TEXT["4"]],
    ]
    for row in rule_rows:
        rules.append(row)
    for col, width in zip("ABCD", [12, 12, 16, 72]):
        rules.column_dimensions[col].width = width
    for cell in rules[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
    rules.freeze_panes = "A2"
    add_table(rules, rules.dimensions, "RuleSummaryTable")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--product-list", required=True)
    parser.add_argument("--revenue-files", nargs="+", required=True)
    parser.add_argument("--output-json", required=True)
    parser.add_argument("--output-xlsx", required=True)
    args = parser.parse_args()

    payload = make_payload(Path(args.product_list), [Path(item) for item in args.revenue_files])
    output_json = Path(args.output_json)
    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    write_output_workbook(payload, Path(args.output_xlsx))
    print(json.dumps({"ok": True, "metadata": payload["metadata"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
