"""Local XLSX reconciliation. All financial comparisons use Decimal."""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import re
import tempfile
import time as clock_time
from typing import Callable
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import ColorFilter, FilterColumn


RED = "FFFFC7CE"
FUEL_ALIASES = {
    "gun": ["枪号", "油枪号"], "time": ["交易时间", "加油时间"],
    "oil": ["油品", "油品名称"], "volume": ["加油量", "升数", "加油升数"],
    "trade": ["交易金额"], "discount": ["优惠金额"],
    "station": ["油站名称", "加油站名称", "油站"], "card": ["卡号", "加油卡号"],
}
INVOICE_ALIASES = {
    "time": ["开票时间", "开票日期"], "amount": ["合计金额（含税）", "价税合计", "合计金额(含税)"],
    "remark": ["备注"], "oil": ["商品名称", "油品"], "volume": ["数量", "升数"],
    "electronic": ["数电票号码", "电子发票号码"], "number": ["发票号码"],
    "flag": ["红票标志"], "invoice_status": ["开票状态"],
    "station": ["油站名称", "加油站名称", "开票门店", "部门门店"],
}


class ReconciliationError(ValueError):
    """A problem with the user's workbook or settings."""


def _replace_file(source, target):
    # Windows scanners can briefly hold a newly created XLSX or reservation.
    # Retry only transient sharing/access errors, without changing the target.
    for attempt in range(5):
        try:
            os.replace(source, target)
            return
        except PermissionError as exc:
            if getattr(exc, "winerror", None) not in (5, 32, 33) or attempt == 4:
                raise
            clock_time.sleep(0.05 * (attempt + 1))


@dataclass(frozen=True)
class Period:
    start: datetime | None
    end: datetime | None
    label: str

    def contains(self, value: datetime | None) -> bool:
        return self.start is None or value is None or self.start <= value <= self.end


def parse_period(date_text="2026-10-11", start_text=None, end_text=None) -> Period:
    if start_text is None and end_text is None:
        legacy = str(date_text or "").strip()
        if not legacy:
            return Period(None, None, "全部日期")
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", legacy):
            raise ReconciliationError("单日日期请使用 YYYY-MM-DD 格式；时间段请填写开始时间和结束时间")
        start_text = end_text = legacy
    start_text, end_text = str(start_text or "").strip(), str(end_text or "").strip()
    if not start_text and not end_text:
        return Period(None, None, "全部日期")
    if not start_text or not end_text:
        raise ReconciliationError("请同时填写开始时间和结束时间，或同时留空核对全部")

    def endpoint(text, is_end):
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2}(?:\.\d{1,6})?)?)?", text):
            raise ReconciliationError("开始/结束时间请使用 YYYY-MM-DD 或 YYYY-MM-DD HH:MM[:SS]")
        try:
            if len(text) == 10:
                return datetime.combine(date.fromisoformat(text), time.max if is_end else time.min)
            return datetime.fromisoformat(text)
        except ValueError as exc:
            raise ReconciliationError("开始/结束时间不是有效的日期时间") from exc

    start, end = endpoint(start_text, False), endpoint(end_text, True)
    if start > end:
        raise ReconciliationError("开始时间不得晚于结束时间")
    label = start_text if start_text == end_text and len(start_text) == 10 else f"{start_text}至{end_text}"
    return Period(start, end, label)


@dataclass
class Record:
    source: str
    row: int
    kind: str
    time: datetime | None
    gun: str
    oil: str
    volume: Decimal | None
    amount: Decimal | None
    number: str = ""
    time_source: str = "交易时间"
    issues: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    in_scope: bool = True
    station: str = ""
    card: str = ""
    trade: Decimal | None = None
    discount: Decimal | None = None
    end_time: datetime | None = None
    invoice_time: datetime | None = None


def decimal_value(value, label: str, *, blank_zero=False) -> Decimal:
    if value is None or str(value).strip() == "":
        if blank_zero:
            return Decimal(0)
        raise ReconciliationError(f"{label}为空")
    if isinstance(value, bool):
        raise ReconciliationError(f"{label}不是有效数字")
    try:
        result = Decimal(str(value).strip().replace(",", "").replace("，", ""))
    except InvalidOperation as exc:
        raise ReconciliationError(f"{label}不是有效数字：{value}") from exc
    if not result.is_finite():
        raise ReconciliationError(f"{label}必须为有限数字")
    return result


def parse_time(value) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if value is None:
        return None
    text = str(value).strip().replace("年", "-").replace("月", "-").replace("日", " ")
    text = text.replace("/", "-").replace("T", " ").replace("：", ":")
    match = re.search(r"(?<!\d)(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,6}))?)?(?=$|\s|[;；,，])", text)
    if not match:
        return None
    y, m, d, h, minute, second, micro = match.groups()
    try:
        return datetime(int(y), int(m), int(d), int(h or 0), int(minute or 0),
                        int(second or 0), int((micro or "0").ljust(6, "0")))
    except ValueError:
        return None


def normalize_gun(value) -> str:
    text = str(value or "").strip()
    match = re.fullmatch(r"(?:油?枪(?:号)?\s*[:：]?)?\s*(\d+)(?:\.0+)?(?:号)?", text)
    return str(int(match.group(1))) if match else text


def normalize_oil(value) -> str:
    text = str(value if value is not None else "").strip()
    match = re.search(r"(?<!\d)(-?\d{1,3})\s*[#＃号]", text)
    if match:
        return str(int(match.group(1)))
    if re.fullmatch(r"-?\d{1,3}(?:\.0+)?", text):
        return str(int(float(text)))
    return text


def normalize_station(value) -> str:
    text = re.sub(r"\s+", "", str(value or ""))
    # The supplied invoice remarks add the GAWA brand before the station name.
    return re.sub(r"^(?:中南加华|加华)", "", text)


def _identifier(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = Decimal(str(value))
        if number.is_finite():
            return format(number, "f").removesuffix(".0")
    return str(value).strip()


def _header_key(value) -> str:
    return re.sub(r"[\s（）()\-—_]", "", str(value or ""))


def _find_header(ws, aliases, required):
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 40)):
        keys = {_header_key(cell.value): cell.column for cell in row if cell.value is not None}
        mapping = {}
        for name, choices in aliases.items():
            for choice in choices:
                if _header_key(choice) in keys:
                    mapping[name] = keys[_header_key(choice)]
                    break
        if set(required) <= mapping.keys():
            return row[0].row, mapping
    return None


def _blank(value):
    return value is None or str(value).strip() == ""


def _summary_row(values):
    # The supplied exports put totals / printing metadata below their records.
    text = [str(v).strip() for v in values if v is not None]
    return any(re.match(r"^(?:合计|总计|小计|打印时间\s*[:：]|制表人\s*[:：])", v) for v in text[:4])


def _read_records(ws, values_ws, header, mapping, kind, period):
    records = []
    for row in range(header + 1, ws.max_row + 1):
        vals = [values_ws.cell(row, c).value for c in range(1, ws.max_column + 1)]
        raw = [ws.cell(row, c).value for c in range(1, ws.max_column + 1)]
        if not any(not _blank(v) for v in raw) or _summary_row(raw):
            continue
        get = lambda name: vals[mapping[name] - 1] if name in mapping else None
        # A formula-only footer with no identifying columns is a total, not a transaction.
        identity = ("time", "gun", "oil") if kind == "fuel" else ("time", "electronic", "number", "remark", "oil")
        if not any(not _blank(get(k)) for k in identity) and any(isinstance(v, str) and v.startswith("=SUM(") for v in raw):
            continue
        issues, notes = [], []
        trade = discount = end_time = invoice_time = None
        card = ""
        station = str(get("station") or "").strip()
        def number(name, label, blank_zero=False):
            try:
                # data_only may have no cached value for an input formula.
                original_value = raw[mapping[name] - 1] if name in mapping else None
                if isinstance(original_value, str) and original_value.startswith("=") and _blank(get(name)):
                    raise ReconciliationError(f"{label}公式缺少计算缓存，请先在Excel中重算保存")
                return decimal_value(get(name), label, blank_zero=blank_zero)
            except ReconciliationError as exc:
                issues.append(str(exc))
                return None
        volume = number("volume", "升数")
        if kind == "fuel":
            trade = number("trade", "交易金额")
            discount = number("discount", "优惠金额", blank_zero=True)
            if _blank(get("discount")) and _blank(raw[mapping["discount"] - 1]):
                notes.append("空优惠按0计算")
            amount = trade - discount if trade is not None and discount is not None else None
            if trade is not None and discount is not None and discount > trade:
                notes.append("原始优惠金额大于交易金额，按原值扣减")
            record_time = parse_time(get("time"))
            gun = normalize_gun(get("gun"))
            time_source, invoice_number = "加油开始时间", ""
            card = _identifier(get("card"))
            if not card:
                notes.append("卡号缺失")
        else:
            amount = number("amount", "发票合计金额（含税）")
            remark = str(get("remark") or "")
            invoice_time = parse_time(get("time"))
            if invoice_time is None:
                issues.append("开票时间缺失或格式无效")
            station_label = re.search(r"(?:加油站点|油站名称|加油站名称|站点)\s*[:：]\s*(.*?)(?=交易时间\s*[:：]|(?:油枪号|枪号)\s*[:：]|$)", remark)
            if station_label and station_label.group(1).strip(" \t;；,，"):
                station = station_label.group(1).strip(" \t;；,，")
            time_label = re.search(r"交易时间\s*[:：]\s*(.*?)(?=(?:油枪号|枪号|站点|加油站点)\s*[:：]|$)", remark)
            remark_time = time_label.group(1).strip(" \t;；,，") if time_label else ""
            if remark_time:
                record_time = parse_time(remark_time)
                end_time = record_time
                time_source = "发票备注交易时间"
            else:
                record_time = invoice_time
                time_source = "开票时间（备注缺失）"
                notes.append("备注无交易时间，使用开票时间")
            gun_match = re.search(r"(?:油枪号|枪号)\s*[:：]\s*(\d+)", remark)
            gun = normalize_gun(gun_match.group(1)) if gun_match else ""
            id_value = get("electronic") if not _blank(get("electronic")) else get("number")
            invoice_number = _identifier(id_value)
            if not invoice_number:
                issues.append("发票号码缺失")
            elif isinstance(id_value, (int, float)) and len(invoice_number.split(".")[0]) > 15:
                issues.append("长发票号码以数值存储，需确认原始完整号码")
            flag, inv_status = str(get("flag") or ""), str(get("invoice_status") or "")
            if "红" in flag or any(s in inv_status for s in ["作废", "失败", "冲红", "红冲"]):
                issues.append("红票或非有效开票状态，需人工确认")
        if record_time is None:
            issues.append(f"{time_source}缺失或格式无效")
        if not gun:
            notes.append("枪号缺失")
        oil = normalize_oil(get("oil"))
        if not oil:
            notes.append("油品缺失")
        in_scope = period.contains(record_time)
        records.append(Record(ws.title, row, kind, record_time, gun, oil, volume, amount,
                              invoice_number, time_source, issues, notes, in_scope,
                              station, card, trade, discount, end_time, invoice_time))
    return records


def _time_seconds(a, b):
    delta = b.time - a.time
    return Decimal(delta.days * 86400 + delta.seconds) + Decimal(delta.microseconds) / Decimal(1000000)


def _compatible(a, b):
    return ((not a.gun or not b.gun or a.gun == b.gun)
            and (not a.oil or not b.oil or a.oil == b.oil)
            and (not a.station or not b.station or normalize_station(a.station) == normalize_station(b.station)))


def _edge_type(a, b, amount_limit, seconds_limit):
    if not _compatible(a, b):
        return None
    amount_ok = abs(b.amount - a.amount) <= amount_limit
    volume_ok = b.volume == a.volume
    time_near = abs(_time_seconds(a, b)) <= seconds_limit
    nearby_day = abs((a.time.date() - b.time.date()).days) <= 1
    strong_identity = bool(a.gun and b.gun and a.oil and b.oil)
    # Amount is the result of the audit, not a reason to reorder otherwise
    # identical transactions. Reserve same-volume chronological counterparts
    # even when their amounts disagree.
    # Include over-limit counterparts in the SAME chronological alignment.
    # Splitting them into a later stage can consume a later fueling's earlier
    # invoice first and reverse two otherwise identical transactions.
    if volume_ok and nearby_day and (strong_identity or time_near):
        return 0
    # Do not infer a counterpart across distant days from a repeated amount alone.
    if not nearby_day:
        return None
    if not time_near and volume_ok and (amount_ok or (a.gun and b.gun and a.oil and b.oil)):
        return 1
    if time_near and (amount_ok or volume_ok):
        return 1
    if time_near and a.gun and b.gun and a.oil and b.oil:
        return 2
    return None


def _assign_ordered(fuel, invoices, edges):
    """Maximize pairs per connected candidate group, keeping chronological order.

    Each component has its own sequence alignment, so unrelated guns need not
    have the same global ordering. Equal timestamps are broken by source row.
    """
    if not edges:
        return []
    adjacency = defaultdict(set)
    for i, j in edges:
        adjacency[(0, i)].add((1, j))
        adjacency[(1, j)].add((0, i))
    seen, pairs = set(), []
    for seed in sorted(adjacency):
        if seed in seen:
            continue
        stack, left, right = [seed], [], []
        seen.add(seed)
        while stack:
            node = stack.pop()
            (left if node[0] == 0 else right).append(node[1])
            for child in adjacency[node]:
                if child not in seen:
                    seen.add(child)
                    stack.append(child)
        left.sort(key=lambda i: (fuel[i].time, fuel[i].row))
        right.sort(key=lambda j: (invoices[j].time, invoices[j].row))
        n, m = len(left), len(right)
        directions = [bytearray(m + 1) for _ in range(n + 1)]
        previous = [(0, Decimal(0))] * (m + 1)
        for x in range(1, n + 1):
            current = [(0, Decimal(0))] * (m + 1)
            for y in range(1, m + 1):
                best, direction = previous[y], 1
                if current[y - 1] > best:
                    best, direction = current[y - 1], 2
                edge = edges.get((left[x - 1], right[y - 1]))
                if edge is not None:
                    prior = previous[y - 1]
                    candidate = (prior[0] + 1, prior[1] - edge)
                    if candidate >= best:
                        best, direction = candidate, 3
                current[y], directions[x][y] = best, direction
            previous = current
        x, y = n, m
        while x and y:
            direction = directions[x][y]
            if direction == 3:
                pairs.append((left[x - 1], right[y - 1]))
                x, y = x - 1, y - 1
            elif direction == 1:
                x -= 1
            else:
                y -= 1
    return pairs


def _plain(value):
    return format(value, "f") if value is not None else None


def _format_time(value):
    return value.isoformat(sep=" ", timespec="microseconds" if value.microsecond else "seconds") if value else None


def _detail(a, b):
    return "\n".join([
        f"油站：{a.station or b.station or '未提供'}",
        f"枪号：{a.gun}；卡号：{a.card or '未提供'}；油品：{a.oil}",
        f"发票号码：{b.number}",
        f"加油开始时间：{_format_time(a.time)}",
        f"加油结束时间：{_format_time(b.end_time) or '未提供（使用开票时间辅助判断）'}",
        f"开票时间：{_format_time(b.invoice_time) or '未提供'}",
        f"加油升数：{_plain(a.volume)}；发票升数：{_plain(b.volume)}",
        f"交易金额：{_plain(a.trade)}；优惠金额：{_plain(a.discount)}",
        f"应开票金额：{_plain(a.amount)}；实际发票金额：{_plain(b.amount)}",
        f"加油位置：'{a.source}'!A{a.row}（原始第{a.row}行）",
        f"发票位置：'{b.source}'!A{b.row}（原始第{b.row}行）",
    ])


def _pair_result(a, b, amount_limit, seconds_limit, duplicate=False):
    amount_diff, volume_diff = b.amount - a.amount, b.volume - a.volume
    seconds = _time_seconds(a, b)
    failures = []
    if abs(amount_diff) > amount_limit:
        failures.append("金额不符")
    if volume_diff != 0:
        failures.append("升数不符")
    if seconds > seconds_limit:
        failures.append("时间超限")
    if seconds < 0:
        failures.append("时间先后异常：结束早于开始")
    incomplete = not (a.gun and b.gun and a.oil and b.oil)
    if amount_diff != 0:
        status = "金额不符"
    elif seconds < 0 or len(failures) > 1 or incomplete:
        status = "待确认"
    elif failures:
        status = failures[0]
    else:
        status = "已匹配"
    basis = [f"枪号{a.gun or b.gun or '缺失'}；油品{a.oil or b.oil or '缺失'}；油站{a.station or b.station or '未提供'}",
             f"应开=交易金额-优惠金额；金额必须完全一致；结束-开始在0—{seconds_limit / 60}分钟内",
             f"时间使用{b.time_source}" + ("（结束时间缺失，开票时间仅作辅助）" if b.end_time is None else "（加油结束时间）")]
    if duplicate:
        basis.append("重复候选按时间顺序推定，按交易时间顺序一一配对（同秒按原行号），每笔只用一次")
    if failures:
        basis.append("不符项：" + "、".join(failures))
        basis.append("保留疑似对应关系供复核，未计为已匹配")
    if incomplete:
        basis.append("枪号或油品资料不完整，疑似对应需确认")
    basis.extend(a.notes + b.notes)
    common = dict(expected=_plain(a.amount), actual=_plain(b.amount), amount_diff=_plain(amount_diff),
                  volume_diff=_plain(volume_diff), time_diff=_plain((seconds / 60).quantize(Decimal("0.000001"))),
                  status=status, basis="；".join(basis), time_source=b.time_source,
                  fuel_completion_time=_format_time(b.end_time),
                  fuel_start_time=_format_time(a.time), fuel_end_time=_format_time(b.end_time),
                  invoice_time=_format_time(b.invoice_time), match_details=_detail(a, b) if status == "已匹配" else None,
                  gun=a.gun, card=a.card, oil=a.oil, station=a.station or b.station,
                  fuel_volume=_plain(a.volume), invoice_volume=_plain(b.volume),
                  trade_amount=_plain(a.trade), discount_amount=_plain(a.discount),
                  red=status != "已匹配" or amount_diff != 0)
    return [dict(common, source=a.source, row=a.row, counterpart=f"'{b.source}'!A{b.row}", invoice_number=b.number),
            dict(common, source=b.source, row=b.row, counterpart=f"'{a.source}'!A{a.row}", invoice_number=b.number)]


def _unpaired_result(record):
    if not record.in_scope:
        status, basis = "待确认", "未参与本次核对：记录时间不在所选核对时间段内"
    elif record.issues:
        status, basis = "待确认", "；".join(record.issues)
    else:
        status, basis = "无对应记录", "没有满足一对一对应条件的可用记录（已分配的记录不重复使用）"
    return dict(source=record.source, row=record.row, counterpart="", invoice_number=record.number,
                expected=_plain(record.amount) if record.kind == "fuel" else None,
                actual=_plain(record.amount) if record.kind == "invoice" else None,
                amount_diff=None, volume_diff=None, time_diff=None, status=status,
                basis=basis + ("；" + "；".join(record.notes) if record.notes else ""),
                fuel_completion_time=_format_time(record.end_time),
                fuel_start_time=_format_time(record.time) if record.kind == "fuel" else None,
                fuel_end_time=_format_time(record.end_time), invoice_time=_format_time(record.invoice_time),
                match_details=None,
                gun=record.gun, card=record.card, oil=record.oil, station=record.station,
                fuel_volume=_plain(record.volume) if record.kind == "fuel" else None,
                invoice_volume=_plain(record.volume) if record.kind == "invoice" else None,
                trade_amount=_plain(record.trade), discount_amount=_plain(record.discount),
                time_source=record.time_source, red=True)


def _match(fuel, invoices, amount_limit, seconds_limit):
    active_fuel = {i for i, r in enumerate(fuel) if r.in_scope and not r.issues}
    active_inv = {j for j, r in enumerate(invoices) if r.in_scope and not r.issues}
    used_fuel, used_inv, results = set(), set(), []
    duplicate_keys = Counter((normalize_station(r.station), r.gun, r.oil, r.volume) for r in fuel if not r.issues)
    candidate_counts = Counter()
    # Strong identity/quantity/time counterparts precede weaker diagnostics.
    for stage in (0, 1, 2):
        edges = {}
        available_inv = sorted(active_inv - used_inv, key=lambda j: (invoices[j].time, invoices[j].row))
        times = [invoices[j].time for j in available_inv]
        for i in sorted(active_fuel - used_fuel):
            a = fuel[i]
            window = seconds_limit if stage == 2 else max(seconds_limit, Decimal(172800))
            span = timedelta(seconds=float(window))
            for j in available_inv[bisect_left(times, a.time - span):bisect_right(times, a.time + span)]:
                b = invoices[j]
                if _edge_type(a, b, amount_limit, seconds_limit) != stage:
                    continue
                edges[(i, j)] = abs(_time_seconds(a, b))
                candidate_counts[i] += 1
        for i, j in _assign_ordered(fuel, invoices, edges):
            used_fuel.add(i)
            used_inv.add(j)
            duplicate = candidate_counts[i] > 1 or duplicate_keys[(normalize_station(fuel[i].station), fuel[i].gun, fuel[i].oil, fuel[i].volume)] > 1
            results.extend(_pair_result(fuel[i], invoices[j], amount_limit, seconds_limit, duplicate))
    results.extend(_unpaired_result(r) for i, r in enumerate(fuel) if i not in used_fuel)
    results.extend(_unpaired_result(r) for j, r in enumerate(invoices) if j not in used_inv)
    return results, len(used_fuel)


def _export(wb, configurations, results, path, red_only, invoice_source):
    by_source = defaultdict(list)
    for result in results:
        by_source[result["source"]].append(result)
    red_fill = PatternFill("solid", fgColor=RED)
    for ws, header, original_columns, records in configurations:
        if ws.title == invoice_source:
            continue
        for result in by_source[ws.title]:
            row = result["row"]
            if result["red"]:
                for cells in ws.iter_rows(min_row=row, max_row=row, min_col=1, max_col=original_columns):
                    for cell in cells:
                        cell.fill = copy(red_fill)
            if red_only:
                ws.row_dimensions[row].hidden = not result["red"]
        last_record = max((r.row for r in records), default=header)
        ws.auto_filter.ref = f"A{header}:{get_column_letter(original_columns)}{last_record}"
        ws.auto_filter.filterColumn = []
        if red_only:
            dxf_id = wb._differential_styles.add(DifferentialStyle(fill=red_fill))
            ws.auto_filter.filterColumn.append(FilterColumn(colId=0, colorFilter=ColorFilter(dxfId=dxf_id, cellColor=True)))
    wb.save(path)


def _preserve_cell_values(original, output):
    """Retain original string empties and formula caches lost by a library save.

    Shared strings are expanded to inline strings, keeping rich text. Only cell
    value payloads are restored; the requested new row fills remain in place.
    """
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    tag = lambda name: "{" + ns + "}" + name
    ET.register_namespace("", ns)
    with ZipFile(original) as source, ZipFile(output) as target:
        shared = []
        if "xl/sharedStrings.xml" in source.namelist():
            shared = list(ET.fromstring(source.read("xl/sharedStrings.xml")))
        replacements = {}
        for name in source.namelist():
            if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name) or name not in target.namelist():
                continue
            source_root, target_root = ET.fromstring(source.read(name)), ET.fromstring(target.read(name))
            cells = {c.attrib["r"]: c for c in target_root.iter(tag("c"))}
            for original_cell in source_root.iter(tag("c")):
                cell = cells.get(original_cell.attrib["r"])
                if cell is None:
                    continue
                for child in list(cell):
                    if child.tag in {tag("v"), tag("f"), tag("is")}:
                        cell.remove(child)
                original_type = original_cell.attrib.get("t")
                if original_type == "s":
                    index = original_cell.find(tag("v"))
                    inline = copy(shared[int(index.text)])
                    inline.tag = tag("is")
                    cell.set("t", "inlineStr")
                    cell.append(inline)
                else:
                    if original_type:
                        cell.set("t", original_type)
                    else:
                        cell.attrib.pop("t", None)
                    for child in original_cell:
                        if child.tag in {tag("v"), tag("f"), tag("is")}:
                            cell.append(copy(child))
            replacements[name] = ET.tostring(target_root, encoding="utf-8", xml_declaration=True)
        with tempfile.NamedTemporaryFile(dir=Path(output).parent, suffix=".xlsx", delete=False) as handle:
            staged = Path(handle.name)
        try:
            with ZipFile(staged, "w", ZIP_DEFLATED) as archive:
                for item in target.infolist():
                    archive.writestr(item, replacements.get(item.filename, target.read(item.filename)))
        except BaseException:
            staged.unlink(missing_ok=True)
            raise
    try:
        _replace_file(staged, output)
    finally:
        staged.unlink(missing_ok=True)


def _verify_saved(path, configurations, results, invoice_source):
    check = load_workbook(path, data_only=False)
    try:
        if check.sheetnames != configurations[0][0].parent.sheetnames:
            raise RuntimeError("输出工作表结构发生变化")
        config_by_name = {ws.title: (header, original_columns, records) for ws, header, original_columns, records in configurations}
        for result in results:
            header, original_columns, _ = config_by_name[result["source"]]
            ws = check[result["source"]]
            if result["red"] and any(ws.cell(result["row"], c).fill.fgColor.rgb != RED for c in range(1, original_columns + 1)):
                raise RuntimeError("输出整行红底校验失败")
        for source_ws, header, original_columns, records in configurations:
            target_ws = check[source_ws.title]
            if target_ws.max_column != original_columns:
                raise RuntimeError("原表列数发生变化")
            if source_ws.title != invoice_source and not target_ws.auto_filter.ref:
                raise RuntimeError("输出筛选范围缺失")
            if source_ws.title == invoice_source:
                if (source_ws.auto_filter != target_ws.auto_filter
                        or source_ws.freeze_panes != target_ws.freeze_panes
                        or source_ws.sheet_view.zoomScale != target_ws.sheet_view.zoomScale):
                    raise RuntimeError("发票原有筛选或视图发生变化")
            for row in source_ws.iter_rows(max_col=original_columns):
                for cell in row:
                    saved = target_ws.cell(cell.row, cell.column)
                    if saved.value != cell.value:
                        raise RuntimeError(f"原始字段未完整保留：{source_ws.title}!{cell.coordinate}")
                    original_note = (cell.comment.text, cell.comment.author) if cell.comment else None
                    saved_note = (saved.comment.text, saved.comment.author) if saved.comment else None
                    if original_note != saved_note:
                        raise RuntimeError("原表批注发生变化")
                    if source_ws.title == invoice_source and copy(saved.fill) != copy(cell.fill):
                        raise RuntimeError("发票原有标色发生变化")
    finally:
        check.close()


def _reserve_output_pair(directory, base_name):
    suffix = 1
    while True:
        stem = base_name + ("" if suffix == 1 else "_" + str(suffix))
        workbook_path, log_path = directory / (stem + ".xlsx"), directory / (stem + ".log")
        try:
            with workbook_path.open("xb"):
                pass
        except FileExistsError:
            suffix += 1
            continue
        try:
            with log_path.open("xb"):
                pass
        except FileExistsError:
            workbook_path.unlink()
            suffix += 1
            continue
        except BaseException:
            workbook_path.unlink(missing_ok=True)
            raise
        return workbook_path, log_path


def _write_log(path, input_path, report):
    labels = [("counterpart", "对应记录位置"), ("invoice_number", "完整发票号码"),
              ("station", "油站"), ("gun", "枪号"), ("card", "卡号"), ("oil", "油品"),
              ("fuel_start_time", "加油开始时间"), ("fuel_end_time", "加油结束时间"),
              ("invoice_time", "开票时间"), ("fuel_completion_time", "加油完成时间"),
              ("time_source", "时间来源"), ("fuel_volume", "加油升数"), ("invoice_volume", "发票升数"),
              ("trade_amount", "交易金额"), ("discount_amount", "优惠金额"),
              ("expected", "应开票金额"), ("actual", "实际发票金额"),
              ("amount_diff", "金额差（实际-应开）"), ("volume_diff", "升数差（发票-加油）"),
              ("time_diff", "时间差（结束-开始，分钟）"), ("basis", "配对依据")]
    with Path(path).open("w", encoding="utf-8", newline="\n") as stream:
        stream.write("加华加油与发票核对日志\n")
        stream.write(f"源文件：{input_path}\n核对时间段：{report['date']}\n")
        stream.write(f"加油记录：{report['total_fuel']}；发票记录：{report['total_invoice']}；对应：{report['paired']}对；已匹配：{report['matched']}对\n")
        stream.write(f"仅加油明细新增浅红标记：{report['red_fuel']}行；发票页保持原样\n")
        stream.write("金额严格零差额，升数完全一致；结束-开始在0—30分钟内。\n")
        for result in report["results"]:
            stream.write("\n" + "=" * 72 + "\n")
            stream.write(f"[{result['source']}!A{result['row']}]\n")
            stream.write(f"来源工作表：{result['source']}\n原始行号：{result['row']}\n核对状态：{result['status']}\n")
            for key, label in labels:
                value = result.get(key)
                stream.write(f"{label}：{'' if value is None else value}\n")
            stream.write("匹配明细：\n")
            if result.get("match_details"):
                stream.write(result["match_details"] + "\n")


def run_reconciliation(input_path, output_dir, date_text="2026-10-11", amount_tolerance="0",
                       minutes_tolerance="30", red_only=False, progress: Callable[[str], None] | None = None,
                       start_text=None, end_text=None):
    notify = progress or (lambda _message: None)
    input_path, output_dir = Path(input_path).resolve(), Path(output_dir).resolve()
    if not input_path.is_file() or input_path.suffix.lower() != ".xlsx":
        raise ReconciliationError("请选择存在的 .xlsx 文件（前两页为加油及发票数据）")
    amount_limit = decimal_value(amount_tolerance, "金额容差")
    if amount_limit != 0:
        raise ReconciliationError("当前方案要求金额完全一致，不能设置非零金额容差")
    minutes_limit = decimal_value(minutes_tolerance, "时间容差")
    if amount_limit < 0 or minutes_limit < 0 or minutes_limit > 525600:
        raise ReconciliationError("容差不得为负，时间容差不得大于525600分钟")
    period = parse_period(date_text, start_text, end_text)
    notify("正在读取工作簿，识别前两页表头…")
    try:
        wb = load_workbook(input_path, data_only=False)
        values_wb = load_workbook(input_path, data_only=True)
    except Exception as exc:
        raise ReconciliationError(f"无法读取工作簿，请确认文件未损坏或加密：{exc}") from exc
    temporary_path = None
    reserved_path = None
    temporary_log_path = None
    reserved_log_path = None
    try:
        if len(wb.worksheets) < 2:
            raise ReconciliationError("工作簿至少需要两页：加油明细与发票")
        configurations, kinds, sources = [], {}, {}
        for ws in wb.worksheets[:2]:
            found_fuel = _find_header(ws, FUEL_ALIASES, ["gun", "time", "oil", "volume", "trade", "discount"])
            found_invoice = _find_header(ws, INVOICE_ALIASES, ["time", "amount", "remark", "oil", "volume"])
            if found_fuel:
                kind, (header, mapping) = "fuel", found_fuel
            elif found_invoice:
                kind, (header, mapping) = "invoice", found_invoice
            else:
                raise ReconciliationError(f"“{ws.title}”未找到所需表头。加油须含交易金额、优惠金额、加油量、交易时间、枪号、油品；发票须含合计金额（含税）、数量、备注、开票时间、商品名称。")
            if kind in kinds:
                raise ReconciliationError("前两页必须分别为加油明细和发票")
            if any(ws.cell(header, c).value == "核对状态" for c in range(1, ws.max_column + 1)):
                raise ReconciliationError("该文件已含核对结果，请选择原始工作簿重新核对")
            records = _read_records(ws, values_wb[ws.title], header, mapping, kind, period)
            kinds[kind] = records
            sources[kind] = ws.title
            configurations.append((ws, header, ws.max_column, records))
        fuel, invoices = kinds["fuel"], kinds["invoice"]
        if not fuel and not invoices:
            raise ReconciliationError("前两页均没有可核对的明细记录")
        duplicate_numbers = Counter(r.number for r in invoices if r.number)
        for r in invoices:
            if duplicate_numbers[r.number] > 1:
                r.issues.append("重复发票号码或同票多商品行，需人工汇总确认，不自动使用该发票")
        stations = {normalize_station(r.station) for r in fuel + invoices if r.station and r.in_scope}
        if len(stations) > 1:
            for r in fuel + invoices:
                if r.in_scope and not r.station:
                    r.issues.append("涉及多个油站但该记录未提供油站，需人工确认，不跨站推断")
        notify(f"已读取加油{len(fuel)}笔、发票{len(invoices)}笔，正在按时间顺序一对一核对…")
        results, paired = _match(fuel, invoices, amount_limit, minutes_limit * 60)
        for result in results:
            if result["source"] == sources["invoice"]:
                result["red"] = False
        source_order = {ws.title: i for i, ws in enumerate(wb.worksheets)}
        results.sort(key=lambda r: (source_order[r["source"]], r["row"]))
        positions = {(r["source"], r["row"]) for r in results}
        if len(positions) != len(results) or len(results) != len(fuel) + len(invoices):
            raise RuntimeError("记录结果覆盖校验失败")
        fuel_results = [r for r in results if r["source"] == sources["fuel"]]
        invoice_results = [r for r in results if r["source"] == sources["invoice"]]
        counterparts = [r["counterpart"] for r in fuel_results if r["counterpart"]]
        if len(counterparts) != len(set(counterparts)):
            raise RuntimeError("发票重复使用校验失败")
        output_dir.mkdir(parents=True, exist_ok=True)
        safe_period = re.sub(r'[<>:"/\\|?*\s]', "_", period.label)
        base_name = f"{input_path.stem}_核对结果_{safe_period}"
        reserved_path, reserved_log_path = _reserve_output_pair(output_dir, base_name)
        with tempfile.NamedTemporaryFile(dir=output_dir, prefix=".核对中_", suffix=".xlsx", delete=False) as handle:
            temporary_path = Path(handle.name)
        with tempfile.NamedTemporaryFile(dir=output_dir, prefix=".核对中_", suffix=".log", delete=False) as handle:
            temporary_log_path = Path(handle.name)
        notify("正在保留原表字段，仅给加油明细原行标红并启用底色筛选…")
        _export(wb, configurations, results, temporary_path, red_only, sources["invoice"])
        _preserve_cell_values(input_path, temporary_path)
        notify("正在回读检查原始字段、发票原格式及加油行红底…")
        _verify_saved(temporary_path, configurations, results, sources["invoice"])
        counts = dict(Counter(r["status"] for r in fuel_results))
        report = dict(output_path=str(reserved_path), log_path=str(reserved_log_path),
                      total_fuel=len(fuel), total_invoice=len(invoices), paired=paired,
                      matched=sum(r["status"].startswith("已匹配") for r in fuel_results),
                      red_fuel=sum(r["red"] for r in fuel_results),
                      red_invoice=sum(r["red"] for r in invoice_results), status_counts=counts,
                      invoice_status_counts=dict(Counter(r["status"] for r in invoice_results)), results=results,
                      expected_total=_plain(sum((r.amount for r in fuel if r.in_scope and r.amount is not None), Decimal(0))),
                      actual_total=_plain(sum((r.amount for r in invoices if r.in_scope and r.amount is not None), Decimal(0))),
                      date=period.label, period_start=_format_time(period.start), period_end=_format_time(period.end),
                      amount_tolerance=_plain(amount_limit), minutes_tolerance=_plain(minutes_limit))
        notify("正在保存完整核对日志（不写入原表）…")
        _write_log(temporary_log_path, input_path, report)
        saved_log = temporary_log_path.read_text(encoding="utf-8")
        if any(f"[{r['source']}!A{r['row']}]" not in saved_log for r in results):
            raise RuntimeError("核对日志记录覆盖校验失败")
        _replace_file(temporary_path, reserved_path)
        temporary_path = None
        _replace_file(temporary_log_path, reserved_log_path)
        temporary_log_path = reserved_path = reserved_log_path = None
        notify(f"完成：对应{paired}对，成功匹配{report['matched']}对，加油红底{report['red_fuel']}笔。")
        return report
    except PermissionError as exc:
        raise ReconciliationError("无法保存结果，请关闭已打开的输出文件，或选择可写入的输出目录") from exc
    finally:
        values_wb.close()
        wb.close()
        for path in (temporary_path, reserved_path, temporary_log_path, reserved_log_path):
            if path is not None:
                path.unlink(missing_ok=True)
