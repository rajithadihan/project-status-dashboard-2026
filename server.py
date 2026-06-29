
#!/usr/bin/env python3
"""
IPD Data Enrichment Excel Dashboard local server.
Reads Project-status.xlsx from this folder and returns analytics for dashboard.html.
Uses only Python standard library.
"""
from __future__ import annotations

import json
import mimetypes
import os
import re
import webbrowser
from collections import defaultdict
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from zipfile import ZipFile
from xml.etree import ElementTree as ET

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Project-status.xlsx"
PORT = int(os.environ.get("IPD_DASHBOARD_PORT", "8765"))
HOST = "127.0.0.1"
PROJECT_DEADLINE = "2026-09-30"
HOLIDAYS_2026 = {
    "2026-01-03", "2026-01-15", "2026-02-01", "2026-02-04", "2026-03-02", "2026-04-01",
    "2026-04-13", "2026-04-14", "2026-05-01", "2026-05-02", "2026-05-30", "2026-06-29",
    "2026-07-29", "2026-08-26", "2026-08-27", "2026-09-26", "2026-10-25", "2026-11-24",
    "2026-12-23", "2026-12-25"
}

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

COMPLETED_STATUSES = {"completed", "complete", "done", "finished", "closed"}
WORKED_STATUSES = {"worked", "work", "yes", "y", "1", "present", "available", "active"}


def norm_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def norm_code(value) -> str:
    text = str(value or "").replace("\u00a0", " ").strip()
    if re.match(r"^\d+\.0$", text):
        text = text[:-2]
    return text.upper()


def excel_serial_to_date(value) -> str:
    try:
        n = float(value)
    except Exception:
        return str(value or "").strip()
    dt = datetime(1899, 12, 30) + timedelta(days=n)
    return dt.date().isoformat()


def parse_date_value(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if re.match(r"^\d+(\.\d+)?$", text):
        return excel_serial_to_date(text)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%d %b %Y", "%d-%b-%Y", "%d %B %Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except Exception:
            pass
    try:
        return datetime.fromisoformat(text.replace("Z", "")).date().isoformat()
    except Exception:
        return text


def col_index_from_cell_ref(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref or "A1")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_shared_strings(z: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall(NS_MAIN + "si"):
        texts = []
        for t in si.iter(NS_MAIN + "t"):
            texts.append(t.text or "")
        out.append("".join(texts))
    return out


def workbook_sheet_paths(z: ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(NS_PKG_REL + "Relationship")}
    sheets = {}
    for sheet in wb.find(NS_MAIN + "sheets"):
        name = sheet.attrib.get("name", "")
        rid = sheet.attrib.get(NS_REL + "id")
        target = rel_map.get(rid, "")
        if target:
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheets[name] = target
    return sheets


def read_sheet(z: ZipFile, sheet_path: str, shared_strings: list[str]) -> list[dict[str, str]]:
    root = ET.fromstring(z.read(sheet_path))
    raw_rows = []
    for row in root.iter(NS_MAIN + "row"):
        vals = {}
        for c in row.findall(NS_MAIN + "c"):
            idx = col_index_from_cell_ref(c.attrib.get("r", "A1"))
            t = c.attrib.get("t")
            value = ""
            v = c.find(NS_MAIN + "v")
            if t == "s" and v is not None:
                try:
                    value = shared_strings[int(v.text)]
                except Exception:
                    value = ""
            elif t == "inlineStr":
                value = "".join(tx.text or "" for tx in c.iter(NS_MAIN + "t"))
            elif v is not None:
                value = v.text or ""
            vals[idx] = str(value).strip()
        if vals:
            max_idx = max(vals)
            raw_rows.append([vals.get(i, "") for i in range(max_idx + 1)])

    if not raw_rows:
        return []

    header = [str(x).strip() for x in raw_rows[0]]
    rows = []
    for r in raw_rows[1:]:
        if not any(str(x).strip() for x in r):
            continue
        obj = {}
        for i, h in enumerate(header):
            if h:
                obj[h] = r[i] if i < len(r) else ""
        rows.append(obj)
    return rows


def get_col(row: dict, candidates: list[str]) -> str:
    lookup = {norm_header(k): v for k, v in row.items()}
    for c in candidates:
        key = norm_header(c)
        if key in lookup:
            return lookup[key]
    return ""


def read_excel() -> tuple[list[dict], list[dict], list[dict]]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"{EXCEL_FILE.name} was not found in the dashboard folder.")
    with ZipFile(EXCEL_FILE) as z:
        ss = read_shared_strings(z)
        paths = workbook_sheet_paths(z)
        if "All" not in paths or "Task" not in paths:
            raise ValueError("Workbook must contain sheets named All and Task.")
        all_rows = read_sheet(z, paths["All"], ss)
        task_rows = read_sheet(z, paths["Task"], ss)
        working_rows = read_sheet(z, paths["Working Days"], ss) if "Working Days" in paths else []
    return all_rows, task_rows, working_rows


def add_group(group: dict, key: str, total_inc=0, completed_inc=0):
    key = str(key or "Unassigned").strip() or "Unassigned"
    row = group.setdefault(key, {"name": key, "total": 0, "completed": 0, "remaining": 0, "pct": 0})
    row["total"] += total_inc
    row["completed"] += completed_inc


def finish_group(group: dict) -> list[dict]:
    rows = []
    for row in group.values():
        row["remaining"] = max(row["total"] - row["completed"], 0)
        row["pct"] = (row["completed"] / row["total"] * 100) if row["total"] else 0
        rows.append(row)
    rows.sort(key=lambda x: (-x["remaining"], x["name"]))
    return rows


def abc_sort_key(row: dict):
    name = str(row.get("name", "") or "Unassigned").strip().upper()
    order = {"A": 0, "B": 1, "C": 2, "D": 3, "UNASSIGNED": 4, "": 4}
    return (order.get(name, 99), name)


def add_tree_count(tree: dict, category: str, stock_group: str, class_name: str, done: int):
    category = str(category or "Unassigned").strip() or "Unassigned"
    stock_group = str(stock_group or "Unassigned").strip() or "Unassigned"
    class_name = str(class_name or "Unassigned").strip() or "Unassigned"

    cat = tree.setdefault(category, {"name": category, "total": 0, "completed": 0, "remaining": 0, "pct": 0, "classes": {}})
    cat["total"] += 1
    cat["completed"] += done

    cl = cat["classes"].setdefault(class_name, {"name": class_name, "total": 0, "completed": 0, "remaining": 0, "pct": 0, "stockGroups": {}})
    cl["total"] += 1
    cl["completed"] += done

    sg = cl["stockGroups"].setdefault(stock_group, {"name": stock_group, "total": 0, "completed": 0, "remaining": 0, "pct": 0})
    sg["total"] += 1
    sg["completed"] += done


def finalize_tree(tree: dict) -> list[dict]:
    categories = []
    for cat in tree.values():
        cat["remaining"] = max(cat["total"] - cat["completed"], 0)
        cat["pct"] = (cat["completed"] / cat["total"] * 100) if cat["total"] else 0
        classes = []
        for cl in cat["classes"].values():
            cl["remaining"] = max(cl["total"] - cl["completed"], 0)
            cl["pct"] = (cl["completed"] / cl["total"] * 100) if cl["total"] else 0
            stock_groups = []
            for sg in cl["stockGroups"].values():
                sg["remaining"] = max(sg["total"] - sg["completed"], 0)
                sg["pct"] = (sg["completed"] / sg["total"] * 100) if sg["total"] else 0
                stock_groups.append(sg)
            stock_groups.sort(key=lambda x: (-x["remaining"], x["name"]))
            cl["stockGroups"] = stock_groups
            classes.append(cl)
        classes.sort(key=lambda x: (-x["remaining"], x["name"]))
        cat["classes"] = classes
        categories.append(cat)
    categories.sort(key=lambda x: (-x["remaining"], x["name"]))
    return categories


def parse_working_days(rows: list[dict]) -> dict:
    dates = []
    engineers = []
    worked = defaultdict(dict)
    raw = []
    for row in rows:
        date_val = parse_date_value(get_col(row, ["Date", "Working Date", "Day"]))
        if not date_val:
            continue
        if date_val not in dates:
            dates.append(date_val)
        row_out = {"date": date_val, "engineers": {}}
        for h, v in row.items():
            if norm_header(h) in {"date", "workingdate", "day"}:
                continue
            eng = str(h or "").strip()
            if not eng:
                continue
            if eng not in engineers:
                engineers.append(eng)
            status = str(v or "").strip()
            is_worked = norm_header(status) in WORKED_STATUSES
            worked[eng][date_val] = is_worked
            row_out["engineers"][eng] = status
        raw.append(row_out)
    dates.sort()
    engineers.sort()
    return {"dates": dates, "engineers": engineers, "worked": worked, "raw": raw}




def date_leq(a: str, b: str) -> bool:
    return bool(a and b and a <= b)


def future_business_dates(start_date: str, deadline: str) -> list[str]:
    dates = []
    if not start_date or not deadline:
        return dates
    try:
        d = datetime.fromisoformat(start_date).date() + timedelta(days=1)
        end = datetime.fromisoformat(deadline).date()
    except Exception:
        return dates
    while d <= end:
        iso_d = d.isoformat()
        if d.weekday() < 5 and iso_d not in HOLIDAYS_2026:
            dates.append(iso_d)
        d += timedelta(days=1)
    return dates


def compute_dashboard() -> dict:
    all_rows_raw, task_rows_raw, working_rows_raw = read_excel()
    generated_at = datetime.now().isoformat(timespec="seconds")
    stat = EXCEL_FILE.stat()

    working = parse_working_days(working_rows_raw)
    working_dates = working["dates"]
    working_engineers = working["engineers"]

    master = {}
    duplicate_master = 0
    for r in all_rows_raw:
        code = norm_code(get_col(r, ["Product code", "Product Code", "Product Number", "Code"]))
        if not code:
            continue
        if code in master:
            duplicate_master += 1
            continue
        master[code] = {
            "code": code,
            "abc": str(get_col(r, ["ABC Class", "ABC", "ABCClass"])).strip() or "Unassigned",
            "pm": str(get_col(r, ["Product Manager", "PM"])).strip() or "Unassigned",
            "category": str(get_col(r, ["Category"])).strip() or "Unassigned",
            "stockGroup": str(get_col(r, ["Stock Group", "StockGroup"])).strip() or "Unassigned",
            "class": str(get_col(r, ["Class"])).strip() or "Unassigned",
        }

    latest_task = {}
    task_row_count = 0
    task_duplicates = 0
    unmatched_task_codes = set()
    status_counts_all_rows = defaultdict(int)

    for idx, r in enumerate(task_rows_raw):
        code = norm_code(get_col(r, ["Product code", "Product Code", "Product Number", "Code"]))
        if not code:
            continue
        task_row_count += 1
        engineer = str(get_col(r, ["Engineer", "Owner", "Assigned Engineer"])).strip() or "Unassigned"
        status = str(get_col(r, ["Status", "Task Status"])).strip() or "Unknown"
        date_value = parse_date_value(get_col(r, ["Date", "Completed Date", "Date Completed", "Completion Date"]))
        status_key = norm_header(status)
        # A product is counted as completed when it has a completed-like status OR a completion date.
        is_completed = status_key in COMPLETED_STATUSES or bool(date_value)
        status_counts_all_rows[status] += 1
        if code not in master:
            unmatched_task_codes.add(code)
        sort_key = (date_value or "0000-00-00", idx)
        if code in latest_task:
            task_duplicates += 1
        if code not in latest_task or sort_key >= latest_task[code]["sortKey"]:
            latest_task[code] = {
                "code": code,
                "engineer": engineer,
                "status": status,
                "date": date_value,
                "isCompleted": is_completed,
                "sortKey": sort_key,
                "matched": code in master,
            }

    completed_codes = {code for code, t in latest_task.items() if t["matched"] and t["isCompleted"]}
    total = len(master)
    completed = len(completed_codes)
    remaining = max(total - completed, 0)
    completion_pct = (completed / total * 100) if total else 0

    by_abc = defaultdict(lambda: {"name": "", "total": 0, "completed": 0, "remaining": 0, "pct": 0})
    by_pm, by_category, category_tree = {}, {}, {}

    for code, item in master.items():
        done = 1 if code in completed_codes else 0
        add_group(by_pm, item["pm"], 1, done)
        add_group(by_category, item["category"], 1, done)
        add_tree_count(category_tree, item["category"], item["stockGroup"], item["class"], done)
        abc = item["abc"] or "Unassigned"
        row = by_abc[abc]
        row["name"] = abc
        row["total"] += 1
        row["completed"] += done

    by_abc_list = finish_group(by_abc)
    by_abc_list.sort(key=abc_sort_key)
    by_pm_list = finish_group(by_pm)
    by_category_list = finish_group(by_category)
    category_tree_list = finalize_tree(category_tree)

    # Completed task records and daily metrics.
    completed_tasks = [t for t in latest_task.values() if t["matched"] and t["isCompleted"]]
    completed_tasks.sort(key=lambda x: (x["date"] or "9999-12-31", x["engineer"], x["code"]))

    task_dates = sorted({t["date"] for t in completed_tasks if t["date"]})
    if not working_dates:
        # Fallback only if Working Days sheet is missing/empty.
        working_dates = task_dates
    team_working_dates = []
    for d in working_dates:
        # A team work day exists when at least one engineer is marked Worked. If there are no engineer headers,
        # keep the date as a working date.
        if not working_engineers or any(working["worked"].get(e, {}).get(d, False) for e in working_engineers):
            team_working_dates.append(d)

    # Actual productivity uses only elapsed working dates from the Working Days sheet.
    # Future working dates in the sheet are used only for forecasting to the deadline.
    latest_completed_date = task_dates[-1] if task_dates else (team_working_dates[-1] if team_working_dates else "")
    elapsed_team_working_dates = [d for d in team_working_dates if not latest_completed_date or d <= latest_completed_date]
    future_team_working_dates = [d for d in team_working_dates if latest_completed_date and latest_completed_date < d <= PROJECT_DEADLINE]
    # If future dates are not yet listed in the work calendar, use weekday/holiday-adjusted business days
    # from the latest completed work date to the deadline for forward forecasting only.
    forecast_future_dates_are_calendar_generated = False
    if latest_completed_date and not future_team_working_dates:
        future_team_working_dates = future_business_dates(latest_completed_date, PROJECT_DEADLINE)
        forecast_future_dates_are_calendar_generated = True

    all_engineers = sorted(set(working_engineers) | {t["engineer"] for t in completed_tasks if t["engineer"]})
    daily_counts_raw = defaultdict(int)
    engineer_counts = defaultdict(int)
    engineer_daily_raw = defaultdict(lambda: defaultdict(int))
    status_counts_latest = defaultdict(int)

    for t in latest_task.values():
        if t["matched"]:
            status_counts_latest[t["status"]] += 1
        if t["matched"] and t["isCompleted"]:
            d = t["date"] or "No Date"
            daily_counts_raw[d] += 1
            engineer_counts[t["engineer"]] += 1
            engineer_daily_raw[t["engineer"]][d] += 1

    # Use dates from Working Days sheet for all daily charts. Zero means worked/no completed.
    daily_series = []
    cumulative = []
    running = 0
    for d in working_dates:
        count = daily_counts_raw.get(d, 0)
        running += count
        daily_series.append({"date": d, "completed": count})
        cumulative.append({"date": d, "completed": count, "cumulative": running})

    engineer_daily_series = []
    for engineer in all_engineers:
        values = []
        for d in working_dates:
            values.append({
                "date": d,
                "completed": engineer_daily_raw[engineer].get(d, 0),
                "worked": bool(working["worked"].get(engineer, {}).get(d, False)) if working_dates else True,
            })
        engineer_daily_series.append({"engineer": engineer, "values": values})

    engineer_summary = []
    for engineer in all_engineers:
        worked_days = sum(1 for d in elapsed_team_working_dates if working["worked"].get(engineer, {}).get(d, False))
        # If Working Days did not include this engineer but task rows do, use task date count as fallback.
        if worked_days == 0 and engineer in engineer_daily_raw:
            worked_days = len(engineer_daily_raw[engineer])
        comp = engineer_counts.get(engineer, 0)
        engineer_summary.append({
            "engineer": engineer,
            "completed": comp,
            "workedDays": worked_days,
            "avgPerWorkedDay": comp / worked_days if worked_days else 0,
            "sharePct": (comp / completed * 100) if completed else 0,
        })
    engineer_summary.sort(key=lambda x: (-x["completed"], x["engineer"]))

    if elapsed_team_working_dates:
        actual_rate = completed / len(elapsed_team_working_dates)
        latest_date = elapsed_team_working_dates[-1]
        first_date = elapsed_team_working_dates[0]
        latest_day_count = daily_counts_raw.get(latest_date, 0)
    elif task_dates:
        actual_rate = completed / max(1, len(task_dates))
        latest_date = task_dates[-1]
        first_date = task_dates[0]
        latest_day_count = daily_counts_raw.get(latest_date, 0)
    else:
        actual_rate = latest_day_count = 0
        first_date = latest_date = ""

    recent_tasks = []
    for t in completed_tasks[-200:][::-1]:
        m = master.get(t["code"], {})
        recent_tasks.append({
            "code": t["code"],
            "engineer": t["engineer"],
            "status": t["status"],
            "date": t["date"],
            "abc": m.get("abc", ""),
            "pm": m.get("pm", ""),
            "category": m.get("category", ""),
            "stockGroup": m.get("stockGroup", ""),
            "class": m.get("class", ""),
        })

    # Forecast to fixed management deadline.
    remaining_workdays_to_deadline = len(future_team_working_dates)
    required_daily_rate = remaining / remaining_workdays_to_deadline if remaining_workdays_to_deadline else None
    projected_additional = actual_rate * remaining_workdays_to_deadline
    projected_completed_by_deadline = min(total, completed + projected_additional)
    projected_remaining_at_deadline = max(total - projected_completed_by_deadline, 0)
    rate_gap = (required_daily_rate - actual_rate) if required_daily_rate is not None else None

    projected_completion_date = ""
    if actual_rate > 0 and future_team_working_dates:
        need = remaining
        for d in future_team_working_dates:
            need -= actual_rate
            if need <= 0:
                projected_completion_date = d
                break
        if not projected_completion_date:
            projected_completion_date = "After listed working plan"
    elif remaining == 0:
        projected_completion_date = latest_date

    latest_active_engineers = 0
    if latest_date:
        latest_active_engineers = sum(1 for e in all_engineers if working["worked"].get(e, {}).get(latest_date, False))
    if latest_active_engineers == 0:
        latest_active_engineers = len([e for e in all_engineers if engineer_counts.get(e, 0) > 0])

    total_engineer_worked_days = sum(1 for e in all_engineers for d in elapsed_team_working_dates if working["worked"].get(e, {}).get(d, False))
    avg_per_engineer_day = completed / total_engineer_worked_days if total_engineer_worked_days else 0
    required_engineers = None
    if required_daily_rate is not None and avg_per_engineer_day > 0:
        import math
        required_engineers = math.ceil(required_daily_rate / avg_per_engineer_day)

    if required_daily_rate is None:
        forecast_status = "Need Working Date Plan"
        forecast_status_class = "watch"
        suggestion = "The forecast needs working dates before the deadline to calculate the required delivery rate."
    elif actual_rate >= required_daily_rate:
        forecast_status = "Safe Line"
        forecast_status_class = "safe"
        suggestion = "Current delivery rate is above the rate required to meet the 30 Sep 2026 deadline. Maintain current cadence and continue daily monitoring."
    elif actual_rate >= required_daily_rate * 0.9:
        forecast_status = "Watch Line"
        forecast_status_class = "watch"
        suggestion = "Current rate is close to the required line. Monitor daily output and recover any shortfall quickly."
    else:
        forecast_status = "At Risk"
        forecast_status_class = "risk"
        extra = f" Increase output by about {rate_gap:.1f} products/day." if rate_gap is not None else ""
        eng = f" Estimated required engineers: {required_engineers}." if required_engineers else ""
        suggestion = "Current delivery rate is below the rate required for the 30 Sep 2026 deadline." + extra + eng

    forecast_dates = elapsed_team_working_dates + future_team_working_dates
    run = 0
    actual_cumulative_map = {}
    for d in working_dates:
        run += daily_counts_raw.get(d, 0)
        actual_cumulative_map[d] = run
    actual_cumulative = [actual_cumulative_map.get(d, None) if d in elapsed_team_working_dates else None for d in forecast_dates]
    target_cumulative = []
    projected_cumulative = []
    for d in forecast_dates:
        if d in elapsed_team_working_dates:
            target_cumulative.append(None)
            projected_cumulative.append(None)
        else:
            idx = future_team_working_dates.index(d) + 1 if d in future_team_working_dates else 0
            target_cumulative.append(completed + (required_daily_rate or 0) * idx if required_daily_rate is not None else None)
            projected_cumulative.append(min(total, completed + actual_rate * idx))

    forecast = {
        "deadline": PROJECT_DEADLINE,
        "status": forecast_status,
        "statusClass": forecast_status_class,
        "suggestion": suggestion,
        "actualDailyRate": actual_rate,
        "requiredDailyRate": required_daily_rate,
        "rateGap": rate_gap,
        "remainingWorkdaysToDeadline": remaining_workdays_to_deadline,
        "elapsedWorkingDays": len(elapsed_team_working_dates),
        "latestDate": latest_date,
        "latestActiveEngineers": latest_active_engineers,
        "avgPerEngineerDay": avg_per_engineer_day,
        "requiredEngineers": required_engineers,
        "projectedCompletedByDeadline": projected_completed_by_deadline,
        "projectedRemainingAtDeadline": projected_remaining_at_deadline,
        "projectedCompletionDate": projected_completion_date,
        "futureDatesGenerated": forecast_future_dates_are_calendar_generated,
        "dates": forecast_dates,
        "actualCumulative": actual_cumulative,
        "targetCumulative": target_cumulative,
        "projectedCumulative": projected_cumulative,
        "requiredDailyLine": [required_daily_rate for _ in elapsed_team_working_dates] if required_daily_rate is not None else [],
        "elapsedDates": elapsed_team_working_dates,
        "elapsedDailyCompleted": [daily_counts_raw.get(d, 0) for d in elapsed_team_working_dates],
    }

    return {
        "metadata": {
            "generatedAt": generated_at,
            "excelModified": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
        },
        "overview": {
            "total": total,
            "completed": completed,
            "remaining": remaining,
            "completionPct": completion_pct,
            "actualDailyRate": actual_rate,
            "latestDayCount": latest_day_count,
            "firstWorkDate": first_date,
            "lastWorkDate": latest_date,
            "teamWorkingDays": len(team_working_dates),
            "taskRows": task_row_count,
            "uniqueTaskCodes": len(latest_task),
            "matchedTaskCodes": len([1 for t in latest_task.values() if t["matched"]]),
            "unmatchedTaskCodes": len(unmatched_task_codes),
            "duplicateTaskRows": task_duplicates,
            "duplicateMasterRows": duplicate_master,
        },
        "groups": {
            "abc": by_abc_list,
            "productManager": by_pm_list,
            "category": by_category_list,
            "categoryTree": category_tree_list,
        },
        "engineers": {
            "summary": engineer_summary,
            "dailySeries": engineer_daily_series,
        },
        "workingDays": {
            "dates": working_dates,
            "teamDates": team_working_dates,
            "elapsedTeamDates": elapsed_team_working_dates,
            "futureTeamDates": future_team_working_dates,
            "engineers": all_engineers,
        },
        "forecast": forecast,
        "trends": {
            "daily": daily_series,
            "cumulative": cumulative,
        },
        "status": {
            "latest": [{"status": k, "count": v} for k, v in sorted(status_counts_latest.items())],
            "allRows": [{"status": k, "count": v} for k, v in sorted(status_counts_all_rows.items())],
        },
        "recentTasks": recent_tasks,
    }


class DashboardHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        if self.path.startswith("/api"):
            return
        super().log_message(fmt, *args)

    def send_json(self, data, status=200):
        payload = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/api/data":
            try:
                self.send_json({"ok": True, "data": compute_dashboard()})
            except Exception as e:
                self.send_json({"ok": False, "error": str(e)}, status=500)
            return
        if path in ("/", "/dashboard.html"):
            file_path = BASE_DIR / "dashboard.html"
        else:
            file_path = BASE_DIR / path.lstrip("/")
        if not file_path.exists() or not file_path.is_file():
            self.send_error(404, "Not found")
            return
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        data = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)


def main():
    print("IPD Data Enrichment Dashboard")
    print(f"Folder: {BASE_DIR}")
    print(f"Workbook: {EXCEL_FILE.name}")
    if not EXCEL_FILE.exists():
        print("WARNING: Project-status.xlsx was not found in this folder.")
    url = f"http://{HOST}:{PORT}/"
    print(f"Opening: {url}")
    server = ThreadingHTTPServer((HOST, PORT), DashboardHandler)
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping dashboard server...")
        server.server_close()


if __name__ == "__main__":
    main()
