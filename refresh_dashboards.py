import re, json, sys, os, openpyxl, win32com.client
from datetime import datetime

FOLDER = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.join(FOLDER, "2026_Active_WorkOrders.xlsm")
HTML_FILES = [
    os.path.join(FOLDER, "work_center_dashboard.html"),
    os.path.join(FOLDER, "measurex_dashboard.html"),
    os.path.join(FOLDER, "iml_dashboard.html"),
]

def open_close_wb():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(WORKBOOK)
    excel.CalculateUntilAsyncQueriesDone()
    wb.Close(SaveChanges=True)
    excel.Quit()
    return

def safe_num(v):
    # Return #N/A or DIV errors as None.
    if isinstance(v, (int, float)):
        return v
    return None

def load_rows():
    open_close_wb()
    wb2 = openpyxl.load_workbook(WORKBOOK, data_only=True)
    ws2 = wb2["Active_Sched"]
    tbl = ws2.tables["Active_Sched"]
    min_col, min_row, max_col, max_row = openpyxl.utils.cell.range_boundaries(tbl.ref)
    headers = [ws2.cell(row=min_row, column=c).value for c in range(min_col, max_col + 1)]
    rows = []
    for r in range(min_row + 1, max_row + 1):
        row = {headers[i]: ws2.cell(row=r, column=min_col + i).value for i in range(len(headers))}
        rows.append(row)
    return rows

def build_data():
    rows = load_rows()
    now = datetime.now()

    by_wc = {}
    for row in rows:
        wc = row.get("WorkCenter")
        if not wc:
            continue
        by_wc.setdefault(wc, []).append(row)

    work_centers = []
    for wc, items in by_wc.items():
        items.sort(key=lambda r: r["PROD_START_TIME"])

        cur_idx = None
        live_idxs = [i for i, r in enumerate(items) if safe_num(r.get("CURR_CYCLE_TIME")) is not None]
        if live_idxs:
            cur_idx = live_idxs[-1]
        else:
            window_idxs = [i for i, r in enumerate(items) if r["PROD_START_TIME"] <= now <= r["PROD_END_TIME"]]
            if window_idxs:
                cur_idx = window_idxs[0]
            else:
                past_idxs = [i for i, r in enumerate(items) if r["PROD_END_TIME"] < now]
                if past_idxs:
                    cur_idx = past_idxs[-1]

        if cur_idx is not None:
            cur = items[cur_idx]
            nxt = items[cur_idx + 1] if cur_idx + 1 < len(items) else None
            nxt2 = items[cur_idx + 2] if cur_idx + 2 < len(items) else None
        else:
            cur = None
            nxt = items[0] if items else None
            nxt2 = items[1] if len(items) > 1 else None

        def mk(r, with_cur_fields=True):
            if r is None:
                return None
            cph_raw = safe_num(r.get("Cartons Per Hour"))
            d = {
                "item": r["ITEMNO"],
                "desc": r["DESCRIP"],
                "start": r["PROD_START_TIME"].isoformat(),
                "end": r["PROD_END_TIME"].isoformat(),
                "cph": round(cph_raw, 1) if cph_raw is not None else None,
            }
            if with_cur_fields:
                d["cycle_time"] = safe_num(r.get("CYCLE_TIME"))
                d["curr_cycle_time"] = safe_num(r.get("CURR_CYCLE_TIME"))
                d["cavitation"] = safe_num(r.get("Cavitation %"))
            return d

        work_centers.append({
            "wc": wc,
            "plant": items[0]["EPLANT_ID"],
            "mfg": items[0]["MFGCELL"],
            "cur": mk(cur, True),
            "nxt": mk(nxt, False),
            "nxt2": mk(nxt2, False),
        })

    work_centers.sort(key=lambda w: w["wc"])
    return {"exported_at": now.isoformat(timespec="seconds"), "work_centers": work_centers}

def splice(html_path, data):
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    m = re.search(r"const DATA = (\{.*?\});\n\nfunction \w+\(", content, re.S)
    if not m:
        raise RuntimeError(f"DATA block not found in {html_path}")
    new_content = content[: m.start(1)] + json.dumps(data, indent=2) + content[m.end(1):]
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)

def main():
    if not os.path.exists(WORKBOOK):
        print(f"ERROR: workbook not found at {WORKBOOK}", file=sys.stderr)
        sys.exit(1)
    try:
        data = build_data()
    except Exception as e:
        print(f"ERROR reading workbook: {e}", file=sys.stderr)
        sys.exit(1)

    updated = []
    for path in HTML_FILES:
        splice(path, data)
        updated.append(os.path.basename(path))

    down = [w["wc"] for w in data["work_centers"] if not w["cur"] or not w["cur"].get("curr_cycle_time")]
    print(f"Exported at {data['exported_at']}")
    print(f"Updated: {', '.join(updated) if updated else 'NONE'}")
    print(f"Work centers with no live cycle reading (will show Down): {', '.join(down) if down else 'none'}")

if __name__ == "__main__":
    main()
