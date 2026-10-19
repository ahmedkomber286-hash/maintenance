"""Spare-parts maintenance tracker: Streamlit + Google Sheets."""
import hmac
import json
import re
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

# ====== Settings ======
TZ = ZoneInfo("Africa/Cairo")
COL_LINE = "Production Line"
COL_MACHINE = "Machine Name"
COL_NAME = "Part Name"
COL_PART_NO = "Part No."
COL_FREQ = "Frequancy Rate (month)"
COL_LAST = "Last change"
COL_NEXT = "Next change"
COL_NOTES = "Notes"
COL_STOCK = "Stock"
COL_LIMIT = "Stock limit"
REF_TAB = "Dropdowns"   # auto-created: per-part history (source of the dropdown lists)
LOG_TAB = "Log"         # auto-created: full audit log (who changed what, and when)
MAX_HIST = 100          # max history values kept per cell
# ======================

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
RED = {"red": 1, "green": 0.78, "blue": 0.81}
YELLOW = {"red": 1, "green": 0.92, "blue": 0.61}


# ---------- Helpers ----------
def today() -> date:
    return datetime.now(TZ).date()


def key(s) -> str:
    return str(s).strip().casefold()


def col_index(headers):
    """Header name (casefolded) -> 1-based column number."""
    return {key(h): i + 1 for i, h in enumerate(headers) if str(h).strip()}


def col_letter(c: int) -> str:
    return re.sub(r"\d", "", rowcol_to_a1(1, c))


def num(v):
    if isinstance(v, bool) or v in (None, ""):
        return None
    try:
        return float(str(v).replace(",", "")) if isinstance(v, str) else float(v)
    except ValueError:
        return None


def to_date(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return date(1899, 12, 30) + timedelta(days=int(v))     # Google/Excel date serial
    if isinstance(v, str):
        for f in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


def add_freq(start: date, months):
    try:
        return start + relativedelta(months=int(float(months)))
    except (TypeError, ValueError):
        return None


def grid(sid, row, col):
    return {"sheetId": sid, "startRowIndex": row - 1, "endRowIndex": row,
            "startColumnIndex": col - 1, "endColumnIndex": col}


def date_fmt_request(sid, row, col):
    return {"repeatCell": {"range": grid(sid, row, col), "cell": {"userEnteredFormat": {
        "numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}},
        "fields": "userEnteredFormat.numberFormat"}}


# ---------- Google Sheets connection ----------
@st.cache_resource
def get_book():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])


@st.cache_resource
def main_ws():
    name = st.secrets.get("sheet_name")
    return get_book().worksheet(name) if name else get_book().sheet1


@st.cache_resource
def get_tab(name):
    return get_book().worksheet(name)


@st.cache_resource
def setup_sheet():
    """One-time preparation: Stock columns, history/log tabs, and row colouring in the sheet."""
    book, ws = get_book(), main_ws()
    hdr = ws.row_values(1)
    ci = col_index(hdr)
    for name in (COL_STOCK, COL_LIMIT):
        if key(name) not in ci:
            c = len(hdr) + 1
            if c > ws.col_count:
                ws.add_cols(1)
            ws.update_cell(1, c, name)
            hdr.append(name)
            ci[key(name)] = c
    missing = [n for n in (COL_PART_NO, COL_NAME, COL_LAST, COL_NEXT, COL_NOTES) if key(n) not in ci]
    if missing:
        raise RuntimeError(f"Missing columns in the sheet: {missing}")

    titles = [w.title for w in book.worksheets()]
    if REF_TAB not in titles:
        t = book.add_worksheet(title=REF_TAB, rows=500, cols=2 + MAX_HIST)
        t.update(range_name="A1:C1", values=[["Part No.", "Field", "History (newest first)"]])
    if LOG_TAB not in titles:
        t = book.add_worksheet(title=LOG_TAB, rows=2000, cols=6)
        t.update(range_name="A1:F1", values=[["Time", "User", "Part No.", "Field", "Old", "New"]])

    # Row colouring inside the sheet (replaces any older/wrong rules of the same kind)
    meta = book.fetch_sheet_metadata({"fields": "sheets(properties.sheetId,conditionalFormats)"})
    fmts = next((s.get("conditionalFormats", []) for s in meta["sheets"]
                 if s["properties"]["sheetId"] == ws.id), [])
    cn, cs, cm = (col_letter(ci[key(x)]) for x in (COL_NEXT, COL_STOCK, COL_LIMIT))
    over_f = f"=AND(ISNUMBER(${cn}2),${cn}2<TODAY())"
    low_f = f"=AND(ISNUMBER(${cs}2),ISNUMBER(${cm}2),${cs}2<${cm}2)"
    stale = [i for i, f in enumerate(fmts)
             if "TODAY()" in json.dumps(f) or "ISNUMBER" in json.dumps(f)]
    rng = {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": ws.row_count,
           "startColumnIndex": 0, "endColumnIndex": len(hdr)}

    def rule(formula, color, idx):
        return {"addConditionalFormatRule": {"index": idx, "rule": {"ranges": [rng], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
            "format": {"backgroundColor": color}}}}}

    reqs = [{"deleteConditionalFormatRule": {"sheetId": ws.id, "index": i}} for i in sorted(stale, reverse=True)]
    reqs.append(rule(over_f, RED, 0))
    reqs.append(rule(low_f, YELLOW, len(fmts) - len(stale) + 1))
    book.batch_update({"requests": reqs})
    return True


# ---------- Reading data ----------
@st.cache_data(ttl=30, show_spinner="Loading data...")
def load_data():
    values = main_ws().get_all_values(value_render_option="UNFORMATTED_VALUE")
    headers = [str(h).strip() for h in values[0]]
    ci = col_index(headers)

    def cell(row, name):
        i = ci.get(key(name))
        return row[i - 1] if i and i - 1 < len(row) else ""

    rows = []
    for n, row in enumerate(values[1:], start=2):
        pn = str(cell(row, COL_PART_NO)).strip()
        if not pn:
            continue
        d = {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers) if h}
        d[COL_PART_NO] = pn
        d["_row"] = n
        d["_last"], d["_next"] = to_date(cell(row, COL_LAST)), to_date(cell(row, COL_NEXT))
        d["_freq"] = num(cell(row, COL_FREQ))
        d["_stock"], d["_limit"] = num(cell(row, COL_STOCK)), num(cell(row, COL_LIMIT))
        d["_overdue"] = bool(d["_next"] and d["_next"] < today())
        d["_low"] = d["_stock"] is not None and d["_limit"] is not None and d["_stock"] < d["_limit"]
        rows.append(d)
    return [h for h in headers if h], rows


def status_of(r):
    s = []
    if r["_overdue"]:
        s.append("⏰ Next change date has passed")
    if r["_low"]:
        s.append("📉 Stock below limit")
    return " + ".join(s)


# ---------- Writing data ----------
def push_history(ref, ref_vals, part, field, entry, seed):
    """Prepend a value to the (part, field) history row. Returns (row number, entries)."""
    r = next((i + 1 for i, v in enumerate(ref_vals)
              if len(v) > 1 and v[0] == part and v[1] == field), None)
    entries = []
    if r is None:
        r = len(ref_vals) + 1
        ref_vals.append([part, field])
    else:
        entries = [x for x in ref_vals[r - 1][2:] if x != ""]
    if not entries and seed:
        entries = [seed]
    entries = ([entry] + entries)[:MAX_HIST]
    ref_vals[r - 1] = [part, field] + entries
    if r > ref.row_count:
        ref.add_rows(max(50, r - ref.row_count))
    return r, entries


def save_changes(part_no, new, user):
    """new = {'last': date|None, 'next': date|None, 'note': str, 'stock': float|None}. Returns changed field names."""
    ws, ref, log, book = main_ws(), get_tab(REF_TAB), get_tab(LOG_TAB), get_book()
    ci = col_index(ws.row_values(1))
    c = {n: ci[key(n)] for n in (COL_PART_NO, COL_LAST, COL_NEXT, COL_NOTES, COL_STOCK)}
    parts = [str(p).strip() for p in ws.col_values(c[COL_PART_NO])]
    if part_no not in parts:
        raise RuntimeError("This part no longer exists in the sheet.")
    row = parts.index(part_no) + 1
    cur = ws.row_values(row, value_render_option="UNFORMATTED_VALUE")
    get = lambda n: cur[c[n] - 1] if c[n] - 1 < len(cur) else ""
    now = datetime.now(TZ)

    changes = []   # (field, column, value to write, history text, old value text, is_date)
    for field, new_d in ((COL_LAST, new.get("last")), (COL_NEXT, new.get("next"))):
        old_d = to_date(get(field))
        if new_d and new_d != old_d:
            s = f"{new_d:%Y-%m-%d}"
            changes.append((field, c[field], s, s, f"{old_d:%Y-%m-%d}" if old_d else "", True))
    if (new.get("note") or "").strip():
        e = f"{now:%Y-%m-%d} | {new['note'].strip()}"
        changes.append((COL_NOTES, c[COL_NOTES], e, e, str(get(COL_NOTES) or ""), False))
    old_stock = num(get(COL_STOCK))
    if new.get("stock") is not None and new["stock"] != old_stock:
        changes.append((COL_STOCK, c[COL_STOCK], new["stock"], f"{new['stock']:g}",
                        f"{old_stock:g}" if old_stock is not None else "", False))
    if not changes:
        return []

    ref_vals = ref.get_all_values()
    ref_ups, reqs, main_ups, log_rows = [], [], [], []
    last_col = 2 + MAX_HIST
    for field, col, write, hist, old, is_date in changes:
        r, entries = push_history(ref, ref_vals, part_no, field, hist, old)
        ref_ups.append({"range": f"A{r}", "values": [[part_no, field] + entries + [""] * (MAX_HIST - len(entries))]})
        reqs.append({"setDataValidation": {"range": grid(ws.id, row, col), "rule": {
            "condition": {"type": "ONE_OF_RANGE", "values": [
                {"userEnteredValue": f"='{REF_TAB}'!{rowcol_to_a1(r, 3)}:{rowcol_to_a1(r, last_col)}"}]},
            "showCustomUi": True, "strict": False}}})
        main_ups.append({"range": rowcol_to_a1(row, col), "values": [[write]]})
        log_rows.append([f"{now:%Y-%m-%d %H:%M:%S}", user, part_no, field, old, hist])
        if is_date:
            reqs.append(date_fmt_request(ws.id, row, col))

    ref.batch_update(ref_ups, value_input_option="RAW")
    ws.batch_update(main_ups, value_input_option="USER_ENTERED")
    book.batch_update({"requests": reqs})
    log.append_rows(log_rows, value_input_option="RAW")
    return [ch[0] for ch in changes]


def create_part(vals, user):
    """vals: {column name: value}. Appends a new row at the end of the data."""
    ws, log = main_ws(), get_tab(LOG_TAB)
    hdr = ws.row_values(1)
    ci = col_index(hdr)
    parts = [str(p).strip() for p in ws.col_values(ci[key(COL_PART_NO)])]
    pn = str(vals[COL_PART_NO]).strip()
    if pn.casefold() in [p.casefold() for p in parts]:
        raise RuntimeError(f"Part No. '{pn}' already exists.")
    row = len(parts) + 1
    if row > ws.row_count:
        ws.add_rows(50)
    out = [""] * len(hdr)
    for name, v in vals.items():
        if key(name) in ci:
            out[ci[key(name)] - 1] = f"{v:%Y-%m-%d}" if isinstance(v, date) else ("" if v is None else v)
    ws.update(range_name=f"A{row}", values=[out], value_input_option="USER_ENTERED")
    reqs = [date_fmt_request(ws.id, row, ci[key(n)]) for n in (COL_LAST, COL_NEXT)
            if isinstance(vals.get(n), date)]
    if reqs:
        get_book().batch_update({"requests": reqs})
    log.append_rows([[f"{datetime.now(TZ):%Y-%m-%d %H:%M:%S}", user, pn, "NEW PART",
                      "", vals.get(COL_NAME, "")]], value_input_option="RAW")
    return pn


# ---------- UI ----------
def login():
    if st.session_state.get("user"):
        return st.session_state["user"]
    users = {str(k): str(v) for k, v in dict(st.secrets.get("users", {})).items()}
    if not users:
        st.error("No users configured in Secrets (section [users]).")
        st.stop()
    st.title("🛠️ Spare Parts Maintenance Tracker")
    name = st.text_input("Username")
    pwd = st.text_input("Password", type="password")
    if st.button("Sign in", type="primary"):
        if name in users and hmac.compare_digest(users[name], pwd):
            st.session_state["user"] = name
            st.rerun()
        st.error("Wrong username or password")
    st.stop()


def colored_row(headers, r):
    data = {}
    for h in headers:
        v = r.get(h, "")
        if key(h) == key(COL_LAST):
            v = r["_last"]
        elif key(h) == key(COL_NEXT):
            v = r["_next"]
        data[h] = v.strftime("%Y-%m-%d") if isinstance(v, date) else ("—" if v in ("", None) else str(v))
    color = "#ffb3b3" if r["_overdue"] else "#ffe699" if r["_low"] else ""
    return pd.DataFrame([data]).style.apply(
        lambda row: [f"background-color: {color}; color: #000" if color else ""] * len(row), axis=1)


def search_tab(headers, rows, user):
    by_part = {r[COL_PART_NO]: r for r in rows}
    if "_goto" in st.session_state:
        st.session_state["part_sel"] = st.session_state.pop("_goto")
    part_no = st.selectbox("Search by Part No.", options=list(by_part), index=None, key="part_sel",
                           format_func=lambda p: f"{p}  —  {by_part[p].get(COL_NAME, '')}",
                           placeholder="Type a Part No. or part name…")
    if not part_no:
        return
    r = by_part[part_no]
    st.subheader(f"{r.get(COL_NAME, '')}  —  {part_no}")
    if r["_overdue"]:
        st.error("⏰ The Next change date has passed and the part has not been changed yet.")
    if r["_low"]:
        st.warning(f"📉 Stock ({r['_stock']:g}) is below the stock limit ({r['_limit']:g}).")
    st.dataframe(colored_row(headers, r), hide_index=True, width="stretch")

    st.divider()
    st.markdown("#### Update")
    v = st.session_state["ver"]
    c1, c2 = st.columns(2)
    new_last = c1.date_input("Last change", value=r["_last"], min_value=date(2000, 1, 1),
                             max_value=date(2100, 1, 1), key=f"last_{part_no}_{v}")
    new_next = c2.date_input("Next change", value=r["_next"], min_value=date(2000, 1, 1),
                             max_value=date(2100, 1, 1), key=f"next_{part_no}_{v}")
    if r["_freq"] and new_last and add_freq(new_last, r["_freq"]):
        st.caption(f"Frequency is {r['_freq']:g} month(s) → suggested next change: "
                   f"**{add_freq(new_last, r['_freq']):%Y-%m-%d}**")
    new_note = st.text_input("Notes (new note — leave empty to keep the current one)", key=f"note_{part_no}_{v}")
    new_stock = st.number_input("Stock", min_value=0.0, step=1.0,
                                value=r["_stock"] if r["_stock"] is not None else 0.0, key=f"stock_{part_no}_{v}")

    stock_changed = new_stock != (r["_stock"] if r["_stock"] is not None else 0.0)
    changed = (new_last != r["_last"]) or (new_next != r["_next"]) or bool(new_note.strip()) or stock_changed
    if st.button("💾 Save", type="primary", disabled=not changed):
        try:
            done = save_changes(part_no, {"last": new_last, "next": new_next, "note": new_note,
                                          "stock": new_stock if (stock_changed or r["_stock"] is not None) else None},
                                user)
        except (gspread.exceptions.APIError, RuntimeError) as e:
            st.error(f"Could not save: {e}")
            return
        load_data.clear()
        st.session_state["ver"] += 1
        st.session_state["msg"] = f"Saved ✅  {part_no}: updated {', '.join(done) or 'nothing'}"
        if r["_limit"] is not None and new_stock < r["_limit"]:
            st.session_state["warn"] = (f"📉 Alert: stock of {part_no} is now {new_stock:g}, "
                                        f"below the limit ({r['_limit']:g}).")
        st.rerun()


def new_part_tab(user):
    st.markdown("#### Add a new spare part")
    with st.form("new_part", clear_on_submit=True):
        c1, c2 = st.columns(2)
        pn = c1.text_input("Part No. *")
        name = c2.text_input("Part Name *")
        line = c1.text_input("Production Line")
        machine = c2.text_input("Machine Name")
        freq = c1.number_input("Frequency rate (months)", min_value=0.0, step=1.0, value=0.0)
        notes = c2.text_input("Notes")
        last = c1.date_input("Last change", value=None, min_value=date(2000, 1, 1), max_value=date(2100, 1, 1))
        nxt = c2.date_input("Next change", value=None, min_value=date(2000, 1, 1), max_value=date(2100, 1, 1))
        stock = c1.number_input("Stock", min_value=0.0, step=1.0, value=0.0)
        limit = c2.number_input("Stock limit", min_value=0.0, step=1.0, value=0.0)
        submitted = st.form_submit_button("➕ Create part", type="primary")
    if submitted:
        if not pn.strip() or not name.strip():
            st.error("Part No. and Part Name are required.")
            return
        vals = {COL_PART_NO: pn.strip(), COL_NAME: name.strip(), COL_LINE: line.strip(),
                COL_MACHINE: machine.strip(), COL_FREQ: freq or "", COL_LAST: last, COL_NEXT: nxt,
                COL_NOTES: f"{today():%Y-%m-%d} | {notes.strip()}" if notes.strip() else "",
                COL_STOCK: stock, COL_LIMIT: limit}
        try:
            created = create_part(vals, user)
        except (gspread.exceptions.APIError, RuntimeError) as e:
            st.error(f"Could not create the part: {e}")
            return
        load_data.clear()
        st.session_state["_goto"] = created
        st.session_state["msg"] = f"Part {created} was added to the sheet ✅ — find it in the Search tab."
        st.rerun()


def main():
    st.set_page_config(page_title="Spare Parts Tracker", page_icon="🛠️", layout="wide")
    user = login()
    try:
        setup_sheet()
        headers, rows = load_data()
    except gspread.exceptions.APIError as e:
        st.error(f"Google Sheets connection problem: {e}")
        st.stop()
    except RuntimeError as e:
        st.error(str(e))
        st.stop()

    st.title("🛠️ Spare Parts Maintenance Tracker")
    st.caption(f"Signed in as: {user}")
    st.session_state.setdefault("ver", 0)
    if "msg" in st.session_state:
        st.success(st.session_state.pop("msg"))
    if "warn" in st.session_state:
        st.warning(st.session_state.pop("warn"))

    flagged = [r for r in rows if r["_overdue"] or r["_low"]]
    if flagged:
        with st.expander(f"🔔 Alerts ({len(flagged)} parts)"):
            st.dataframe(pd.DataFrame([{"Part No.": r[COL_PART_NO], "Part Name": r.get(COL_NAME, ""),
                                        "Alert": status_of(r)} for r in flagged]),
                         hide_index=True, width="stretch", height=min(400, 40 + 35 * len(flagged)))

    tab1, tab2 = st.tabs(["🔎 Search & update", "➕ Add new part"])
    with tab1:
        search_tab(headers, rows, user)
    with tab2:
        new_part_tab(user)


if __name__ == "__main__":
    main()
