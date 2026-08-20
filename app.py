import io
import json
import os
import time
import threading
import urllib.request
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

app = Flask(__name__)

SHEET_ID = "1Z6C7_4niXWeDJIvhkPzhqioKXNbxppzWSBo64MsaAw4"
SHEET_XLSX_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
TARGET_FILE = "data/target.xlsx"
STAFF_STATUS_FILE = "data/staff_status.json"
OWNER_OVERRIDES_FILE = "data/owner_overrides.json"

CACHE_SECONDS = 180
_cache = {"data": None, "ts": 0}
_lock = threading.Lock()
_staff_lock = threading.Lock()
_owner_lock = threading.Lock()

# On hosts without a persistent disk (e.g. Render's free tier), small pieces
# of durable state (staff status, owner overrides) are kept in Upstash Redis
# instead of local files, so they survive restarts/redeploys/sleep cycles.
UPSTASH_URL = os.environ.get("UPSTASH_REDIS_REST_URL", "").rstrip("/")
UPSTASH_TOKEN = os.environ.get("UPSTASH_REDIS_REST_TOKEN", "")
USE_UPSTASH = bool(UPSTASH_URL and UPSTASH_TOKEN)


def _upstash_get(key):
    req = urllib.request.Request(
        f"{UPSTASH_URL}/get/{key}",
        headers={"Authorization": f"Bearer {UPSTASH_TOKEN}"},
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        result = json.loads(resp.read().decode("utf-8"))
    return result.get("result")


def _upstash_set(key, value_str):
    req = urllib.request.Request(
        f"{UPSTASH_URL}/set/{key}",
        data=value_str.encode("utf-8"),
        headers={"Authorization": f"Bearer {UPSTASH_TOKEN}", "Content-Type": "text/plain"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        resp.read()


def _kv_get(store_key, file_path, default):
    if USE_UPSTASH:
        raw = _upstash_get(store_key)
        return json.loads(raw) if raw else default
    if not os.path.exists(file_path):
        return default
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _kv_set(store_key, file_path, obj):
    payload = json.dumps(obj, ensure_ascii=False, indent=2)
    if USE_UPSTASH:
        _upstash_set(store_key, payload)
        return
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(payload)


def load_resigned_staff():
    return set(_kv_get("staff_status", STAFF_STATUS_FILE, {"resigned": []}).get("resigned", []))


def save_resigned_staff(names):
    _kv_set("staff_status", STAFF_STATUS_FILE, {"resigned": sorted(names)})


def load_owner_overrides():
    return _kv_get("owner_overrides", OWNER_OVERRIDES_FILE, {"overrides": {}}).get("overrides", {})


def save_owner_overrides(overrides):
    _kv_set("owner_overrides", OWNER_OVERRIDES_FILE, {"overrides": overrides})

# nickname (target file) -> real name (Google Sheet visit log)
SL_NICK_TO_REAL = {
    "Tew": "Phaksiri",
    "Bas": "Teerachate",
    "O": "Prasertchai",
    "Poom": "Pisud",
    "Note": "Arthit",
    "Kid": "Somkid",
    "View": "Chatchon",
    "Aim": "Amika",
    "Big": "Ritthikorn",
}
SL_REAL_TO_NICK = {v: k for k, v in SL_NICK_TO_REAL.items()}

IE_NORMALIZE = {
    "roj": "ROJ", "rojana": "ROJ",
    "bpo": "BPO",
    "all area": "ALL AREA",
    "bhs": "BHS",
}


def normalize_ie(v):
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s == "-":
        return ""
    key = s.lower()
    return IE_NORMALIZE.get(key, s.upper())


TOPIC_GROUPS = [
    ("เชิงสัมพันธ์", ["สวัสดีตามเทศกาล", "พูดคุยทั่วไป"]),
    ("สัญญา", ["สัญญาซื้อขาย", "BG", "Contract"]),
    ("ราคา/ต้นทุน", ["โครงสร้างราคา", "ส่วนต่างค่าก๊าซ"]),
    ("เทคนิค/บริการ", ["service flow", "ส่งผล PSV", "Site survey", "Survey", "Safety Training"]),
    ("ประชุม", ["Online meeting"]),
]


def group_topic(t):
    t = (t or "").strip()
    if not t:
        return "อื่นๆ"
    for group, keys in TOPIC_GROUPS:
        for k in keys:
            if k.lower() in t.lower():
                return group
    return "อื่นๆ"


def fetch_workbook_bytes():
    req = urllib.request.Request(SHEET_XLSX_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def load_visit_and_docs():
    raw = fetch_workbook_bytes()
    xls = pd.ExcelFile(io.BytesIO(raw))

    visit = pd.read_excel(xls, sheet_name="Visit Record", header=0, skiprows=[1])
    visit = visit.dropna(subset=["Company_Name"]).copy()
    visit["Company_Name"] = visit["Company_Name"].astype(str).str.strip()
    visit["CustomerID"] = visit["CustomerID"].fillna("").astype(str).str.strip()
    visit["IE"] = visit["IE"].apply(normalize_ie)
    visit["Date"] = pd.to_datetime(visit["Date"], errors="coerce")
    visit["NGD_Staff_Name"] = visit["NGD_Staff_Name"].fillna("").astype(str).str.strip()
    visit["Topic"] = visit["Topic"].fillna("").astype(str).str.strip()
    visit["TopicGroup"] = visit["Topic"].apply(group_topic)
    visit["Details"] = visit["Details"].fillna("").astype(str)
    visit["Contact_person"] = visit["Contact_person"].fillna("").astype(str)
    visit = visit.dropna(subset=["Date"])

    doc = pd.read_excel(xls, sheet_name="Document Receive", header=0, skiprows=[1])
    doc = doc.dropna(subset=["Company Name"]).copy()
    doc["Company Name"] = doc["Company Name"].astype(str).str.strip()
    doc["CustomerID"] = doc["CustomerID"].fillna("").astype(str).str.strip()
    doc["IE"] = doc["IE"].apply(normalize_ie)
    doc["Date"] = pd.to_datetime(doc["Date"], errors="coerce")
    doc["NGD Staff Name"] = doc["NGD Staff Name"].fillna("").astype(str).str.strip()
    for col in ["Document 1", "Document 2", "Note", "Reciever"]:
        doc[col] = doc[col].fillna("").astype(str).str.strip()
    doc = doc.dropna(subset=["Date"])

    return visit, doc


def load_targets():
    rows = []
    xls = pd.ExcelFile(TARGET_FILE)
    for sheet in ["PNGD", "ANGD"]:
        df = pd.read_excel(xls, sheet_name=sheet, header=0)
        df = df.dropna(subset=["Project ID."]).copy()
        for _, r in df.iterrows():
            plan = str(r.get("Plan Visit", "")).strip()
            if "4 เดือน" in plan:
                per_year, months = 3, 4
            elif "6 เดือน" in plan:
                per_year, months = 2, 6
            else:
                per_year, months = 2, 6
            nick = str(r.get("SL", "")).strip()
            rows.append({
                "CustomerID": str(r.get("Project ID.", "")).strip(),
                "CompanyName": str(r.get("Customer Name", "")).strip(),
                "IE": normalize_ie(r.get("IE", "")),
                "SL_nick": nick,
                "SL_real": SL_NICK_TO_REAL.get(nick, nick),
                "CSI2025": str(r.get("CSI 2025", "")).strip(),
                "TargetPerYear": per_year,
                "TargetMonths": months,
                "BusinessUnit": sheet,
            })
    return pd.DataFrame(rows)


def build_dataset():
    visit, doc = load_visit_and_docs()
    targets = load_targets()

    now = pd.Timestamp.now()
    year_start = pd.Timestamp(year=now.year, month=1, day=1)
    day_of_year = (now - year_start).days + 1
    year_fraction = day_of_year / 365.0

    # A "touchpoint" toward target = a visit OR a document receive, since both
    # involve in-person customer contact.
    touch = pd.concat([
        visit[["CustomerID", "Date"]],
        doc[["CustomerID", "Date"]],
    ], ignore_index=True)
    touch = touch[touch["CustomerID"] != ""]

    touch_this_year = touch[touch["Date"] >= year_start]
    touch_count_year = touch_this_year.groupby("CustomerID").size()
    touch_count_total = touch.groupby("CustomerID").size()
    last_touch = touch.groupby("CustomerID")["Date"].max()

    elapsed_months = year_fraction * 12
    owner_overrides = load_owner_overrides()

    customers = []
    for _, t in targets.iterrows():
        cid = t["CustomerID"]
        override_owner = owner_overrides.get(cid)
        sl_real = override_owner if override_owner else t["SL_real"]
        sl_nick = SL_REAL_TO_NICK.get(sl_real, "") if override_owner else t["SL_nick"]
        v_year = int(touch_count_year.get(cid, 0))
        v_total = int(touch_count_total.get(cid, 0))
        lv = last_touch.get(cid)
        target_per_year = t["TargetPerYear"]
        target_months = t["TargetMonths"]
        needed = max(0, round(target_per_year - v_year))

        days_since = None if pd.isna(lv) else (now - lv).days

        # Calendar-checkpoint based overdue logic: the year is split into
        # `target_per_year` checkpoints spaced `target_months` apart. As long
        # as the customer has met the visit count required by the most
        # recently passed checkpoint, they're not overdue - regardless of
        # exactly when within the year those visits happened.
        checkpoints_passed = int(elapsed_months // target_months)
        required_by_now = min(checkpoints_passed, target_per_year)

        if v_year >= target_per_year:
            status = "done"
        elif v_year < required_by_now:
            status = "overdue"
        else:
            status = "in_progress"

        next_checkpoint_num = v_year + 1
        next_due = (
            year_start + pd.DateOffset(months=int(target_months * next_checkpoint_num))
            if next_checkpoint_num <= target_per_year else None
        )

        customers.append({
            "customer_id": cid,
            "company_name": t["CompanyName"],
            "ie": t["IE"],
            "sl_real": sl_real,
            "sl_nick": sl_nick,
            "owner_overridden": bool(override_owner),
            "business_unit": t["BusinessUnit"],
            "csi2025": t["CSI2025"],
            "target_per_year": target_per_year,
            "target_months": target_months,
            "visits_this_year": v_year,
            "visits_total": v_total,
            "visits_needed": int(needed),
            "last_visit_date": lv.strftime("%Y-%m-%d") if pd.notna(lv) else None,
            "days_since_last_visit": int(days_since) if days_since is not None else None,
            "next_due_date": next_due.strftime("%Y-%m-%d") if next_due is not None and pd.notna(next_due) else None,
            "status": status,
        })

    target_ids = set(targets["CustomerID"])
    off_target_ids = sorted(set(visit["CustomerID"]) - target_ids - {""})
    off_target = []
    for cid in off_target_ids:
        sub = visit[visit["CustomerID"] == cid]
        off_target.append({
            "customer_id": cid,
            "company_name": sub["Company_Name"].iloc[0],
            "ie": sub["IE"].iloc[-1],
            "visits_total": int(len(sub)),
            "visits_this_year": int(len(sub[sub["Date"] >= year_start])),
            "last_visit_date": sub["Date"].max().strftime("%Y-%m-%d"),
        })

    visits_out = []
    for _, r in visit.sort_values("Date", ascending=False).iterrows():
        visits_out.append({
            "company_name": r["Company_Name"],
            "customer_id": r["CustomerID"],
            "ie": r["IE"],
            "date": r["Date"].strftime("%Y-%m-%d"),
            "month": r["Date"].strftime("%Y-%m"),
            "topic": r["Topic"],
            "topic_group": r["TopicGroup"],
            "details": r["Details"],
            "contact": r["Contact_person"],
            "staff": r["NGD_Staff_Name"],
            "sl_nick": SL_REAL_TO_NICK.get(r["NGD_Staff_Name"], r["NGD_Staff_Name"]),
        })

    docs_out = []
    for _, r in doc.sort_values("Date", ascending=False).iterrows():
        docs_out.append({
            "company_name": r["Company Name"],
            "customer_id": r["CustomerID"],
            "ie": r["IE"],
            "date": r["Date"].strftime("%Y-%m-%d"),
            "doc1": r["Document 1"],
            "doc2": r["Document 2"],
            "note": r["Note"],
            "receiver": r["Reciever"],
            "staff": r["NGD Staff Name"],
        })

    all_staff = sorted(set(t["sl_real"] for t in customers) | set(v["staff"] for v in visits_out if v["staff"]))
    resigned = load_resigned_staff()
    sl_list = [s for s in all_staff if s not in resigned]

    return {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "customers": customers,
        "off_target_customers": off_target,
        "visits": visits_out,
        "documents": docs_out,
        "sl_list": sl_list,
        "all_staff": all_staff,
        "resigned_staff": sorted(resigned),
    }


def get_dataset(force=False):
    with _lock:
        if not force and _cache["data"] is not None and (time.time() - _cache["ts"]) < CACHE_SECONDS:
            return _cache["data"]
        data = build_dataset()
        _cache["data"] = data
        _cache["ts"] = time.time()
        return data


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/data")
def api_data():
    return jsonify(get_dataset())


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    data = get_dataset(force=True)
    return jsonify({"ok": True, "generated_at": data["generated_at"]})


@app.route("/api/staff/set_status", methods=["POST"])
def api_staff_set_status():
    body = request.get_json(force=True) or {}
    name = str(body.get("name", "")).strip()
    resigned = bool(body.get("resigned", False))
    if not name:
        return jsonify({"ok": False, "error": "missing name"}), 400
    with _staff_lock:
        resigned_set = load_resigned_staff()
        if resigned:
            resigned_set.add(name)
        else:
            resigned_set.discard(name)
        save_resigned_staff(resigned_set)
    get_dataset(force=True)
    return jsonify({"ok": True, "resigned_staff": sorted(resigned_set)})


@app.route("/api/owners/export")
def api_owners_export():
    data = get_dataset()
    rows = [{
        "CustomerID": c["customer_id"],
        "บริษัท": c["company_name"],
        "พื้นที่ (IE)": c["ie"],
        "Owner ปัจจุบัน": c["sl_real"],
        "Owner ใหม่ (กรอกเฉพาะรายที่ต้องการเปลี่ยน)": "",
    } for c in data["customers"]]
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="โอนย้าย Owner")
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="owner_transfer_form.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/owners/import", methods=["POST"])
def api_owners_import():
    file = request.files.get("file")
    if not file:
        return jsonify({"ok": False, "error": "ไม่พบไฟล์ที่อัปโหลด"}), 400

    try:
        df = pd.read_excel(file, header=0)
    except Exception as e:
        return jsonify({"ok": False, "error": f"อ่านไฟล์ Excel ไม่สำเร็จ: {e}"}), 400

    cols = {str(c).strip().lower(): c for c in df.columns}
    id_col = next((cols[k] for k in cols if "customerid" in k.replace(" ", "")), None)
    new_owner_col = next((cols[k] for k in cols if "ใหม่" in k or "new" in k), None)
    if id_col is None or new_owner_col is None:
        return jsonify({"ok": False, "error": "ไม่พบคอลัมน์ CustomerID หรือ 'Owner ใหม่' ในไฟล์ กรุณาใช้แบบฟอร์มที่ดาวน์โหลดจากระบบ"}), 400

    current_data = get_dataset()
    valid_customers = {c["customer_id"]: c for c in current_data["customers"]}

    with _owner_lock:
        overrides = load_owner_overrides()
        applied = []
        skipped = []
        for _, row in df.iterrows():
            cid = str(row.get(id_col, "")).strip()
            new_owner = str(row.get(new_owner_col, "")).strip()
            if not cid or not new_owner or new_owner.lower() == "nan":
                continue
            if cid not in valid_customers:
                skipped.append({"customer_id": cid, "reason": "ไม่พบรหัสลูกค้านี้ในระบบ"})
                continue
            old_owner = valid_customers[cid]["sl_real"]
            overrides[cid] = new_owner
            applied.append({
                "customer_id": cid,
                "company_name": valid_customers[cid]["company_name"],
                "old_owner": old_owner,
                "new_owner": new_owner,
            })
        save_owner_overrides(overrides)

    get_dataset(force=True)
    return jsonify({"ok": True, "applied": applied, "skipped": skipped, "applied_count": len(applied), "skipped_count": len(skipped)})


@app.route("/api/owners/reset", methods=["POST"])
def api_owners_reset():
    body = request.get_json(force=True) or {}
    cid = str(body.get("customer_id", "")).strip()
    if not cid:
        return jsonify({"ok": False, "error": "missing customer_id"}), 400
    with _owner_lock:
        overrides = load_owner_overrides()
        overrides.pop(cid, None)
        save_owner_overrides(overrides)
    get_dataset(force=True)
    return jsonify({"ok": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=not os.environ.get("PORT"))
