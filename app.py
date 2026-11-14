import io
import os
import sqlite3
import zipfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "data" / "hills_reports.db"
DB_PATH.parent.mkdir(exist_ok=True)

st.set_page_config(page_title="HILLS Daily Report", page_icon="📋", layout="wide")

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_date TEXT NOT NULL,
            weather TEXT,
            temperature TEXT,
            manager TEXT,
            currency TEXT DEFAULT 'JPY',
            target REAL DEFAULT 0,
            forecast REAL DEFAULT 0,
            actual REAL DEFAULT 0,
            customers INTEGER DEFAULT 0,
            notes TEXT,
            photo_name TEXT,
            photo BLOB,
            created_at TEXT
        )""")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS hourly_sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_id INTEGER,
            hour TEXT,
            sales REAL DEFAULT 0,
            customers INTEGER DEFAULT 0,
            FOREIGN KEY(report_id) REFERENCES reports(id)
        )""")
        conn.execute("""
        CREATE TABLE IF NOT EXISTS staff (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_id INTEGER,
            name TEXT,
            role TEXT,
            start_time TEXT,
            end_time TEXT,
            break_time TEXT,
            FOREIGN KEY(report_id) REFERENCES reports(id)
        )""")

def money(v):
    try:
        return f"¥{float(v):,.0f}"
    except Exception:
        return str(v)

init_db()

st.title("HILLS Daily Report")
st.caption("บันทึกรายงานประจำวัน • ถ่ายรูปแนบต้นฉบับ • ส่งออก Excel")

tab_new, tab_history = st.tabs(["＋ บันทึกรายงาน", "ประวัติ / ส่งออก"])

with tab_new:
    st.subheader("1. รูปภาพรายงานต้นฉบับ")
    photo = st.camera_input("ถ่ายรูปรายงาน") 
    uploaded = st.file_uploader("หรือเลือกภาพจากเครื่อง", type=["jpg", "jpeg", "png"])
    image_file = uploaded if uploaded is not None else photo
    if image_file:
        st.image(image_file, caption="รูปที่แนบ", use_container_width=True)

    st.divider()
    st.subheader("2. ข้อมูลรายงาน")
    with st.form("report_form", clear_on_submit=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            report_date = st.date_input("วันที่", value=date.today())
            weather = st.selectbox("สภาพอากาศ", ["晴れ (แดดออก)", "曇り (มีเมฆ)", "雨 (ฝน)", "雪 (หิมะ)", "その他 (อื่น ๆ)"])
        with c2:
            temperature = st.text_input("อุณหภูมิ", placeholder="เช่น 18°C")
            manager = st.text_input("ผู้รับผิดชอบ / Manager")
        with c3:
            currency = st.selectbox("สกุลเงิน", ["JPY", "THB", "USD"], index=0)
            customers = st.number_input("จำนวนลูกค้า", min_value=0, step=1)

        st.markdown("**ยอดขาย**")
        a, b, c = st.columns(3)
        with a:
            target = st.number_input("เป้าหมาย (Target)", min_value=0.0, step=1000.0)
        with b:
            forecast = st.number_input("ยอดคาดการณ์ (Forecast)", min_value=0.0, step=1000.0)
        with c:
            actual = st.number_input("ยอดขายจริง (Actual)", min_value=0.0, step=1000.0)
        notes = st.text_area("หมายเหตุ / Notes", height=90)

        st.divider()
        st.subheader("3. ยอดขายรายชั่วโมง")
        default_hours = pd.DataFrame({
            "เวลา": ["10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00", "20:00"],
            "ยอดขาย": [0]*11,
            "จำนวนลูกค้า": [0]*11
        })
        hourly = st.data_editor(default_hours, num_rows="dynamic", use_container_width=True, key="hourly_editor")

        st.divider()
        st.subheader("4. ตารางพนักงาน")
        default_staff = pd.DataFrame({
            "ชื่อพนักงาน": [""],
            "ตำแหน่ง": [""],
            "เวลาเข้า": [""],
            "เวลาออก": [""],
            "เวลาพัก": [""]
        })
        staff_df = st.data_editor(default_staff, num_rows="dynamic", use_container_width=True, key="staff_editor")

        submitted = st.form_submit_button("บันทึกรายงาน", type="primary", use_container_width=True)

    if submitted:
        photo_bytes = image_file.getvalue() if image_file else None
        photo_name = f"{report_date.isoformat()}_report.{image_file.type.split('/')[-1]}" if image_file else ""
        with db() as conn:
            cur = conn.execute("""
                INSERT INTO reports
                (report_date, weather, temperature, manager, currency, target, forecast, actual,
                 customers, notes, photo_name, photo, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (report_date.isoformat(), weather, temperature, manager, currency, target,
                  forecast, actual, int(customers), notes, photo_name, photo_bytes,
                  datetime.now().isoformat(timespec="seconds")))
            rid = cur.lastrowid
            for _, row in hourly.iterrows():
                hour = str(row.get("เวลา", "")).strip()
                if hour and hour.lower() != "nan":
                    conn.execute("INSERT INTO hourly_sales(report_id,hour,sales,customers) VALUES(?,?,?,?)",
                                 (rid, hour, float(row.get("ยอดขาย", 0) or 0), int(row.get("จำนวนลูกค้า", 0) or 0)))
            for _, row in staff_df.iterrows():
                name = str(row.get("ชื่อพนักงาน", "")).strip()
                if name and name.lower() != "nan":
                    conn.execute("INSERT INTO staff(report_id,name,role,start_time,end_time,break_time) VALUES(?,?,?,?,?,?)",
                                 (rid, name, str(row.get("ตำแหน่ง", "")), str(row.get("เวลาเข้า", "")),
                                  str(row.get("เวลาออก", "")), str(row.get("เวลาพัก", ""))))
        st.success(f"บันทึกรายงานวันที่ {report_date.isoformat()} แล้ว")

with tab_history:
    st.subheader("ประวัติรายงาน")
    with db() as conn:
        reports = pd.read_sql_query("""
            SELECT id, report_date AS วันที่, weather AS สภาพอากาศ, manager AS ผู้รับผิดชอบ,
                   target AS เป้าหมาย, forecast AS คาดการณ์, actual AS ยอดขายจริง,
                   customers AS จำนวนลูกค้า, photo_name AS รูปภาพ
            FROM reports ORDER BY report_date DESC, id DESC
        """, conn)
    if reports.empty:
        st.info("ยังไม่มีรายงานที่บันทึก")
    else:
        search = st.text_input("ค้นหาวันที่หรือผู้รับผิดชอบ")
        shown = reports
        if search:
            shown = reports[reports.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)]
        st.dataframe(shown, use_container_width=True, hide_index=True)
        st.caption(f"ทั้งหมด {len(reports)} รายงาน")
        st.divider()
        st.subheader("ดาวน์โหลดข้อมูล")
        # Build workbook with normalized sheets.
        with db() as conn:
            summary = pd.read_sql_query("""
                SELECT id AS Report_ID, report_date AS Date, weather AS Weather, temperature AS Temperature,
                manager AS Manager, currency AS Currency, target AS Target, forecast AS Forecast,
                actual AS Actual, customers AS Customers, notes AS Notes, photo_name AS Photo_File
                FROM reports ORDER BY report_date, id
            """, conn)
            hourly_all = pd.read_sql_query("""
                SELECT h.report_id AS Report_ID, r.report_date AS Date, h.hour AS Hour,
                h.sales AS Sales, h.customers AS Customers
                FROM hourly_sales h JOIN reports r ON r.id=h.report_id ORDER BY r.report_date,h.hour
            """, conn)
            staff_all = pd.read_sql_query("""
                SELECT s.report_id AS Report_ID, r.report_date AS Date, s.name AS Name, s.role AS Role,
                s.start_time AS Start, s.end_time AS End, s.break_time AS Break
                FROM staff s JOIN reports r ON r.id=s.report_id ORDER BY r.report_date,s.name
            """, conn)
            photos = conn.execute("SELECT id, report_date, photo_name, photo FROM reports WHERE photo IS NOT NULL").fetchall()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary.to_excel(writer, sheet_name="Daily Summary", index=False)
            hourly_all.to_excel(writer, sheet_name="Hourly Sales", index=False)
            staff_all.to_excel(writer, sheet_name="Staff", index=False)
            for ws in writer.book.worksheets:
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="315B4C")
                    cell.alignment = Alignment(horizontal="center")
                for col in ws.columns:
                    max_len = min(max(len(str(c.value or "")) for c in col) + 2, 38)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = max(12, max_len)
        xlsx_bytes = output.getvalue()

        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("HILLS_Report.xlsx", xlsx_bytes)
            for p in photos:
                ext = Path(p["photo_name"] or "report.jpg").suffix or ".jpg"
                zf.writestr(f"photos/{p['report_date']}_ID{p['id']}{ext}", p["photo"])
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("ดาวน์โหลด Excel", xlsx_bytes,
                               file_name="HILLS_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with c2:
            st.download_button("ดาวน์โหลด Excel + รูปต้นฉบับ (ZIP)", zbuf.getvalue(),
                               file_name="HILLS_Report_with_Photos.zip",
                               mime="application/zip", use_container_width=True)

        st.caption("Excel มี 3 ชีต: Daily Summary, Hourly Sales และ Staff ส่วนรูปต้นฉบับจะอยู่ใน ZIP แยกตามวันที่")
        selected_id = st.selectbox("ดูรายละเอียดรายงาน", reports["id"].tolist(), format_func=lambda x: f"ID {x} — {reports.loc[reports.id==x, 'วันที่'].iloc[0]}")
        with db() as conn:
            rec = conn.execute("SELECT * FROM reports WHERE id=?", (int(selected_id),)).fetchone()
            h = pd.read_sql_query("SELECT hour AS เวลา,sales AS ยอดขาย,customers AS จำนวนลูกค้า FROM hourly_sales WHERE report_id=?", conn, params=(int(selected_id),))
            s = pd.read_sql_query("SELECT name AS ชื่อพนักงาน,role AS ตำแหน่ง,start_time AS เวลาเข้า,end_time AS เวลาออก,break_time AS เวลาพัก FROM staff WHERE report_id=?", conn, params=(int(selected_id),))
        st.markdown(f"**วันที่:** {rec['report_date']}　 **ยอดขายจริง:** {money(rec['actual'])}　 **เป้าหมาย:** {money(rec['target'])}")
        if rec["photo"]:
            st.image(rec["photo"], caption=rec["photo_name"], width=420)
        st.write("ยอดขายรายชั่วโมง")
        st.dataframe(h, use_container_width=True, hide_index=True)
        st.write("พนักงาน")
        st.dataframe(s, use_container_width=True, hide_index=True)

st.divider()
st.caption("หมายเหตุ: เวอร์ชันทดลองนี้ใช้ SQLite ในเครื่อง/พื้นที่ของแอป หาก Deploy บน Streamlit Community Cloud ข้อมูลอาจหายเมื่อระบบรีสตาร์ตหรือสร้างแอปใหม่ ควรเพิ่มฐานข้อมูลถาวรก่อนใช้งานจริง")
