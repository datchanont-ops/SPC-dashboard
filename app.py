import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import os
import base64
import requests

# 1. ตั้งค่าหน้ากระดาษเป็น Wide Mode
st.set_page_config(page_title="SPC Production Shortage Dashboard", layout="wide")

# ==========================================
# 🔗 ระบบจำไฟล์ถาวรผ่าน GitHub (Persistence Layer)
# ==========================================
try:
    GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
    GITHUB_REPO = st.secrets["GITHUB_REPO"]
    GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")
    GITHUB_DATA_DIR = st.secrets.get("GITHUB_DATA_DIR", "data")
    GITHUB_ENABLED = True
except Exception:
    GITHUB_ENABLED = False

GITHUB_API = "https://api.github.com"

def gh_headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }

def gh_get_file(remote_path):
    if not GITHUB_ENABLED:
        return None, None
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}?ref={GITHUB_BRANCH}"
    try:
        r = requests.get(url, headers=gh_headers(), timeout=15)
        if r.status_code == 200:
            data = r.json()
            content = base64.b64decode(data["content"])
            return content, data["sha"]
    except Exception:
        pass
    return None, None

def gh_put_file(remote_path, content_bytes, message):
    if not GITHUB_ENABLED:
        return False
    _, sha = gh_get_file(remote_path)
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    try:
        r = requests.put(url, headers=gh_headers(), json=payload, timeout=15)
        return r.status_code in (200, 201)
    except Exception:
        return False

# ==========================================
# 🎨 CSS Modern Corporate Styling
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Prompt', sans-serif !important;
    }
    
    /* พื้นหลังหลักสไตล์ Light Theme */
    .stApp {
        background-color: #f8fafc !important;
    }
    
    /* ปรับแต่งกล่อง Metric / KPI Cards ให้ดูแพง */
    [data-testid="stMetric"] {
        background-color: #ffffff !important;
        padding: 20px 24px !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.04) !important;
        border: 1px solid #e2e8f0 !important;
        border-top: 4px solid #3b82f6 !important;
        position: relative !important;
        transition: transform 0.2s ease;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 15px rgba(0,0,0,0.08) !important;
    }
    
    /* สีของขอบบนการ์ดแต่ละใบ */
    [data-testid="column"]:nth-child(1) [data-testid="stMetric"] { border-top-color: #ef4444 !important; } /* แดง */
    [data-testid="column"]:nth-child(2) [data-testid="stMetric"] { border-top-color: #f59e0b !important; } /* ส้ม/เหลือง */
    [data-testid="column"]:nth-child(3) [data-testid="stMetric"] { border-top-color: #10b981 !important; } /* เขียว */

    /* ข้อความหัวข้อในการ์ด */
    [data-testid="stMetricLabel"] > div {
        color: #64748b !important;
        font-size: 15px !important;
        font-weight: 600 !important;
    }
    
    /* ตัวเลขค่าสรุปในการ์ด */
    [data-testid="stMetricValue"] > div {
        color: #0f172a !important;
        font-size: 32px !important;
        font-weight: 700 !important;
    }

    /* กล่อง Header หลัก */
    .dashboard-header {
        background-color: #ffffff;
        padding: 20px 28px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        margin-bottom: 20px;
    }
    
    /* ปรับตาราง */
    div[data-testid="stDataFrame"] {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    }
    
    .alert-box {
        background-color: #fee2e2;
        color: #b91c1c;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-weight: 500;
        border-left: 5px solid #b91c1c;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Shortage Report')
    return output.getvalue()

# 3. ฟังก์ชันประมวลผลข้อมูล
@st.cache_data
def process_data(uploaded_file, file_template, include_plan=True):
    xls = pd.ExcelFile(uploaded_file)
    
    def get_sheet(sheet_names, target):
        for s in sheet_names:
            if target.lower() in s.strip().lower(): 
                return s
        raise ValueError(f"หา Sheet ชื่อ '{target}' ไม่เจอเลยครับ! (ในไฟล์ที่อัปโหลดมีแค่ Sheet: {', '.join(sheet_names)})")

    sheet_ord = get_sheet(xls.sheet_names, 'ord bac')
    sheet_plan = get_sheet(xls.sheet_names, 'plan')
    sheet_wip = get_sheet(xls.sheet_names, 'wip fg')

    df_ord = pd.read_excel(uploaded_file, sheet_name=sheet_ord)
    df_plan = pd.read_excel(uploaded_file, sheet_name=sheet_plan)
    df_wip = pd.read_excel(uploaded_file, sheet_name=sheet_wip)
    
    df_ord.columns = df_ord.columns.astype(str).str.strip()
    df_plan.columns = df_plan.columns.astype(str).str.strip()
    df_wip.columns = df_wip.columns.astype(str).str.strip()

    def find_col(df, possible_names):
        for name in possible_names:
            for col in df.columns:
                if name.strip().lower() == col.strip().lower(): return col
        return None

    part_col_plan = find_col(df_plan, ['Part No.', 'Part no.', 'Partno']) or df_plan.columns[4]
    date_col_plan = find_col(df_plan, ['วันที่.1', 'date', 'วันที่']) or df_plan.columns[1]
    order_col_plan = find_col(df_plan, ['Order', 'Plan Order']) or df_plan.columns[5]

    mat_col_wip = find_col(df_wip, ['Material', 'Mat']) or df_wip.columns[0]
    unr_col_wip = find_col(df_wip, ['Unrestricted', 'Stock']) or df_wip.columns[5]

    sap_col_ord = find_col(df_ord, ['SAP Mat.', 'SAP Material', 'Material']) or df_ord.columns[2]
    dlv_col_ord = find_col(df_ord, ['Dlv. Date', 'Delivery Date', 'Date']) or df_ord.columns[6]
    qty_col_ord = find_col(df_ord, ['Outstd.Base Qty', 'Base Qty', 'Qty']) or df_ord.columns[5]

    df_plan[part_col_plan] = df_plan[part_col_plan].astype(str).str.strip()
    df_ord[sap_col_ord] = df_ord[sap_col_ord].astype(str).str.strip()
    
    # ทำความสะอาดชื่อ Part ใน sheet wip fg ตัด ;S1 และ ;S2 ออก
    df_wip[mat_col_wip] = df_wip[mat_col_wip].astype(str).str.strip().str.replace(';S1', '', regex=False).str.replace(';S2', '', regex=False)

    try:
        df_sum_raw = pd.read_excel(file_template, sheet_name='summary v1', header=1)
        df_sum_raw_headerless = pd.read_excel(file_template, sheet_name='summary v1', header=None)
    except Exception as e:
        raise ValueError(f"หาไฟล์/Sheet Master ไม่เจอ หรือชื่อผิด: {e}")
    
    ord_parts = df_ord[sap_col_ord].dropna().str.upper().unique()
    sum_parts = df_sum_raw['fg1'].dropna().astype(str).str.strip().str.upper().unique()
    missing_set = set(ord_parts) - set(sum_parts)
    
    df_missing = pd.DataFrame()
    if missing_set:
        df_missing_raw = df_ord[df_ord[sap_col_ord].str.upper().isin(missing_set)]
        df_missing = df_missing_raw.groupby(sap_col_ord)[qty_col_ord].sum().reset_index()
        df_missing.rename(columns={sap_col_ord: 'Missing Part (from Ord)', qty_col_ord: 'Outstanding Qty'}, inplace=True)
    
    sche_map = {}
    note_map = {}
    for idx, row in df_sum_raw.iterrows():
        part = row.get('Part No.')
        if pd.notna(part):
            p_key = str(part).strip()
            if pd.notna(row.get('SCHE')): sche_map[p_key] = str(row.get('SCHE')).strip()
            if pd.notna(row.get('Note')): note_map[p_key] = str(row.get('Note')).strip()

    df_plan['_parsed_date'] = pd.to_datetime(df_plan[date_col_plan], errors='coerce')
    df_ord['_parsed_date'] = pd.to_datetime(df_ord[dlv_col_ord], errors='coerce')

    dates_row = pd.to_datetime(df_sum_raw_headerless.iloc[1, 6:38].values)
    start_date = pd.to_datetime('2025-01-01')
    today_date = dates_row[0]

    df_plan['Norm_Date'] = df_plan['_parsed_date'].dt.normalize()
    df_ord['Norm_Date'] = df_ord['_parsed_date'].dt.normalize()

    wip_dict = df_wip.groupby(mat_col_wip)[unr_col_wip].sum().to_dict()
    
    plan_hist_mask = (df_plan['_parsed_date'] >= start_date) & (df_plan['_parsed_date'] <= today_date)
    plan_hist_dict = df_plan[plan_hist_mask].groupby(part_col_plan)[order_col_plan].sum().to_dict()
    plan_daily_dict = df_plan.groupby([part_col_plan, 'Norm_Date'])[order_col_plan].sum().to_dict()

    ord_hist_mask = (df_ord['_parsed_date'] >= start_date) & (df_ord['_parsed_date'] <= today_date)
    ord_hist_dict = df_ord[ord_hist_mask].groupby(sap_col_ord)[qty_col_ord].sum().to_dict()
    ord_daily_dict = df_ord.groupby([sap_col_ord, 'Norm_Date'])[qty_col_ord].sum().to_dict()
    
    dashboard_data = []
    for idx in range(2, len(df_sum_raw_headerless), 3):
        if idx + 2 >= len(df_sum_raw_headerless): break
        raw_part = df_sum_raw_headerless.iloc[idx, 2]
        raw_fg1 = df_sum_raw_headerless.iloc[idx, 3]
        if pd.isna(raw_part) and pd.isna(raw_fg1): continue
        
        part_no = str(raw_part).strip()
        fg1 = str(raw_fg1).strip()
        
        sche_val = sche_map.get(part_no, "OTHER")
        note_val = note_map.get(part_no, "-")
        
        # 📌 แยกคำนวณ WIP และ FG
        wip_val = wip_dict.get(part_no, 0)
        fg_val = wip_dict.get(fg1, 0)
        total_val = wip_val + fg_val
                  
        plan_arr = np.zeros(32)
        ord_arr = np.zeros(32)
        
        if include_plan:
            plan_arr[0] = plan_hist_dict.get(part_no, 0)
                              
        ord_arr[0] = ord_hist_dict.get(fg1, 0)
        total_orders_item = ord_arr[0]
                            
        for i in range(1, 32):
            d = pd.to_datetime(dates_row[i]).normalize()
            if include_plan:
                plan_arr[i] = plan_daily_dict.get((part_no, d), 0)

            day_ord = ord_daily_dict.get((fg1, d), 0)
            ord_arr[i] = day_ord
            total_orders_item += day_ord
            
        bl_arr = np.zeros(32)
        bl_arr[1] = total_val + plan_arr[0] - ord_arr[0] - ord_arr[1]
        for i in range(2, 32): bl_arr[i] = bl_arr[i-1] + plan_arr[i-1] - ord_arr[i]
            
        shot_date = "OK"
        for i in range(1, 32):
            if bl_arr[i] < 0:
                shot_date = dates_row[i].strftime('%Y-%m-%d')
                break
                
        row_data = {
            'SCHE': sche_val, 
            'Part No.': part_no, 
            'FG1': fg1, 
            'WIP': int(wip_val), 
            'FG': int(fg_val), 
            'รวม (WIP+FG)': int(total_val), 
            'Orders': int(total_orders_item), 
            'B/O Date': shot_date, # 📌 เปลี่ยนชื่อตัวแปรเป็น B/O Date
            'Note': note_val
        }
        
        for i in range(1, 32):
            date_key = dates_row[i].strftime('%Y-%m-%d')
            row_data[date_key] = int(bl_arr[i])
            
        dashboard_data.append(row_data)
        
    return pd.DataFrame(dashboard_data), [d.date() for d in dates_row[1:32]], df_missing

template_path = "Copy of daily check spc Aug26 rev2.2-1 .xlsx"
saved_up_file = "saved_database.xlsx"

# -------------------------------------------------------------
# 🔄 4. ระบบซิงค์ข้อมูลจาก GitHub ลง Local (ทำครั้งเดียวต่อ Session)
# -------------------------------------------------------------
if GITHUB_ENABLED and not st.session_state.get("github_synced"):
    with st.spinner("🔄 กำลังซิงค์ไฟล์ข้อมูลล่าสุดจาก GitHub..."):
        content, _ = gh_get_file(f"{GITHUB_DATA_DIR}/{saved_up_file}")
        if content:
            with open(saved_up_file, "wb") as f:
                f.write(content)
    st.session_state["github_synced"] = True
elif not GITHUB_ENABLED:
    pass

# แบ่งเลย์เอาต์ส่วนบน
col_header, col_filter, col_plan_toggle, col_upload = st.columns([1.6, 0.9, 0.9, 1.0])

target_file = None

with col_upload:
    st.markdown("<p style='font-weight:600; color:#475569; margin-bottom:5px;'>📂 อัปโหลดไฟล์ Database</p>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
    
    if uploaded_file is not None: 
        target_file = uploaded_file
        if st.button("💾 บันทึกไฟล์ข้อมูลนี้ไว้ใช้รอบหน้า", use_container_width=True):
            file_bytes = bytes(uploaded_file.getbuffer())
            with open(saved_up_file, "wb") as f:
                f.write(file_bytes)
                
            if GITHUB_ENABLED:
                with st.spinner("☁️ กำลังบันทึกไฟล์ขึ้น GitHub..."):
                    ok = gh_put_file(f"{GITHUB_DATA_DIR}/{saved_up_file}", file_bytes, f"Auto-save db upload: {uploaded_file.name}")
                if ok:
                    st.success("✅ บันทึกไฟล์ขึ้น GitHub เรียบร้อย!")
                else:
                    st.warning("⚠️ บันทึกลงระบบชั่วคราวได้ แต่ push ไป GitHub ไม่สำเร็จ")
            else:
                st.success("✅ บันทึกไฟล์เรียบร้อย!")
                
        st.caption("🟢 กำลังแสดงผลจาก: **ไฟล์ที่เพิ่งอัปโหลด**")
        
    elif os.path.exists(saved_up_file):
        target_file = saved_up_file
        st.caption("📌 กำลังแสดงผลจาก: **ไฟล์ที่บันทึกไว้ล่าสุด**")
        if st.button("🗑️ ล้างข้อมูลไฟล์ที่บันทึกไว้", use_container_width=True):
            os.remove(saved_up_file)
            st.rerun()

with col_plan_toggle:
    st.markdown("<p style='font-weight:600; color:#475569; margin-bottom:5px;'>⚙️ โหมดคำนวณ</p>", unsafe_allow_html=True)
    include_plan = st.checkbox("รวมแผนผลิต (Plan)", value=True)

# -------------------------------------------------------------
# 5. การแสดงผลและการคำนวณ Dashboard
# -------------------------------------------------------------
if target_file is None:
    with col_header:
        st.markdown(f"""
            <div class="dashboard-header">
                <h1 style="margin:0; font-size:26px; color:#0f172a;">📈 SPC Production Shortage Dashboard</h1>
                <p style="margin:0; color:#64748b; font-size:14px; margin-top:4px;">ระบบรันสำเร็จ! พร้อมใช้งานแล้ว</p>
            </div>
        """, unsafe_allow_html=True)
    st.info("👋 ยินดีต้อนรับ! ระบบไม่พบไฟล์ข้อมูลตั้งต้น **กรุณาอัปโหลดไฟล์ Database ประจำวัน** ที่ช่องมุมขวาบน เพื่อเริ่มต้นการแสดงผลครับ")
else:
    try:
        with st.spinner("กำลังประมวลผลข้อมูล..."):
            df_result, available_dates, df_missing = process_data(target_file, template_path, include_plan=include_plan)
        
        with col_filter:
            st.markdown("<p style='font-weight:600; color:#475569; margin-bottom:5px;'>📅 ดู Balance ถึงวันที่</p>", unsafe_allow_html=True)
            min_d = min(available_dates)
            max_d = max(available_dates)
            max_selectable_date = (pd.to_datetime(max_d) + pd.DateOffset(months=2)).date()
            
            selected_date = st.date_input("", value=max_d, min_value=min_d, max_value=max_selectable_date, format="DD/MM/YYYY", label_visibility="collapsed")
            selected_date_str = selected_date.strftime('%Y-%m-%d')
        
        df_result['B/O Date DT'] = pd.to_datetime(df_result['B/O Date'], errors='coerce')
        mask_short = df_result['B/O Date'] != 'OK'
        mask_date = df_result['B/O Date DT'].dt.date <= selected_date
        
        shortage_df = df_result[mask_short & mask_date].copy()
        
        total_short = len(shortage_df)
        total_orders_sum = shortage_df['Orders'].sum()
        
        mode_plan_text = " (รวมแผนผลิต)" if include_plan else " (ไม่รวมแผนผลิต)"
        
        with col_header:
            st.markdown(f"""
                <div class="dashboard-header">
                    <h1 style="margin:0; font-size:26px; color:#0f172a;">📈 SPC Production Shortage Dashboard</h1>
                    <p style="margin:0; color:#64748b; font-size:14px; margin-top:4px;">แสดงผลข้อมูลและสถานะ B/O Date{mode_plan_text}</p>
                </div>
            """, unsafe_allow_html=True)

        if not df_missing.empty:
            st.markdown(f'''
                <div class="alert-box">⚠️ <b>แจ้งเตือนความเสี่ยงหลุด Balance:</b> พบ {len(df_missing)} Part ที่มีรายการออเดอร์ (ord bac) แต่ไม่ได้ถูกบันทึกโครงสร้างไว้ใน Master (summary v1)</div>
            ''', unsafe_allow_html=True)
            with st.expander("คลิกเพื่อดูรายการ Part ที่ตกหล่น", expanded=False):
                st.dataframe(df_missing.style.map(lambda _: 'color: #b91c1c; font-weight: bold;', subset=['Missing Part (from Ord)']), hide_index=True)
            st.markdown("<br>", unsafe_allow_html=True)

        # 6. การ์ด KPI (ใช้ UI ใหม่)
        st.markdown("<div style='margin-bottom: -15px;'></div>", unsafe_allow_html=True)
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric("Part ที่ต้องระวัง (ติดลบ/B/O)", f"{total_short:,}")
        with col_m2:
            st.metric("จำนวน Order ค้างส่ง (ชิ้น)", f"{total_orders_sum:,}")
        with col_m3:
            st.metric("สถานะระบบ", "✨ ข้อมูลอัปเดตล่าสุด" if uploaded_file is not None else "🕒 ข้อมูลบันทึกล่าสุด")
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 📌 7. ระบบค้นหาสถานะ Part ขั้นสูง (หลายรายการ)
        st.markdown("<h3 style='color:#1e293b; font-size:18px;'>🔍 ค้นหาสถานะ Part ข้อมูลทั้งหมด (เทียบหลายรายการได้)</h3>", unsafe_allow_html=True)
        search_query = st.text_input("พิมพ์รหัส Part No. หรือ FG1 (คั่นด้วยลูกน้ำ ',' หากต้องการเทียบหลายตัว เช่น 1184469, BZ130)", "")
        
        if search_query:
            queries = [q.strip() for q in search_query.split(',') if q.strip()]
            if queries:
                search_mask = pd.Series(False, index=df_result.index)
                for q in queries:
                    mask = df_result['Part No.'].astype(str).str.contains(q, case=False, na=False) | \
                           df_result['FG1'].astype(str).str.contains(q, case=False, na=False)
                    search_mask = search_mask | mask
                
                searched_df = df_result[search_mask].copy()
                if not searched_df.empty:
                    # เรียงคอลัมน์และกรองวันให้แสดงผลสวยงาม
                    date_col_to_show_search = selected_date_str if selected_date_str in searched_df.columns else max(available_dates).strftime('%Y-%m-%d')
                    search_disp = searched_df[['SCHE', 'Part No.', 'FG1', 'WIP', 'FG', 'รวม (WIP+FG)', 'Orders', date_col_to_show_search, 'B/O Date', 'Note']].copy()
                    
                    col_bal_name_search = f'Balance ณ {selected_date.strftime("%d/%m/%Y")}'
                    search_disp.rename(columns={date_col_to_show_search: col_bal_name_search}, inplace=True)
                    
                    search_disp['B/O Date'] = pd.to_datetime(search_disp['B/O Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('OK')
                    
                    # ตกแต่งสีตารางค้นหา
                    styled_search = search_disp.style.map(
                        lambda x: 'background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center;' if x != 'OK' else '', 
                        subset=['B/O Date']
                    ).map(
                        lambda x: 'color: #b91c1c; font-weight: bold;' if isinstance(x, (int, float)) and x < 0 else 'color: #10b981; font-weight: bold;', 
                        subset=[col_bal_name_search]
                    )
                    st.dataframe(styled_search, use_container_width=True, hide_index=True)
                else:
                    st.warning(f"❌ ไม่พบข้อมูล Part ที่ตรงกับ '{search_query}' ในฐานข้อมูล")
        
        st.markdown("<hr style='margin-top:20px; margin-bottom:20px;'>", unsafe_allow_html=True)

        # 8. แบ่งหน้าจอ กราฟ & ตาราง
        left_col, right_col = st.columns([1, 2.5])

        with left_col:
            st.markdown("<h3 style='color:#1e293b; font-size:16px;'>แยกตามแผนก (SCHE)</h3>", unsafe_allow_html=True)
            if total_short > 0:
                sche_counts = shortage_df['SCHE'].value_counts().reset_index()
                sche_counts.columns = ['SCHE', 'Count']
                custom_colors = ['#ef4444', '#f97316', '#f59e0b', '#84cc16', '#10b981', '#06b6d4', '#3b82f6', '#8b5cf6', '#ec4899']
                fig = px.pie(sche_counts, values='Count', names='SCHE', hole=0.55, color_discrete_sequence=custom_colors)
                fig.update_traces(textinfo='none', hoverinfo='label+percent')
                fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=380, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5, font=dict(size=12)))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("ไม่มี Part ที่ติด B/O Date")

        with right_col:
            c_header, c_btn = st.columns([2.5, 1])
            with c_header:
                st.markdown(f"<h3 style='color:#1e293b; font-size:18px;'>รายการ Part ที่ติดลบ {mode_plan_text}</h3>", unsafe_allow_html=True)
            
            if total_short > 0:
                date_col_to_show = selected_date_str if selected_date_str in shortage_df.columns else max(available_dates).strftime('%Y-%m-%d')
                
                # 📌 สลับใช้คอลัมน์ที่แยก WIP และ FG ตามโจทย์
                display_df = shortage_df[['SCHE', 'Part No.', 'FG1', 'WIP', 'FG', 'รวม (WIP+FG)', 'Orders', date_col_to_show, 'B/O Date', 'Note']].copy()
                col_bal_name = f'Balance ณ {selected_date.strftime("%d/%m/%Y")}'
                display_df.rename(columns={date_col_to_show: col_bal_name}, inplace=True)
                
                display_df['B/O Date'] = pd.to_datetime(display_df['B/O Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('OK')
                
                with c_btn:
                    excel_data = to_excel(display_df)
                    st.download_button(label="📥 Download Excel", data=excel_data, file_name=f"Shortage_Report_{selected_date.strftime('%d_%m_%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
                dynamic_height = int((len(display_df) + 1) * 36) + 40
                if dynamic_height < 250: dynamic_height = 250
                
                # ตกแต่งสีตารางหลัก
                st.dataframe(
                    display_df.style.map(
                        lambda x: 'background-color: #fee2e2; color: #b91c1c; font-weight: bold; text-align: center; font-size: 14px;' if x != 'OK' else '', 
                        subset=['B/O Date']
                    ).map(
                        lambda x: 'color: #ef4444; font-weight: 500; font-size: 14px;' if isinstance(x, (int, float)) and x > 0 else 'font-size: 14px;',
                        subset=['Orders']
                    ).map(
                        lambda x: 'color: #2563eb; font-weight: 600; font-size: 14px;',
                        subset=['Part No.', 'FG1']
                    ).map(
                        lambda x: 'color: #b91c1c; font-weight: 700; font-size: 14px;' if isinstance(x, (int, float)) and x < 0 else 'color: #10b981; font-weight: 600; font-size: 14px;',
                        subset=[col_bal_name]
                    ).map(
                        lambda x: 'color: #64748b; font-size: 13px;',
                        subset=['Note']
                    ),
                    use_container_width=True,
                    height=dynamic_height,
                    hide_index=True,
                    column_config={
                        "WIP": st.column_config.NumberColumn("WIP", format="%d"),
                        "FG": st.column_config.NumberColumn("FG", format="%d"),
                        "รวม (WIP+FG)": st.column_config.NumberColumn("รวม", format="%d")
                    }
                )
            else:
                st.info("🎉 ไม่พบรายการ Part ที่ติดลบ (ยอดเยี่ยมมาก!)")

    except Exception as e:
        st.error(f"🚨 ระบบตรวจพบข้อผิดพลาด: {e}")
        st.warning("คำแนะนำ: โปรดตรวจสอบให้แน่ใจว่าคุณได้นำไฟล์ Master Template อัปโหลดเข้า GitHub หรือโฟลเดอร์ไว้แล้ว")
        st.exception(e)