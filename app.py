import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import os
import base64
import requests

# 1. ตั้งค่าหน้ากระดาษเป็น Wide Mode พร้อมชื่อเพจใหม่
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

# 2. CSS ขั้นสูง
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    
    html, body, p, h1, h2, h3, h4, h5, h6, label, th, td, div:not([class*="Icon"]):not([class*="icon"]) { 
        font-family: 'Prompt', sans-serif !important; 
    }
    
    .material-icons, .material-symbols-rounded, [class*="Icon"], [class*="icon"] {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
    }
    
    .stApp { background-color: #f1f5f9 !important; }
    
    .custom-header {
        background-color: #ffffff;
        padding: 24px 28px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        margin-bottom: 25px;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .kpi-row { display: flex; gap: 20px; margin-bottom: 25px; }
    .kpi-card {
        background-color: #ffffff; padding: 24px; border-radius: 12px;
        border: 1px solid #e2e8f0; box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        flex: 1; display: flex; flex-direction: column; justify-content: center; min-height: 150px;
    }
    .kpi-label { color: #475569; font-size: 18px; font-weight: 600; margin-bottom: 8px; }
    .kpi-val { color: #0f172a; font-size: 44px; font-weight: 700; line-height: 1.2; }
    .kpi-val-text { color: #0f172a; font-size: 24px; font-weight: 700; line-height: 1.4; padding-top: 5px; }
    
    .b-red { border-bottom: 5px solid #ef4444 !important; }
    .b-yellow { border-bottom: 5px solid #f59e0b !important; }
    .b-green { border-bottom: 5px solid #10b981 !important; }
    
    .filter-title { color: #334155; font-size: 18px; font-weight: 600; margin-bottom: 5px; margin-top: 5px; }
    div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; border: 1px solid #e2e8f0; }
    .stDateInput input { font-size: 15px !important; font-weight: 500 !important; cursor: pointer; }
    th { font-size: 16px !important; }
    td { font-size: 15px !important; }
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
        
        wip_qty = wip_dict.get(part_no, 0) + wip_dict.get(fg1, 0)
                  
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
        bl_arr[1] = wip_qty + plan_arr[0] - ord_arr[0] - ord_arr[1]
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
            'WIP+FG': int(wip_qty), 
            'Orders': int(total_orders_item), 
            'Short Date': shot_date,
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
    # หากยังไม่ได้เชื่อมต่อ GitHub ระบบจะแจ้งเตือนเบาๆ แต่ยังทำงานแบบ Local ได้
    pass

# แบ่งเลย์เอาต์
col_header, col_filter, col_plan_toggle, col_upload = st.columns([1.6, 0.9, 0.9, 1.0])

target_file = None

with col_upload:
    st.markdown('<div class="filter-title">📂 อัปโหลดไฟล์ Database</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
    
    # 📌 ระบบจำไฟล์ข้อมูลลง GitHub / Local
    if uploaded_file is not None: 
        target_file = uploaded_file
        if st.button("💾 บันทึกไฟล์ข้อมูลนี้ไว้ใช้รอบหน้า", use_container_width=True):
            file_bytes = bytes(uploaded_file.getbuffer())
            
            # 1. บันทึกลงเครื่องเซิร์ฟเวอร์ชั่วคราว
            with open(saved_up_file, "wb") as f:
                f.write(file_bytes)
                
            # 2. ส่งขึ้น GitHub เพื่อจำแบบถาวร
            if GITHUB_ENABLED:
                with st.spinner("☁️ กำลังบันทึกไฟล์ขึ้น GitHub..."):
                    ok = gh_put_file(f"{GITHUB_DATA_DIR}/{saved_up_file}", file_bytes, f"Auto-save db upload: {uploaded_file.name}")
                if ok:
                    st.success("✅ บันทึกไฟล์ขึ้น GitHub เรียบร้อย! คราวหน้าไม่ต้องอัปโหลดซ้ำแล้วครับ")
                else:
                    st.warning("⚠️ บันทึกลงระบบชั่วคราวได้ แต่ push ไป GitHub ไม่สำเร็จ")
            else:
                st.success("✅ บันทึกไฟล์เรียบร้อย! (บันทึกแค่ในเซิร์ฟเวอร์ชั่วคราว เนื่องจากไม่ได้ต่อ GitHub)")
                
        st.caption("🟢 กำลังแสดงผลจาก: **ไฟล์ที่เพิ่งอัปโหลด**")
        
    elif os.path.exists(saved_up_file):
        target_file = saved_up_file
        st.caption("📌 กำลังแสดงผลจาก: **ไฟล์ที่บันทึกไว้ล่าสุด**")
        if st.button("🗑️ ล้างข้อมูลไฟล์ที่บันทึกไว้", use_container_width=True):
            os.remove(saved_up_file)
            st.rerun()

with col_plan_toggle:
    st.markdown('<div class="filter-title">⚙️ โหมดคำนวณ</div>', unsafe_allow_html=True)
    include_plan = st.checkbox("รวมแผนผลิต (Plan)", value=True)

# -------------------------------------------------------------
# 5. การแสดงผลและการคำนวณ Dashboard
# -------------------------------------------------------------
if target_file is None:
    with col_header:
        st.markdown(f"""
            <div class="custom-header">
                <h2 style="margin:0; color:#0f172a; font-size: 28px; display:flex; align-items:center;">
                    📈 SPC Production Shortage Dashboard
                </h2>
                <p style="margin:6px 0 0 0; color:#64748b; font-size: 15px;">ระบบรันสำเร็จ! พร้อมใช้งานแล้ว</p>
            </div>
        """, unsafe_allow_html=True)
    st.info("👋 ยินดีต้อนรับ! ระบบไม่พบไฟล์ข้อมูลตั้งต้น **กรุณาอัปโหลดไฟล์ Database ประจำวัน** ที่ช่องมุมขวาบน เพื่อเริ่มต้นการแสดงผลครับ")
    st.warning("💡 หมายเหตุ: หากอัปโหลดไฟล์แล้วยัง Error กรุณาตรวจสอบให้แน่ใจว่าใน GitHub ของคุณมีไฟล์ `Copy of daily check spc Aug26 rev2.2-1 .xlsx` (Master Template) เก็บไว้ใน Repository แล้ว")
else:
    try:
        with st.spinner("กำลังประมวลผลข้อมูล..."):
            df_result, available_dates, df_missing = process_data(target_file, template_path, include_plan=include_plan)
        
        with col_filter:
            st.markdown('<div class="filter-title">📅 ดู Balance ถึงวันที่</div>', unsafe_allow_html=True)
            min_d = min(available_dates)
            max_d = max(available_dates)
            
            # ขยายวันที่ในปฏิทินให้กดเลือกล่วงหน้าเพิ่มได้อีก 2 เดือน
            max_selectable_date = (pd.to_datetime(max_d) + pd.DateOffset(months=2)).date()
            
            selected_date = st.date_input("", value=max_d, min_value=min_d, max_value=max_selectable_date, format="DD/MM/YYYY", label_visibility="collapsed")
            selected_date_str = selected_date.strftime('%Y-%m-%d')
        
        df_result['Short Date DT'] = pd.to_datetime(df_result['Short Date'], errors='coerce')
        mask_short = df_result['Short Date'] != 'OK'
        mask_date = df_result['Short Date DT'].dt.date <= selected_date
        
        shortage_df = df_result[mask_short & mask_date].copy()
        
        total_short = len(shortage_df)
        total_orders_sum = shortage_df['Orders'].sum()
        
        mode_plan_text = " (รวมแผนผลิต)" if include_plan else " (ไม่รวมแผนผลิต)"
        
        with col_header:
            st.markdown(f"""
                <div class="custom-header">
                    <h2 style="margin:0; color:#0f172a; font-size: 28px; display:flex; align-items:center;">
                        📈 SPC Production Shortage Dashboard
                    </h2>
                    <p style="margin:6px 0 0 0; color:#64748b; font-size: 15px;">แสดงผลข้อมูลและสถานะการ Short{mode_plan_text}</p>
                </div>
            """, unsafe_allow_html=True)

        if not df_missing.empty:
            st.error(f"⚠️ **แจ้งเตือนความเสี่ยงหลุด Balance:** พบ {len(df_missing)} Part ที่มีรายการออเดอร์ (ord bac) แต่ไม่ได้ถูกบันทึกโครงสร้างไว้ใน Master (summary v1)")
            with st.expander("คลิกเพื่อดูรายการ Part ที่ตกหล่น", expanded=False):
                st.dataframe(df_missing.style.map(lambda _: 'color: #b91c1c; font-weight: bold;', subset=['Missing Part (from Ord)']), hide_index=True)
            st.markdown("<br>", unsafe_allow_html=True)

        mode_status = "✨ ข้อมูลอัปเดตล่าสุด<br>(Live File)" if uploaded_file is not None else "🕒 ข้อมูลเดิม<br>(Master File)"

        # 6. การ์ด KPI 
        st.markdown(f"""
            <div class="kpi-row">
                <div class="kpi-card b-red">
                    <div class="kpi-label">Part ที่สถานะ Short (จนถึง {selected_date.strftime('%d/%m/%Y')})</div>
                    <div class="kpi-val">{total_short:,}</div>
                </div>
                <div class="kpi-card b-yellow">
                    <div class="kpi-label">จำนวน Order ค้างส่ง (ชิ้น)</div>
                    <div class="kpi-val">{total_orders_sum:,}</div>
                </div>
                <div class="kpi-card b-green">
                    <div class="kpi-label">สถานะระบบ</div>
                    <div class="kpi-val-text">{mode_status}</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # 7. แบ่งหน้าจอ
        left_col, right_col = st.columns([1.2, 2.8])

        with left_col:
            st.markdown("<h3 style='color:#1e293b; font-weight:600; font-size: 18px; margin-bottom:15px;'>แยกตามแผนก (SCHE)</h3>", unsafe_allow_html=True)
            if total_short > 0:
                sche_counts = shortage_df['SCHE'].value_counts().reset_index()
                sche_counts.columns = ['SCHE', 'Count']
                custom_colors = ['#ef4444', '#f97316', '#f59e0b', '#84cc16', '#10b981', '#06b6d4', '#3b82f6', '#6366f1', '#8b5cf6', '#ec4899']
                fig = px.pie(sche_counts, values='Count', names='SCHE', hole=0.55, color_discrete_sequence=custom_colors)
                fig.update_traces(textinfo='none', hoverinfo='label+percent')
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=450, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.0, font=dict(size=14)))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("ไม่มี Part ที่ติด Short")

        with right_col:
            col_title, col_export = st.columns([3, 1])
            with col_title:
                st.markdown(f"<h3 style='color:#1e293b; font-weight:600; font-size: 18px; margin-bottom:15px;'>รายการ Part ที่ติดลบ {mode_plan_text}</h3>", unsafe_allow_html=True)
            
            if total_short > 0:
                date_col_to_show = selected_date_str if selected_date_str in shortage_df.columns else max(available_dates).strftime('%Y-%m-%d')
                
                display_df = shortage_df[['SCHE', 'Part No.', 'FG1', 'WIP+FG', 'Orders', date_col_to_show, 'Short Date', 'Note']].copy()
                col_bal_name = f'Balance ณ {selected_date.strftime("%d/%m/%Y")}'
                display_df.rename(columns={date_col_to_show: col_bal_name}, inplace=True)
                
                display_df['Short Date'] = pd.to_datetime(display_df['Short Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna('OK')
                
                with col_export:
                    excel_data = to_excel(display_df)
                    st.download_button(label="📥 Download Excel", data=excel_data, file_name=f"Shortage_Report_{selected_date.strftime('%d_%m_%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                dynamic_height = int((len(display_df) + 1) * 36) + 40
                
                st.dataframe(
                    display_df.style.map(
                        lambda x: 'color: #b91c1c; background-color: #fee2e2; font-weight: 600; text-align: center; font-size: 14px;' if x != 'OK' else '', 
                        subset=['Short Date']
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
                    hide_index=True
                )
            else:
                st.info("ไม่พบรายการ Part ที่ติดลบ")

    except Exception as e:
        st.error(f"🚨 ระบบตรวจพบข้อผิดพลาด: {e}")
        st.warning("คำแนะนำ: โปรดตรวจสอบให้แน่ใจว่าคุณได้นำไฟล์ Master Template (Copy of daily check...) อัปโหลดเข้า GitHub ไว้ในโฟลเดอร์เดียวกับโค้ดแล้ว หรือตรวจสอบว่าไฟล์ที่อัปโหลดมี Sheet ครบถ้วน")
        st.exception(e)