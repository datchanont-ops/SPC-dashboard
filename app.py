import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io

# 1. ตั้งค่าหน้ากระดาษเป็น Wide Mode
st.set_page_config(page_title="Production Shortage Dashboard", layout="wide")

# 2. CSS ขั้นสูง (แก้ไขปัญหาฟอนต์ทับไอคอนระบบแล้ว)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    
    /* 🌟 เปลี่ยนฟอนต์เฉพาะข้อความ ยกเว้นไอคอนระบบ */
    html, body, p, h1, h2, h3, h4, h5, h6, label, th, td, div:not([class*="Icon"]):not([class*="icon"]) { 
        font-family: 'Prompt', sans-serif !important; 
    }
    
    /* 🌟 คืนค่าฟอนต์ไอคอนให้ Streamlit เพื่อแก้ปัญหา arrow_drop_down และเมนูทับกัน */
    .material-icons, .material-symbols-rounded, [class*="Icon"], [class*="icon"] {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
    }
    
    .stApp { background-color: #f1f5f9 !important; }
    
    /* Header Custom */
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
    
    /* KPI Cards Custom */
    .kpi-row {
        display: flex;
        gap: 20px;
        margin-bottom: 25px;
    }
    .kpi-card {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        flex: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
        min-height: 150px;
    }
    .kpi-label { color: #475569; font-size: 18px; font-weight: 600; margin-bottom: 8px; }
    .kpi-val { color: #0f172a; font-size: 44px; font-weight: 700; line-height: 1.2; }
    .kpi-val-text { color: #0f172a; font-size: 24px; font-weight: 700; line-height: 1.4; padding-top: 5px; }
    
    /* แถบขอบสีล่าง */
    .b-red { border-bottom: 5px solid #ef4444 !important; }
    .b-yellow { border-bottom: 5px solid #f59e0b !important; }
    .b-green { border-bottom: 5px solid #10b981 !important; }
    
    /* Upload & Filter Label */
    .filter-title {
        color: #334155;
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 5px;
        margin-top: 5px;
    }
    
    /* ตารางแสดงผลแบบเต็ม */
    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #e2e8f0;
    }
    
    /* ขนาดฟอนต์ */
    .stDateInput input { font-size: 15px !important; font-weight: 500 !important; cursor: pointer; }
    th { font-size: 16px !important; }
    td { font-size: 15px !important; }
    </style>
""", unsafe_allow_html=True)

# ฟังก์ชันสำหรับ Export Excel
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Shortage Report')
    processed_data = output.getvalue()
    return processed_data

# 3. ฟังก์ชันประมวลผลข้อมูล
@st.cache_data
def process_data(uploaded_file, file_template):
    xls = pd.ExcelFile(uploaded_file)
    def get_sheet(sheet_names, target):
        for s in sheet_names:
            if target in s.strip().lower(): return s
        return target

    df_ord = pd.read_excel(uploaded_file, sheet_name=get_sheet(xls.sheet_names, 'ord bac'))
    df_plan = pd.read_excel(uploaded_file, sheet_name=get_sheet(xls.sheet_names, 'plan'))
    df_wip = pd.read_excel(uploaded_file, sheet_name=get_sheet(xls.sheet_names, 'wip fg'))
    
    df_ord.columns = df_ord.columns.str.strip()
    df_plan.columns = df_plan.columns.str.strip()
    df_wip.columns = df_wip.columns.str.strip()

    def find_col(df, possible_names):
        for name in possible_names:
            for col in df.columns:
                if str(col).strip().lower() == str(name).strip().lower(): 
                    return col
        return None

    part_col_plan = find_col(df_plan, ['Part No.', 'Part no.', 'Partno']) or df_plan.columns[4]
    date_col_plan = find_col(df_plan, ['วันที่.1', 'date', 'วันที่']) or df_plan.columns[1]
    order_col_plan = find_col(df_plan, ['Order', 'Plan Order']) or df_plan.columns[5]

    mat_col_wip = find_col(df_wip, ['Material', 'Mat']) or df_wip.columns[0]
    unr_col_wip = find_col(df_wip, ['Unrestricted', 'Stock']) or df_wip.columns[5]

    sap_col_ord = find_col(df_ord, ['SAP Mat.', 'SAP Material', 'Material']) or df_ord.columns[2]
    dlv_col_ord = find_col(df_ord, ['Dlv. Date', 'Delivery Date', 'Date']) or df_ord.columns[6]
    qty_col_ord = find_col(df_ord, ['Outstd.Base Qty', 'Base Qty', 'Qty']) or df_ord.columns[5]

    df_sum_raw = pd.read_excel(file_template, sheet_name='summary v1', header=1)
    
    ord_parts = df_ord[sap_col_ord].dropna().astype(str).str.strip().str.upper().unique()
    sum_parts = df_sum_raw['fg1'].dropna().astype(str).str.strip().str.upper().unique()
    missing_set = set(ord_parts) - set(sum_parts)
    
    df_missing = pd.DataFrame()
    if missing_set:
        df_missing_raw = df_ord[df_ord[sap_col_ord].astype(str).str.strip().str.upper().isin(missing_set)]
        df_missing = df_missing_raw.groupby(sap_col_ord)[qty_col_ord].sum().reset_index()
        df_missing.rename(columns={sap_col_ord: 'Missing Part (from Ord)', qty_col_ord: 'Outstanding Qty'}, inplace=True)
    
    sche_map = {}
    note_map = {}
    for idx, row in df_sum_raw.iterrows():
        part = row.get('Part No.')
        sche = row.get('SCHE')
        note = row.get('Note')
        if pd.notna(part):
            p_key = str(part).strip()
            if pd.notna(sche): sche_map[p_key] = str(sche).strip()
            if pd.notna(note): note_map[p_key] = str(note).strip()

    plan_dates = pd.to_datetime(df_plan[date_col_plan], errors='coerce')
    if isinstance(plan_dates, pd.DataFrame): plan_dates = plan_dates.iloc[:, 0]
    df_plan['_parsed_date'] = plan_dates

    ord_dates = pd.to_datetime(df_ord[dlv_col_ord], errors='coerce')
    if isinstance(ord_dates, pd.DataFrame): ord_dates = ord_dates.iloc[:, 0]
    df_ord['_parsed_date'] = ord_dates

    df_sum_raw_headerless = pd.read_excel(file_template, sheet_name='summary v1', header=None)
    dates_row = pd.to_datetime(df_sum_raw_headerless.iloc[1, 6:38].values)
    start_date = pd.to_datetime('2025-01-01')
    today_date = dates_row[0]
    
    dashboard_data = []
    for idx in range(2, len(df_sum_raw_headerless), 3):
        if idx + 2 >= len(df_sum_raw_headerless): break
        part_no = df_sum_raw_headerless.iloc[idx, 2]
        fg1 = df_sum_raw_headerless.iloc[idx, 3]
        if pd.isna(part_no) and pd.isna(fg1): continue
        
        sche_val = sche_map.get(str(part_no).strip(), "OTHER")
        note_val = note_map.get(str(part_no).strip(), "-")
        
        wip_qty = df_wip[df_wip[mat_col_wip].astype(str).str.strip() == str(part_no).strip()][unr_col_wip].sum() + \
                  df_wip[df_wip[mat_col_wip].astype(str).str.strip() == str(fg1).strip()][unr_col_wip].sum()
                  
        plan_arr = np.zeros(32)
        ord_arr = np.zeros(32)
        
        p_mask = (df_plan[part_col_plan].astype(str).str.strip() == str(part_no).strip()) & (df_plan['_parsed_date'] >= start_date) & (df_plan['_parsed_date'] <= today_date)
        plan_arr[0] = df_plan.loc[p_mask, order_col_plan].sum()
                              
        o_mask = (df_ord[sap_col_ord].astype(str).str.strip() == str(fg1).strip()) & (df_ord['_parsed_date'] >= start_date) & (df_ord['_parsed_date'] <= today_date)
        ord_arr[0] = df_ord.loc[o_mask, qty_col_ord].sum()
        total_orders_item = ord_arr[0]
                            
        for i in range(1, 32):
            d = pd.to_datetime(dates_row[i]).normalize()
            p_day = (df_plan[part_col_plan].astype(str).str.strip() == str(part_no).strip()) & (df_plan['_parsed_date'].dt.normalize() == d)
            plan_arr[i] = df_plan.loc[p_day, order_col_plan].sum()

            o_day = (df_ord[sap_col_ord].astype(str).str.strip() == str(fg1).strip()) & (df_ord['_parsed_date'].dt.normalize() == d)
            day_ord = df_ord.loc[o_day, qty_col_ord].sum()
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

# 4. แบ่งเลย์เอาต์
col_header, col_filter, col_upload = st.columns([2.0, 1.0, 1.0])

target_file = "10-8.xlsx"
with col_upload:
    st.markdown('<div class="filter-title">📂 อัปโหลดไฟล์ Database</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload Database (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    if uploaded_file is not None: target_file = uploaded_file

try:
    df_result, available_dates, df_missing = process_data(target_file, template_path)
    
    with col_filter:
        st.markdown('<div class="filter-title">📅 ดู Balance ถึงวันที่ (คลิก)</div>', unsafe_allow_html=True)
        min_d = min(available_dates)
        max_d = max(available_dates)
        selected_date = st.date_input("Select date", value=max_d, min_value=min_d, max_value=max_d, label_visibility="collapsed")
        selected_date_str = selected_date.strftime('%Y-%m-%d')
    
    df_result['Short Date DT'] = pd.to_datetime(df_result['Short Date'], errors='coerce', format='%Y-%m-%d')
    mask_short = df_result['Short Date'] != 'OK'
    mask_date = df_result['Short Date DT'].dt.date <= selected_date
    
    shortage_df = df_result[mask_short & mask_date].copy()
    
    total_short = len(shortage_df)
    total_orders_sum = shortage_df['Orders'].sum()
    
    with col_header:
        st.markdown(f"""
            <div class="custom-header">
                <h2 style="margin:0; color:#0f172a; font-size: 30px; display:flex; align-items:center;">
                    📈 Production Shortage Dashboard
                </h2>
                <p style="margin:6px 0 0 0; color:#64748b; font-size: 16px;">แสดงผลข้อมูลและสถานะการ Short เฉพาะช่วงเวลาที่เลือก</p>
            </div>
        """, unsafe_allow_html=True)

    if not df_missing.empty:
        st.error(f"⚠️ **แจ้งเตือนความเสี่ยงหลุด Balance:** พบ {len(df_missing)} Part ที่มีรายการออเดอร์ (ord bac) แต่ไม่ได้ถูกบันทึกโครงสร้างไว้ใน Master (summary v1)")
        with st.expander("คลิกเพื่อดูรายการ Part ที่ตกหล่น", expanded=False):
            st.dataframe(df_missing.style.map(lambda _: 'color: #b91c1c; font-weight: bold;', subset=['Missing Part (from Ord)']), hide_index=True)
        st.markdown("<br>", unsafe_allow_html=True)

    mode_status = "✨ ข้อมูลอัปเดตล่าสุด<br>(Live File)" if uploaded_file is not None else "🕒 ข้อมูลเดิม<br>(Master File)"

    # 5. การ์ด KPI 
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

    # 6. แบ่งหน้าจอ
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
            st.plotly_chart(fig, width='stretch')
        else:
            st.success("ไม่มี Part ที่ติด Short")

    with right_col:
        col_title, col_export = st.columns([3, 1])
        with col_title:
            st.markdown("<h3 style='color:#1e293b; font-weight:600; font-size: 18px; margin-bottom:15px;'>รายการ Part ที่ติดลบ (Short Date)</h3>", unsafe_allow_html=True)
        
        if total_short > 0:
            display_df = shortage_df[['SCHE', 'Part No.', 'WIP+FG', 'Orders', selected_date_str, 'Short Date', 'Note']].copy()
            col_bal_name = f'Balance ณ {selected_date.strftime("%d %b")}'
            display_df.rename(columns={selected_date_str: col_bal_name}, inplace=True)
            
            with col_export:
                excel_data = to_excel(display_df)
                st.download_button(label="📥 Download Excel", data=excel_data, file_name=f"Shortage_Report_{selected_date.strftime('%d_%b')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
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
                    subset=['Part No.']
                ).map(
                    lambda x: 'color: #b91c1c; font-weight: 700; font-size: 14px;' if isinstance(x, (int, float)) and x < 0 else 'color: #10b981; font-weight: 600; font-size: 14px;',
                    subset=[col_bal_name]
                ).map(
                    lambda x: 'color: #64748b; font-size: 13px;',
                    subset=['Note']
                ),
                width='stretch',
                height=dynamic_height,
                hide_index=True
            )
        else:
            st.info("ไม่พบรายการ Part ที่ติดลบ")

except Exception as e:
    st.warning("กรุณาอัปโหลดไฟล์ Database ประจำวันเพื่อประมวลผล")