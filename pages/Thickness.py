import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(layout="wide")
st.title("📊 Dashboard")

# =============================
# Company selector
# =============================
if "company" not in st.session_state:
    st.session_state["company"] = "บริษัท Shine"

company = st.selectbox(
    "เลือกบริษัท",
    ["บริษัท Shine", "บริษัท MIKI"],
    index=["บริษัท Shine", "บริษัท MIKI"].index(st.session_state["company"])
)
st.session_state["company"] = company

# =============================
# Upload file
# =============================
file = st.file_uploader("Upload Excel", type=["xlsx"])

def make_month_cols(df, date_col):
    """แปลงคอลัมน์วันที่เป็น Month (เต็ม) และ MonthNum สำหรับเรียงลำดับ"""
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df["Month"] = df[date_col].dt.strftime("%B")   # March, April, ...
    df["MonthNum"] = df[date_col].dt.month
    return df

if file:
    # อ่าน sheet (ถ้าไม่มี sheet จะเป็น DataFrame ว่าง)
    try:
        df1 = pd.read_excel(file, sheet_name="shine 1 MC", engine="openpyxl")
    except Exception:
        df1 = pd.DataFrame()
    try:
        df2 = pd.read_excel(file, sheet_name="shine 2 MC", engine="openpyxl")
    except Exception:
        df2 = pd.DataFrame()
    try:
        df_miki = pd.read_excel(file, sheet_name="thick MIKI", engine="openpyxl")
    except Exception:
        df_miki = pd.DataFrame()

    st.success("✅ โหลดไฟล์เสร็จแล้ว")

    # clean columns
    if not df1.empty:
        df1.columns = df1.columns.str.strip()
    if not df2.empty:
        df2.columns = df2.columns.str.strip()
    if not df_miki.empty:
        df_miki.columns = df_miki.columns.str.strip()

    # =============================
    # Shine (แสดง KPI ค่าเฉลี่ยของ thickness และ ค่าMin + กราฟ)
    # =============================
    if company == "บริษัท Shine":
        if df1.empty or df2.empty:
            st.warning("ไฟล์ Shine ไม่สมบูรณ์ กรุณาตรวจสอบ sheet 'shine 1 MC' และ 'shine 2 MC'")
        else:
            # หา column thickness อัตโนมัติ
            try:
                col1_thick = [c for c in df1.columns if "thickness" in c.lower() or "au thickness" in c.lower()][0]
            except IndexError:
                col1_thick = None
            try:
                col2_thick = [c for c in df2.columns if "thickness" in c.lower() or "au thickness" in c.lower()][0]
            except IndexError:
                col2_thick = None

            # หา Date column อัตโนมัติ แล้วสร้าง Month + MonthNum (ใช้ชื่อคอลัมน์ที่เจอ)
            try:
                date_col1 = [c for c in df1.columns if "date" in c.lower() or "วัน" in c][0]
                df1 = make_month_cols(df1, date_col1)
            except Exception:
                df1["Month"] = pd.NA
                df1["MonthNum"] = pd.NA

            try:
                date_col2 = [c for c in df2.columns if "date" in c.lower() or "วัน" in c][0]
                df2 = make_month_cols(df2, date_col2)
            except Exception:
                df2["Month"] = pd.NA
                df2["MonthNum"] = pd.NA

            # -----------------------------
            # KPI ของ Shine (แสดง Avg ของ thickness และ Avg ของ ค่าMin)
            # -----------------------------
            colA, colB = st.columns(2)
            with colA:
                if col1_thick and col1_thick in df1.columns:
                    avg_thick1 = df1[col1_thick].mean()
                    max_thick1 = df1[col1_thick].max()
                    min_thick1 = df1[col1_thick].min()
                else:
                    avg_thick1 = max_thick1 = min_thick1 = None

                if "ค่าMin" in df1.columns:
                    avg_min1 = df1["ค่าMin"].mean()
                else:
                    avg_min1 = None

                # แสดงกล่อง KPI
                st.markdown(f"""
                <div style='border:2px solid #4CAF50; padding:10px; border-radius:8px; text-align:center;'>
                    <div style='font-weight:600; margin-bottom:6px;'>Shine1 MC</div>
                    <div>Thickness Avg: <b>{round(avg_thick1,2) if avg_thick1 is not None else '-'}</b></div>
                    <div>Thickness Max: <b>{round(max_thick1,2) if max_thick1 is not None else '-'}</b></div>
                    <div>Thickness Min: <b>{round(min_thick1,2) if min_thick1 is not None else '-'}</b></div>
                    <div style='margin-top:6px;'>ค่าMin Avg: <b>{round(avg_min1,2) if avg_min1 is not None else '-'}</b></div>
                </div>
                """, unsafe_allow_html=True)

            with colB:
                if col2_thick and col2_thick in df2.columns:
                    avg_thick2 = df2[col2_thick].mean()
                    max_thick2 = df2[col2_thick].max()
                    min_thick2 = df2[col2_thick].min()
                else:
                    avg_thick2 = max_thick2 = min_thick2 = None

                if "ค่าMin" in df2.columns:
                    avg_min2 = df2["ค่าMin"].mean()
                else:
                    avg_min2 = None

                st.markdown(f"""
                <div style='border:2px solid #2196F3; padding:10px; border-radius:8px; text-align:center;'>
                    <div style='font-weight:600; margin-bottom:6px;'>Shine2 MC</div>
                    <div>Thickness Avg: <b>{round(avg_thick2,2) if avg_thick2 is not None else '-'}</b></div>
                    <div>Thickness Max: <b>{round(max_thick2,2) if max_thick2 is not None else '-'}</b></div>
                    <div>Thickness Min: <b>{round(min_thick2,2) if min_thick2 is not None else '-'}</b></div>
                    <div style='margin-top:6px;'>ค่าMin Avg: <b>{round(avg_min2,2) if avg_min2 is not None else '-'}</b></div>
                </div>
                """, unsafe_allow_html=True)

            # -------- Shine1 graph (เต็มความกว้าง) --------
            if "MonthNum" in df1.columns and col1_thick and "ค่าMin" in df1.columns:
                df1_m = df1.groupby(["MonthNum", "Month"], as_index=False).agg({col1_thick: "mean", "ค่าMin": "mean"})
                df1_m = df1_m.sort_values("MonthNum")
                fig1 = go.Figure()
                fig1.add_trace(go.Scatter(x=df1_m["Month"], y=df1_m[col1_thick], mode="lines+markers", name=col1_thick))
                fig1.add_trace(go.Scatter(x=df1_m["Month"], y=df1_m["ค่าMin"], mode="lines+markers", name="ค่าMin"))
                fig1.add_hline(y=1, line=dict(color="red", dash="dash"))
                fig1.add_hline(y=1.4, line=dict(color="red", dash="dash"))
                fig1.update_layout(title="Shine1 MC (รายเดือน)", template="plotly_white")
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("ไม่มีข้อมูลเพียงพอสำหรับกราฟ Shine1 (ต้องมี Month และ ค่าMin และ thickness)")

            # -------- Shine2 graph (เต็มความกว้าง) --------
            if "MonthNum" in df2.columns and col2_thick and "ค่าMin" in df2.columns:
                df2_m = df2.groupby(["MonthNum", "Month"], as_index=False).agg({col2_thick: "mean", "ค่าMin": "mean"})
                df2_m = df2_m.sort_values("MonthNum")
                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(x=df2_m["Month"], y=df2_m[col2_thick], mode="lines+markers", name=col2_thick))
                fig2.add_trace(go.Scatter(x=df2_m["Month"], y=df2_m["ค่าMin"], mode="lines+markers", name="ค่าMin"))
                fig2.add_hline(y=2, line=dict(color="red", dash="dash"))
                fig2.add_hline(y=2.6, line=dict(color="red", dash="dash"))
                fig2.update_layout(title="Shine2 MC (รายเดือน)", template="plotly_white")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("ไม่มีข้อมูลเพียงพอสำหรับกราฟ Shine2 (ต้องมี Month และ ค่าMin และ thickness)")

    # =============================
    # MIKI (ใช้คอลัมน์ Month โดยตรง)
    # =============================
    elif company == "บริษัท MIKI":
        if df_miki.empty:
            st.error("❌ ไม่มี sheet 'thick MIKI' หรือไฟล์ไม่มีข้อมูล")
            st.stop()

        # ใช้ column เดือนตรง ๆ
        if "Month" not in df_miki.columns:
            st.error("❌ ไม่มีคอลัมน์ Month ใน sheet 'thick MIKI'")
            st.stop()

        # สร้าง MonthNum สำหรับเรียงลำดับ (1–12)
        month_map = {
            "January": 1, "February": 2, "March": 3, "April": 4,
            "May": 5, "June": 6, "July": 7, "August": 8,
            "September": 9, "October": 10, "November": 11, "December": 12
        }
        df_miki["MonthNum"] = df_miki["Month"].map(month_map)

        # mapping YG/RG ของแต่ละขนาด
        cols_map = {
            "0.5": ["YG 0.5 mc", "RG 0.5 mc"],
            "1": ["YG 1 mc", "RG 1 mc"],
            "2": ["YG 2 mc", "RG 2 mc"],
            "3": ["YG 3 mc", "RG 3 mc"]
        }

        # แปลงเป็นตัวเลข (ถ้ามี)
        for group in cols_map.values():
            for col in group:
                if col in df_miki.columns:
                    df_miki[col] = pd.to_numeric(df_miki[col], errors="coerce")

        # --- แสดง KPI เป็นคอลัมน์แนวนอน 4 คอลัมน์ (0.5,1,2,3) ---
        display_cols = st.columns(4)
        for i, (label, group) in enumerate(cols_map.items()):
            yg_col, rg_col = group[0], group[1]

            yg_block = ""
            rg_block = ""
            if yg_col in df_miki.columns:
                yg_avg = df_miki[yg_col].mean()
                yg_max = df_miki[yg_col].max()
                yg_min = df_miki[yg_col].min()
                yg_block = f"""
                <div style='border:1px solid #4CAF50; padding:6px; border-radius:6px; text-align:center; font-size:13px; margin-bottom:4px;'>
                <div style='font-weight:600'>{yg_col}</div>
                <div>Avg: <b>{round(yg_avg,2)}</b></div>
                <div>Max: <b>{round(yg_max,2)}</b></div>
                <div>Min: <b>{round(yg_min,2)}</b></div>
                </div>
                """
            else:
                yg_block = f"<div style='color:#777; text-align:center; font-size:13px; margin-bottom:4px;'>ไม่มี</div>"

            if rg_col in df_miki.columns:
                rg_avg = df_miki[rg_col].mean()
                rg_max = df_miki[rg_col].max()
                rg_min = df_miki[rg_col].min()
                rg_block = f"""
                <div style='border:1px solid #2196F3; padding:6px; border-radius:6px; text-align:center; font-size:13px;'>
                <div style='font-weight:600'>{rg_col}</div>
                <div>Avg: <b>{round(rg_avg,2)}</b></div>
                <div>Max: <b>{round(rg_max,2)}</b></div>
                <div>Min: <b>{round(rg_min,2)}</b></div>
                </div>
                """
            else:
                rg_block = f"<div style='color:#777; text-align:center; font-size:13px;'>ไม่มี</div>"

            html = f"""
            <div style='text-align:center; font-weight:600; margin-bottom:6px;'>{label} mc</div>
            <div style='display:flex; justify-content:space-between; gap:6px;'>
            <div style='flex:1'>{yg_block}</div>
            <div style='flex:1'>{rg_block}</div>
            </div>
            """
            display_cols[i].markdown(html, unsafe_allow_html=True)

        # =============================
        # Dropdown เลือกกราฟ (0.5 / 1 / 2 / 3) — แยกกราฟเป็นคนละชาร์ท เรียงบน-ล่าง
        # =============================
        st.subheader("📈 กราฟ MIKI (เลือกความหนา)")
        options = list(cols_map.keys())
        selected = st.selectbox("เลือกความหนา", options)

        if selected:
            available = [c for c in cols_map[selected] if c in df_miki.columns]
            if not available:
                st.warning("❌ ไม่มีข้อมูลของขนาดนี้ในไฟล์")
            else:
                # แสดงกราฟแต่ละ series แยกกัน เรียงบน-ล่าง (full width)
                for col in available:
                    df_clean = df_miki[["MonthNum", "Month", col]].dropna()
                    if df_clean.empty:
                        st.info(f"ไม่มีข้อมูล {col}")
                        continue

                    df_group = df_clean.groupby(["MonthNum", "Month"], as_index=False)[col].mean()
                    df_group = df_group.sort_values("MonthNum")
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=df_group["Month"], y=df_group[col], mode="lines+markers", name=col))
                    # เส้นมาตรฐานตามขนาด
                    if selected == "0.5":
                        fig.add_hline(y=0.4, line=dict(color="red", dash="dash"))
                        fig.add_hline(y=0.6, line=dict(color="red", dash="dash"))
                    elif selected == "1":
                        fig.add_hline(y=0.8, line=dict(color="red", dash="dash"))
                        fig.add_hline(y=1.2, line=dict(color="red", dash="dash"))
                    elif selected == "2":
                        fig.add_hline(y=1.6, line=dict(color="red", dash="dash"))
                        fig.add_hline(y=2.4, line=dict(color="red", dash="dash"))
                    elif selected == "3":
                        fig.add_hline(y=2.4, line=dict(color="red", dash="dash"))
                        fig.add_hline(y=3.6, line=dict(color="red", dash="dash"))

                    fig.update_layout(title=col, template="plotly_white")
                    st.plotly_chart(fig, use_container_width=True)
