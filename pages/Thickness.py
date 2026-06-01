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
    # Shine
    # =============================
    if company == "บริษัท Shine":

        if df1.empty or df2.empty:
            st.warning(
                "ไฟล์ Shine ไม่สมบูรณ์ กรุณาตรวจสอบ sheet 'shine 1 MC' และ 'shine 2 MC'"
            )

        else:

            month_order = {
                "January": 1,
                "February": 2,
                "March": 3,
                "April": 4,
                "May": 5,
                "June": 6,
                "July": 7,
                "August": 8,
                "September": 9,
                "October": 10,
                "November": 11,
                "December": 12
            }

            # ====================================
            # SHINE 1
            # ====================================

            st.subheader("📊 Shine 1 MC")

            if "Month" not in df1.columns:
                if "Column2" in df1.columns:
                    df1["Month"] = df1["Column2"]

            col1_thick = "Au Thickness"

            if (
                "Month" in df1.columns
                and col1_thick in df1.columns
                and "ค่าMin" in df1.columns
            ):

                df1["Month"] = df1["Month"].astype(str).str.strip()

                df1[col1_thick] = pd.to_numeric(
                    df1[col1_thick],
                    errors="coerce"
                )

                df1["ค่าMin"] = pd.to_numeric(
                    df1["ค่าMin"],
                    errors="coerce"
                )

                c1, c2, c3, c4 = st.columns(4)

                c1.metric(
                    "Avg Thickness",
                    round(df1[col1_thick].mean(), 3)
                )

                c2.metric(
                    "Max Thickness",
                    round(df1[col1_thick].max(), 3)
                )

                c3.metric(
                    "Min Thickness",
                    round(df1[col1_thick].min(), 3)
                )

                c4.metric(
                    "Avg ค่าMin",
                    round(df1["ค่าMin"].mean(), 3)
                )

                df1_month = (
                    df1.groupby("Month", as_index=False)
                    .agg({
                        col1_thick: "mean",
                        "ค่าMin": "mean"
                    })
                )

                df1_month["MonthNum"] = (
                    df1_month["Month"]
                    .map(month_order)
                )

                df1_month = (
                    df1_month
                    .sort_values("MonthNum")
                )

                fig1 = go.Figure()

                fig1.add_trace(
                    go.Scatter(
                        x=df1_month["Month"],
                        y=df1_month[col1_thick],
                        mode="lines+markers",
                        name="Au Thickness"
                    )
                )

                fig1.add_trace(
                    go.Scatter(
                        x=df1_month["Month"],
                        y=df1_month["ค่าMin"],
                        mode="lines+markers",
                        name="ค่าMin"
                    )
                )

                fig1.add_hline(
                    y=1,
                    line=dict(color="red", dash="dash")
                )

                fig1.add_hline(
                    y=1.4,
                    line=dict(color="red", dash="dash")
                )

                fig1.update_layout(
                    title="Shine1 MC (Monthly Average)",
                    template="plotly_white"
                )

                st.plotly_chart(
                    fig1,
                    use_container_width=True
                )

            else:
                st.warning(
                    "ไม่พบคอลัมน์ Month / Au Thickness / ค่าMin ใน Shine1"
                )

            # ====================================
            # SHINE 2
            # ====================================

            st.subheader("📊 Shine 2 MC")

            if "Month" not in df2.columns:
                if "Column2" in df2.columns:
                    df2["Month"] = df2["Column2"]

            col2_thick = "Au thickness"

            if (
                "Month" in df2.columns
                and col2_thick in df2.columns
                and "ค่าMin" in df2.columns
            ):

                df2["Month"] = df2["Month"].astype(str).str.strip()

                df2[col2_thick] = pd.to_numeric(
                    df2[col2_thick],
                    errors="coerce"
                )

                df2["ค่าMin"] = pd.to_numeric(
                    df2["ค่าMin"],
                    errors="coerce"
                )

                c1, c2, c3, c4 = st.columns(4)

                c1.metric(
                    "Avg Thickness",
                    round(df2[col2_thick].mean(), 3)
                )

                c2.metric(
                    "Max Thickness",
                    round(df2[col2_thick].max(), 3)
                )

                c3.metric(
                    "Min Thickness",
                    round(df2[col2_thick].min(), 3)
                )

                c4.metric(
                    "Avg ค่าMin",
                    round(df2["ค่าMin"].mean(), 3)
                )

                df2_month = (
                    df2.groupby("Month", as_index=False)
                    .agg({
                        col2_thick: "mean",
                        "ค่าMin": "mean"
                    })
                )

                df2_month["MonthNum"] = (
                    df2_month["Month"]
                    .map(month_order)
                )

                df2_month = (
                    df2_month
                    .sort_values("MonthNum")
                )

                fig2 = go.Figure()

                fig2.add_trace(
                    go.Scatter(
                        x=df2_month["Month"],
                        y=df2_month[col2_thick],
                        mode="lines+markers",
                        name="Au Thickness"
                    )
                )

                fig2.add_trace(
                    go.Scatter(
                        x=df2_month["Month"],
                        y=df2_month["ค่าMin"],
                        mode="lines+markers",
                        name="ค่าMin"
                    )
                )

                fig2.add_hline(
                    y=2,
                    line=dict(color="red", dash="dash")
                )

                fig2.add_hline(
                    y=2.6,
                    line=dict(color="red", dash="dash")
                )

                fig2.update_layout(
                    title="Shine2 MC (Monthly Average)",
                    template="plotly_white"
                )

                st.plotly_chart(
                    fig2,
                    use_container_width=True
                )

            else:
                st.warning(
                    "ไม่พบคอลัมน์ Month / Au thickness / ค่าMin ใน Shine2"
                )

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
