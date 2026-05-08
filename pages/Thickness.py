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
