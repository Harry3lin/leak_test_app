import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, BarChart

# 1. 設定網頁標題與配置
st.set_page_config(page_title="氣密測試數據自動化分析看板", layout="wide")
st.title("📊 氣密測試 (Leak Test) 數據自動化看板")

# 2. 頂部主選單
app_mode = st.radio(
    "請選擇分析模式：",
    ["📈 詳細趨勢曲線 (Line Plot Mode)", "📊 整體數據分佈 (Box Plot Mode)"],
    horizontal=True
)

st.write("---")

# ==============================================================================
# 模式一：Line Plot 趨勢曲線（新增自動產出包含圖表的多分頁 Excel）
# ==============================================================================
if app_mode == "📈 詳細趨勢曲線 (Line Plot Mode)":
    st.subheader("趨勢曲線分析 (Line Plot)")
    
    st.sidebar.header("⚙️ 趨勢圖控制")
    sample_step = st.sidebar.slider("📊 數據抽樣間隔 (Step)", min_value=1, max_value=10, value=1)
    
    uploaded_file = st.file_uploader("請上傳詳細製程數據 CSV 檔案 (Line Plot)", type=["csv"], key="line_uploader")
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            required_columns = ['SN', 'Timestamp', 'Pressure(Kpa)', 'Leak']
            missing_cols = [col for col in required_columns if col not in df.columns]
            
            if missing_cols:
                st.error(f"❌ 檔案格式不符，缺少必要欄位: {missing_cols}")
            else:
                df['Timestamp'] = pd.to_datetime(df['Timestamp'])
                
                # 自動忽略 Phase 包含 'Stabilization'
                if 'Phase' in df.columns:
                    initial_count = len(df)
                    df = df[~df['Phase'].astype(str).str.contains('Stabilization', case=False, na=False)]
                    filtered_count = initial_count - len(df)
                    st.sidebar.warning(f"🧹 已自動過濾掉 {filtered_count} 筆 Stabilization 數據。")
                
                df = df.sort_values(by='Timestamp')
                unique_sns = df['SN'].dropna().unique()
                st.success(f"✅ 成功讀取數據！偵測到共 {len(unique_sns)} 個不同的 SN 產品。")
                
                # ====== 核心新增：在後台動態生成包含圖表的 Excel 檔案 ======
                output_excel = io.BytesIO()
                wb = Workbook()
                # 移除預設的 Sheet
                wb.remove(wb.active)
                
                tabs = st.tabs([f"SN: {sn.split(':')[-1] if ':' in str(sn) else sn}" for sn in unique_sns])
                
                for i, sn in enumerate(unique_sns):
                    sn_data = df[df['SN'] == sn]
                    short_sn = str(sn).split(':')[-1] if ':' in str(sn) else str(sn)
                    # Excel 工作表名稱限制 31 個字元
                    sheet_name = short_sn[-30:]
                    
                    # 建立該 SN 的 Excel 工作表
                    ws = wb.create_sheet(title=sheet_name)
                    
                    # 寫入欄位名稱
                    headers = ['Timestamp', 'Pressure(Kpa)', 'Leak', 'bResult']
                    ws.append(headers)
                    
                    # 寫入數據
                    for _, row in sn_data.iterrows():
                        ws.append([
                            str(row['Timestamp']), 
                            row['Pressure(Kpa)'], 
                            row['Leak'] if pd.notna(row['Leak']) else "", 
                            row['bResult'] if 'bResult' in df.columns else ""
                        ])
                    
                    # 在 Excel 中自動為該分頁繪製 Pressure 折線圖
                    chart_p = LineChart()
                    chart_p.title = f"SN {sheet_name} - Pressure Trend"
                    chart_p.style = 13
                    chart_p.y_axis.title = "Pressure (Kpa)"
                    chart_p.x_axis.title = "Time"
                    data_p = Reference(ws, min_col=2, min_row=1, max_row=len(sn_data)+1)
                    chart_p.add_data(data_p, titles_from_data=True)
                    ws.add_chart(chart_p, "F2")  # 放在 F2 儲存格位置
                    
                    # 在 Excel 中自動為該分頁繪製 Leak 折線圖
                    if not sn_data['Leak'].dropna().empty:
                        chart_l = LineChart()
                        chart_l.title = f"SN {sheet_name} - Leak Trend"
                        chart_l.style = 13
                        chart_l.y_axis.title = "Leak"
                        chart_l.x_axis.title = "Time"
                        data_l = Reference(ws, min_col=3, min_row=1, max_row=len(sn_data)+1)
                        chart_l.add_data(data_l, titles_from_data=True)
                        ws.add_chart(chart_l, "N2")  # 放在 N2 儲存格位置（不與壓力圖重疊）
                    
                    # 網頁前端渲染圖表 (原有功能)
                    with tabs[i]:
                        st.subheader(f"產品序號: {sn}")
                        if sn_data.empty:
                            st.info("ℹ️ 該產品在排除 Stabilization 階段後無剩餘數據。")
                            continue
                        
                        plot_data = sn_data.iloc[::sample_step] if sample_step > 1 else sn_data
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("有效數據筆數", f"{len(sn_data)} 筆")
                        col2.metric("最大壓力 (Kpa)", f"{sn_data['Pressure(Kpa)'].max():.2f}")
                        col3.metric("測試結果狀態", f"{sn_data['bResult'].iloc[-1] if 'bResult' in sn_data.columns else '未知'}")
                        
                        st.write("---")
                        
                        # Pressure 圖
                        fig_press = go.Figure()
                        fig_press.add_trace(go.Scatter(x=plot_data['Timestamp'], y=plot_data['Pressure(Kpa)'], name='Pressure', mode='lines+markers', line=dict(color='#1f77b4', width=2)))
                        fig_press.update_layout(xaxis=dict(title="Timestamp"), yaxis=dict(title="Pressure (Kpa)"), hovermode="closest")
                        st.plotly_chart(fig_press, use_container_width=True)
                        
                        # Leak 圖
                        if not sn_data['Leak'].dropna().empty:
                            fig_leak = go.Figure()
                            fig_leak.add_trace(go.Scatter(x=plot_data['Timestamp'], y=plot_data['Leak'], name='Leak', mode='lines+markers', line=dict(color='#ff7f0e', width=2)))
                            fig_leak.update_layout(xaxis=dict(title="Timestamp"), yaxis=dict(title="Leak"), hovermode="closest")
                            st.plotly_chart(fig_leak, use_container_width=True)
                        else:
                            st.info("ℹ️ 此 SN 當前數據中的 Leak 欄位皆為空值。")
                
                # 儲存 Excel 檔案至記憶體
                wb.save(output_excel)
                output_excel.seek(0)
                
                # 顯示下載按鈕
                st.write("---")
                st.download_button(
                    label="📥 點此下載自動生成的『多分頁含圖表』Excel 報告",
                    data=output_excel,
                    file_name="Leak_Test_Line_Plots.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                            
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")

# ==============================================================================
# 模式二：Box Plot 箱線圖（新增自動產出統計報告 Excel）
# ==============================================================================
elif app_mode == "📊 整體數據分佈 (Box Plot Mode)":
    st.subheader("整體數據箱線圖分析 (Box Plot)")
    
    uploaded_file_box = st.file_uploader("請上傳統計數據 CSV 檔案 (Box Plot)", type=["csv"], key="box_uploader")
    
    if uploaded_file_box is not None:
        try:
            df_box = pd.read_csv(uploaded_file_box)
            required_box_cols = ['SN', 'Pressure(Kpa)', 'Leak']
            missing_box_cols = [col for col in required_box_cols if col not in df_box.columns]
            
            if missing_box_cols:
                st.error(f"❌ 檔案格式不符，缺少必要欄位: {missing_box_cols}")
            else:
                st.success(f"✅ 成功讀取統計數據！共有 {len(df_box)} 筆產品總點數。")
                
                if 'bResult' in df_box.columns:
                    ng_count = len(df_box[df_box['bResult'] == 'NG'])
                    ok_count = len(df_box[df_box['bResult'] == 'OK'])
                    col_b1, col_b2 = st.columns(2)
                    col_b1.metric("🟢 OK 總數", f"{ok_count} 筆")
                    col_b2.metric("🔴 NG 總數", f"{ng_count} 筆")
                
                st.write("---")
                
                # 網頁前端渲染圖表 (原有功能)
                graph_col1, graph_col2 = st.columns(2)
                with graph_col1:
                    st.markdown("#### 📦 Pressure (Kpa) 箱線分佈圖")
                    fig_box_press = px.box(df_box, y="Pressure(Kpa)", points="all", hover_data=["SN"], color="bResult" if "bResult" in df_box.columns else None, color_discrete_map={"OK": "#1f77b4", "NG": "#ef553b"})
                    st.plotly_chart(fig_box_press, use_container_width=True)
                    
                with graph_col2:
                    st.markdown("#### 📦 Leak 箱線分佈圖")
                    fig_box_leak = px.box(df_box, y="Leak", points="all", hover_data=["SN"], color="bResult" if "bResult" in df_box.columns else None, color_discrete_map={"OK": "#1f77b4", "NG": "#ef553b"})
                    st.plotly_chart(fig_box_leak, use_container_width=True)
                
                # ====== 核心新增：生成 Box Plot 數據的 Excel 報表 ======
                output_box_excel = io.BytesIO()
                wb_box = Workbook()
                ws_box = wb_box.active
                ws_box.title = "Summary Data"
                
                # 寫入數據表格
                ws_box.append(list(df_box.columns))
                for _, row in df_box.iterrows():
                    ws_box.append(list(row))
                
                wb_box.save(output_box_excel)
                output_box_excel.seek(0)
                
                # 顯示下載按鈕
