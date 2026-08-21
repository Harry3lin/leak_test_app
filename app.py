import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# 1. 設定網頁標題與配置
st.set_page_config(page_title="氣密測試數據自動化分析看板", layout="wide")
st.title("📊 氣密測試 (Leak Test) 數據自動化看板")

# 2. 頂部主選單：切換不同的分析功能
app_mode = st.radio(
    "請選擇分析模式：",
    ["📈 詳細趨勢曲線 (Line Plot Mode)", "📊 整體數據分佈 (Box Plot Mode)"],
    horizontal=True
)

st.write("---")

# ==============================================================================
# 模式一：Line Plot 趨勢曲線（原有功能，包含過濾 Stabilization）
# ==============================================================================
if app_mode == "📈 詳細趨勢曲線 (Line Plot Mode)":
    st.subheader("趨勢曲線分析 (Line Plot)")
    
    # 側邊欄控制
    st.sidebar.header("⚙️ 趨勢圖控制")
    sample_step = st.sidebar.slider(
        "📊 數據抽樣間隔 (Step)", 
        min_value=1, max_value=10, value=1, 
        help="1 代表顯示原始數據。若圖表卡頓，建議調大此數值。"
    )
    
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
                
                # 自動忽略 Phase 包含 'Stabilization' 的數據
                if 'Phase' in df.columns:
                    initial_count = len(df)
                    df = df[~df['Phase'].astype(str).str.contains('Stabilization', case=False, na=False)]
                    filtered_count = initial_count - len(df)
                    st.sidebar.warning(f"🧹 已自動過濾掉 {filtered_count} 筆 Stabilization 數據。")
                
                df = df.sort_values(by='Timestamp')
                unique_sns = df['SN'].dropna().unique()
                st.success(f"✅ 成功讀取數據！偵測到共 {len(unique_sns)} 個不同的 SN 產品。")
                
                tabs = st.tabs([f"SN: {sn.split(':')[-1] if ':' in str(sn) else sn}" for sn in unique_sns])
                
                for i, sn in enumerate(unique_sns):
                    with tabs[i]:
                        st.subheader(f"產品序號: {sn}")
                        sn_data = df[df['SN'] == sn]
                        
                        if sn_data.empty:
                            st.info("ℹ️ 該產品在排除 Stabilization 階段後無剩餘數據。")
                            continue
                        
                        if sample_step > 1:
                            plot_data = sn_data.iloc[::sample_step]
                        else:
                            plot_data = sn_data
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("有效數據筆數", f"{len(sn_data)} 筆")
                        col2.metric("最大壓力 (Kpa)", f"{sn_data['Pressure(Kpa)'].max():.2f}")
                        col3.metric("測試結果狀態", f"{sn_data['bResult'].iloc[-1] if 'bResult' in sn_data.columns else '未知'}")
                        
                        st.write("---")
                        
                        # Pressure 圖
                        st.markdown("#### 📈 Pressure (Kpa) 趨勢變化圖")
                        fig_press = go.Figure()
                        fig_press.add_trace(go.Scatter(
                            x=plot_data['Timestamp'], y=plot_data['Pressure(Kpa)'],
                            name='Pressure (Kpa)', mode='lines+markers',
                            line=dict(color='#1f77b4', width=2), marker=dict(size=5)
                        ))
                        fig_press.update_layout(xaxis=dict(title="時間軸 (Timestamp)"), yaxis=dict(title="Pressure (Kpa)"), hovermode="closest")
                        st.plotly_chart(fig_press, use_container_width=True)
                        
                        # Leak 圖
                        st.markdown("#### 📉 Leak 趨勢變化圖")
                        if not sn_data['Leak'].dropna().empty:
                            fig_leak = go.Figure()
                            fig_leak.add_trace(go.Scatter(
                                x=plot_data['Timestamp'], y=plot_data['Leak'],
                                name='Leak', mode='lines+markers',
                                line=dict(color='#ff7f0e', width=2), marker=dict(size=5)
                            ))
                            fig_leak.update_layout(xaxis=dict(title="時間軸 (Timestamp)"), yaxis=dict(title="Leak"), hovermode="closest")
                            st.plotly_chart(fig_leak, use_container_width=True)
                        else:
                            st.info("ℹ️ 此 SN 當前數據中的 Leak 欄位皆為空值，暫不顯示圖表。")
                            
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")

# ==============================================================================
# 模式二：Box Plot 箱線圖（全新新增功能）
# ==============================================================================
elif app_mode == "📊 整體數據分佈 (Box Plot Mode)":
    st.subheader("整體數據箱線圖分析 (Box Plot)")
    st.write("請在上傳區放入摘要數據的 CSV 檔案，系統將自動繪製 Pressure 與 Leak 的箱線分佈圖。")
    
    uploaded_file_box = st.file_uploader("請上傳統計數據 CSV 檔案 (Box Plot)", type=["csv"], key="box_uploader")
    
    if uploaded_file_box is not None:
        try:
            # 讀取數據
            df_box = pd.read_csv(uploaded_file_box)
            
            # 檢查基本必要欄位
            required_box_cols = ['SN', 'Pressure(Kpa)', 'Leak']
            missing_box_cols = [col for col in required_box_cols if col not in df_box.columns]
            
            if missing_box_cols:
                st.error(f"❌ 檔案格式不符，缺少必要欄位: {missing_box_cols}")
            else:
                st.success(f"✅ 成功讀取統計數據！共有 {len(df_box)} 筆產品總點數。")
                
                # 數據摘要摘要
                if 'bResult' in df_box.columns:
                    ng_count = len(df_box[df_box['bResult'] == 'NG'])
                    ok_count = len(df_box[df_box['bResult'] == 'OK'])
                    col_b1, col_b2 = st.columns(2)
                    col_b1.metric("🟢 OK 總數", f"{ok_count} 筆")
                    col_b2.metric("🔴 NG 總數", f"{ng_count} 筆")
                
                st.write("---")
                
                # 建立左右兩欄，同時呈現 Pressure 和 Leak 的 Box plot
                graph_col1, graph_col2 = st.columns(2)
                
                with graph_col1:
                    st.markdown("#### 📦 Pressure (Kpa) 箱線分佈圖")
                    # 使用 Plotly Express 繪製精美箱線圖，並把個別 SN 點散佈在旁邊方便抓異常
                    fig_box_press = px.box(
                        df_box, 
                        y="Pressure(Kpa)", 
                        points="all",  # 顯示所有資料點
                        hover_data=["SN", "bResult"] if "bResult" in df_box.columns else ["SN"],
                        color="bResult" if "bResult" in df_box.columns else None,
                        color_discrete_map={"OK": "#1f77b4", "NG": "#ef553b"}
                    )
                    fig_box_press.update_layout(margin=dict(t=20, b=20))
                    st.plotly_chart(fig_box_press, use_container_width=True)
                    
                with graph_col2:
                    st.markdown("#### 📦 Leak 箱線分佈圖")
                    fig_box_leak = px.box(
                        df_box, 
                        y="Leak", 
                        points="all", 
                        hover_data=["SN", "bResult"] if "bResult" in df_box.columns else ["SN"],
                        color="bResult" if "bResult" in df_box.columns else None,
                        color_discrete_map={"OK": "#1f77b4", "NG": "#ef553b"}
                    )
                    fig_box_leak.update_layout(margin=dict(t=20, b=20))
                    st.plotly_chart(fig_box_leak, use_container_width=True)
                
                # 原始數據檢視器
                with st.expander("查看完整 Box Plot 原始數據清單"):
                    st.dataframe(df_box)
                    
        except Exception as e:
            st.error(f"讀取 Box Plot 檔案時發生錯誤: {e}")
