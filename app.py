
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# 1. 設定網頁標題與配置
st.set_page_config(page_title="氣密測試數據自動化分析工具", layout="wide")
st.title("📊 氣密測試 (Leak Test) 數據分析系統")

# 在側邊欄放一個全域的控制項，讓使用者決定要不要對大數據量進行抽樣檢視
st.sidebar.header("⚙️ 圖表控制面板")
sample_step = st.sidebar.slider(
    "📊 數據抽樣間隔 (Step)", 
    min_value=1, 
    max_value=10, 
    value=1, 
    help="1 代表顯示所有原始數據；2 代表每 2 筆抽 1 筆顯示。若圖表卡頓或點太密集，建議調大此數值。"
)

st.write("請在下方上傳您的 CSV 檔案，系統將自動過濾 Stabilization 數據、依據 SN 分頁並繪製 Pressure 與 Leak 趨勢圖。")

# 2. 檔案上傳元件
uploaded_file = st.file_uploader("選擇上傳 CSV 檔案", type=["csv"])

if uploaded_file is not None:
    try:
        # 讀取 CSV 數據
        df = pd.read_csv(uploaded_file)
        
        # 檢查必要欄位是否存在
        required_columns = ['SN', 'Timestamp', 'Pressure(Kpa)', 'Leak']
        missing_cols = [col for col in required_columns if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ 檔案格式不符，缺少必要欄位: {missing_cols}")
        else:
            # 確保時間欄位格式正確
            df['Timestamp'] = pd.to_datetime(df['Timestamp'])
            
            # 【核心修改：新增條件過濾】
            # 自動忽略 Phase 欄位中包含 'Stabilization' 的所有數據（不區分大小寫）
            if 'Phase' in df.columns:
                initial_count = len(df)
                df = df[~df['Phase'].astype(str).str.contains('Stabilization', case=False, na=False)]
                filtered_count = initial_count - len(df)
                st.sidebar.filter_info = st.sidebar.warning(f"🧹 已自動過濾掉 {filtered_count} 筆 Stabilization 階段的數據。")
            
            # 排序讓折線圖順暢
            df = df.sort_values(by='Timestamp')
            
            # 獲取所有獨特的 SN 清單
            unique_sns = df['SN'].dropna().unique()
            st.success(f"✅ 成功讀取數據！偵測到共 {len(unique_sns)} 個不同的 SN 產品。")
            
            # 3. 動態建立分頁 (Tabs)
            tabs = st.tabs([f"SN: {sn.split(':')[-1] if ':' in str(sn) else sn}" for sn in unique_sns])
            
            # 4. 迭代每個分頁，填充數據與圖表
            for i, sn in enumerate(unique_sns):
                with tabs[i]:
                    st.subheader(f"產品序號: {sn}")
                    
                    # 篩選出該 SN 的數據
                    sn_data = df[df['SN'] == sn]
                    
                    if sn_data.empty:
                        st.info("ℹ️ 該產品在排除 Stabilization 階段後無剩餘數據。")
                        continue
                    
                    # 根據側邊欄的設定進行數據抽樣（減少點的重疊）
                    if sample_step > 1:
                        plot_data = sn_data.iloc[::sample_step]
                        st.info(f"💡 當前圖表已啟用抽樣（每 {sample_step} 筆取 1 筆），目前繪圖點數：{len(plot_data)} 筆。")
                    else:
                        plot_data = sn_data
                    
                    # 顯示數據基本摘要
                    col1, col2, col3 = st.columns(3)
                    col1.metric("有效分析數據筆數", f"{len(sn_data)} 筆")
                    col2.metric("最大壓力 (Kpa)", f"{sn_data['Pressure(Kpa)'].max():.2f}")
                    col3.metric("測試結果狀態", f"{sn_data['bResult'].iloc[-1] if 'bResult' in sn_data.columns else '未知'}")
                    
                    st.write("---") # 分隔線
                    
                    # ======= 第一張圖：Pressure 趨勢圖 =======
                    st.markdown("#### 📈 Pressure (Kpa) 趨勢變化圖")
                    fig_press = go.Figure()
                    fig_press.add_trace(go.Scatter(
                        x=plot_data['Timestamp'],
                        y=plot_data['Pressure(Kpa)'],
                        name='Pressure (Kpa)',
                        mode='lines+markers',
                        line=dict(color='#1f77b4', width=2),
                        marker=dict(size=5)
                    ))
                    fig_press.update_layout(
                        xaxis=dict(title="時間軸 (Timestamp)"),
                        yaxis=dict(title="Pressure (Kpa)"),
                        hovermode="closest",
                        margin=dict(t=30, b=30)
                    )
                    st.plotly_chart(fig_press, use_container_width=True)
                    
                    # ======= 第二張圖：Leak 趨勢圖 (若數據不為全空才顯示) =======
                    st.markdown("#### 📉 Leak 趨勢變化圖")
                    if not sn_data['Leak'].dropna().empty:
                        fig_leak = go.Figure()
                        fig_leak.add_trace(go.Scatter(
                            x=plot_data['Timestamp'],
                            y=plot_data['Leak'],
                            name='Leak',
                            mode='lines+markers',
                            line=dict(color='#ff7f0e', width=2),
                            marker=dict(size=5)
                        ))
                        fig_leak.update_layout(
                            xaxis=dict(title="時間軸 (Timestamp)"),
                            yaxis=dict(title="Leak"),
                            hovermode="closest",
                            margin=dict(t=30, b=30)
                        )
                        st.plotly_chart(fig_leak, use_container_width=True)
                    else:
                        st.info("ℹ️ 此 SN 當前數據中的 Leak 欄位皆為空值，暫不顯示圖表。")
                    
                    # 在圖表下方提供該 SN 的原始數據表格查看
                    with st.expander("查看此 SN 原始數據內容 (已排除 Stabilization)"):
                        st.dataframe(sn_data)
                        
    except Exception as e:
        st.error(f"讀取檔案時發生錯誤: {e}")
