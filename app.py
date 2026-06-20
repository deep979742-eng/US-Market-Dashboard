import streamlit as st
import pandas as pd
import datetime
import time
import os
import json
import base64
import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials
import concurrent.futures
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_autorefresh import st_autorefresh

# ==========================================
# 1. PAGE SETUP & CSS
# ==========================================
st.set_page_config(page_title="US F&O Dashboard", layout="wide")

css_str = """<style>
[data-testid='stAppViewContainer'], [data-testid='stAppViewBlockContainer'], [data-testid='stHeader'], [data-testid='stSidebar'], .stApp, .stApp > div { opacity: 1 !important; filter: none !important; transition: none !important; } 
[data-testid='stDataFrame'], [data-testid='stTabs'] { opacity: 1 !important; filter: none !important; transition: none !important; } 
[data-testid='stStatusWidget'] { visibility: hidden !important; display: none !important; } 

.block-container { padding-top: 3.5rem !important; padding-bottom: 0rem !important; padding-left: 1rem !important; padding-right: 1rem !important; } 
[data-testid='stDataFrameTable'] > thead > tr { background-color: darkblue !important; } 

[data-testid='stDataFrameTable'] > thead > tr > th { 
    background-color: darkblue !important; 
    color: white !important; 
    font-weight: bold !important; 
    text-align: center !important; 
    writing-mode: vertical-rl !important; 
    transform: rotate(180deg) !important; 
    white-space: nowrap !important; 
    padding: 8px 4px !important;
    height: 120px !important;
} 

th { background-color: darkblue !important; color: white !important; } 
* { cursor: default !important; } 

@media (max-width: 768px) { 
    .block-container { padding-top: 4rem !important; padding-left: 0.1rem !important; padding-right: 0.1rem !important; } 
    [data-testid='stDataFrameTable'] th { font-size: 10px !important; height: 100px !important; padding: 4px 2px !important; } 
    [data-testid='stDataFrameTable'] td { font-size: 10px !important; padding: 4px 2px !important; } 
}
</style>"""
st.markdown(css_str, unsafe_allow_html=True)

IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))
now_ist = datetime.datetime.now(IST)
today_str = now_ist.strftime("%Y-%m-%d")

HISTORY_FILE = "yf_chart_history.csv"
SNAPSHOT_FILE = "yf_snapshot_8pm.json" 
AUTO_SAVE_FILE = "yf_auto_save_tracker.txt"

if 'live_base_date' not in st.session_state or st.session_state.live_base_date != today_str:
    st.session_state.live_base = {}
    st.session_state.live_base_date = today_str

# ==========================================
# 2. GOOGLE SHEETS DYNAMIC CONNECTION
# ==========================================
@st.cache_resource
def get_gspread_client():
    try:
        if "gcp_service_account" in st.secrets:
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            return gspread.authorize(creds)
    except: pass
    return None

# ==========================================
# 3. STOCK LIST & HELPER FUNCTIONS
# ==========================================
raw_symbols = [
    "AAPL", "MSFT", "TSLA", "NVDA", "AMZN", "META", "GOOGL", "NFLX", "AMD", "SPY", "QQQ"
]

def calc_vol_pcr(ce_vol, pe_vol): return 0.0 if ce_vol == 0 else round(pe_vol / ce_vol, 2)
def calc_opt_pcr(ce_oi, pe_oi): return 0.0 if ce_oi == 0 else round(pe_oi / ce_oi, 2)
def calc_vol_cpr(ce_vol, pe_vol): return 0.0 if pe_vol == 0 else round(ce_vol / pe_vol, 2)

def get_generic_key(sym):
    try:
        import re
        match = re.match(r"([A-Z]+)(\d{6})([CP])(\d+)", sym)
        if match:
            ticker, date, opt_type, strike_raw = match.groups()
            strike = str(float(strike_raw) / 1000)
            return f"{ticker}_{strike}{opt_type}"
    except: pass
    return None

# ==========================================
# 4. MASTER SCANNER (YAHOO FINANCE)
# ==========================================
@st.cache_data(ttl=290, show_spinner=False)
def run_master_scan(date_str):
    scan_time_ist = datetime.datetime.now(IST)
    time_str = scan_time_ist.strftime('%H:%M')
    
    baseline_prices = {}
    baseline_generic = {} 
    snap_time = {}
    snapshot_changed = False
    
    client = get_gspread_client()
    if client:
        try:
            ss = client.open("US_F&O_Data")
            ws1 = ss.get_worksheet(0) 
            ws2 = ss.worksheet("Sheet2") 
            
            try:
                tab2_date_row = ws2.cell(1, 1).value
                if tab2_date_row and "LAST_SAVED_DATE:" in tab2_date_row:
                    saved_date = tab2_date_row.replace("LAST_SAVED_DATE:", "").strip()
                    if saved_date and saved_date != date_str:
                        tab2_col_vals = ws2.col_values(1)[1:] 
                        full_b64 = "".join(tab2_col_vals)
                        if full_b64:
                            ws1.clear()
                            chunks = [full_b64[i:i+40000] for i in range(0, len(full_b64), 40000)]
                            clist1 = ws1.range(f'A1:A{len(chunks)}')
                            for i, cell in enumerate(clist1): cell.value = chunks[i]
                            ws1.update_cells(clist1)
                            
                            ws2.update_cell(1, 1, f"LAST_SAVED_DATE: {date_str}")
                            ws2.batch_clear(["A2:A100"])
            except: pass

            try:
                col_vals = ws1.col_values(1)
                if col_vals:
                    full_str = "".join(col_vals)
                    decoded_str = base64.b64decode(full_str).decode('utf-8')
                    loaded_prices = json.loads(decoded_str)
                    for k, v in loaded_prices.items():
                        baseline_prices[k] = round(float(v), 2)
                        gen_key = get_generic_key(k)
                        if gen_key:
                            baseline_generic[gen_key] = round(float(v), 2)
            except: pass

            try:
                snap_val = ws2.cell(1, 2).value
                if snap_val: snap_time = json.loads(snap_val)
            except: pass
        except: pass

    st.session_state.baseline_count = len(baseline_prices)
    st.session_state.has_snapshot = bool(snap_time)
        
    final_list = []
    new_csv_rows = []
    live_ltp_data = {} 

    # 🚀 YFINANCE FETCH LOGIC (WITH ADVANCED ERROR CATCHING)
    def fetch_yf_data(sym):
        for attempt in range(3):
            try:
                tk = yf.Ticker(sym)
                hist = tk.history(period="5d")
                
                if hist.empty: 
                    return sym, "Hist Data Empty", None
                
                spot_ltp = float(hist['Close'].iloc[-1])
                open_p = float(hist['Open'].iloc[-1])
                float_c = float(hist['Close'].iloc[-2]) if len(hist) > 1 else open_p
                ltp_ch = spot_ltp - float_c
                chg_pct = (ltp_ch / float_c) * 100 if float_c != 0 else 0

                exps = tk.options
                if not exps: 
                    return sym, "No Options Found", None
                
                opt = tk.option_chain(exps[0])
                return sym, {
                    'spot_ltp': spot_ltp, 'open_p': open_p, 'float_c': float_c,
                    'ltp_ch': ltp_ch, 'chg_pct': chg_pct
                }, opt
            except Exception as e:
                time.sleep(2) # Agar block hua toh 2 second rukega
                if attempt == 2:
                    return sym, f"Err: {str(e)[:15]}", None
                    
        return sym, "Unknown Error", None

    # 🚀 SPEED LIMIT APPLIED (max_workers=2) TO PREVENT BLOCK
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        results = executor.map(fetch_yf_data, raw_symbols)
        for s_name, spot_data, oc in results:
            
            # Error aane par yahan capture hoga
            if spot_data is None or isinstance(spot_data, str) or not oc:
                err_msg = spot_data if isinstance(spot_data, str) else "NA"
                final_list.append({'SYMS': f"{s_name} ({err_msg})", 'OPEN_STATUS': "NA", 'V_PCR': 0.0, 'O_PCR': 0.0, 'V_CPR': 0.0, 'LTP_CH': 0.0, 'CHG_%': 0.0, 'LTP': 0.0, 'VOL_ABS': 0.0, 'PCR_ABS': 0.0, 'VOL_PCT': 0.0, 'PCR_PCT': 0.0, 'CE_CON': 0.0, 'PE_CON': 0.0})
                continue
                
            open_p = spot_data['open_p']
            float_c = spot_data['float_c']
            spot_ltp = spot_data['spot_ltp']
            open_status = "NA" if open_p == 0 or float_c == 0 else "Gap Up 🔼" if open_p > float_c else "Gap Down 🔽" if open_p < float_c else "Same ➖"

            calls = oc.calls
            puts = oc.puts
            
            c_oi = float(calls['openInterest'].sum() if 'openInterest' in calls else 0)
            p_oi = float(puts['openInterest'].sum() if 'openInterest' in puts else 0)
            c_v = float(calls['volume'].sum() if 'volume' in calls else 0)
            p_v = float(puts['volume'].sum() if 'volume' in puts else 0)
            
            for _, row in calls.iterrows():
                sym_str = str(row.get('contractSymbol', ''))
                lp_str = round(float(row.get('lastPrice', 0)), 2)
                if lp_str > 0: live_ltp_data[sym_str] = lp_str

            for _, row in puts.iterrows():
                sym_str = str(row.get('contractSymbol', ''))
                lp_str = round(float(row.get('lastPrice', 0)), 2)
                if lp_str > 0: live_ltp_data[sym_str] = lp_str

            o_pcr = calc_opt_pcr(c_oi, p_oi)
            v_cpr = calc_vol_cpr(c_v, p_v)
            v_pcr = calc_vol_pcr(c_v, p_v)

            if scan_time_ist.time() < datetime.time(20, 0):
                pcr_abs, vol_abs, pcr_pct, vol_pct = 0.0, 0.0, 0.0, 0.0
            else:
                if s_name not in snap_time:
                    snap_time[s_name] = {'pcr': o_pcr, 'vol_cpr': v_cpr}
                    snapshot_changed = True
                    pcr_abs, vol_abs, pcr_pct, vol_pct = 0.0, 0.0, 0.0, 0.0
                else:
                    base = snap_time[s_name]
                    base_pcr_val = base['pcr']
                    base_vol_val = base['vol_cpr']
                    
                    def get_standard_pct(current_val, base_val):
                        if base_val == 0: return 0.0
                        return ((current_val - base_val) / base_val) * 100.0
                        
                    pcr_abs = o_pcr - base_pcr_val
                    vol_abs = v_cpr - base_vol_val
                    pcr_pct = get_standard_pct(o_pcr, base_pcr_val)
                    vol_pct = get_standard_pct(v_cpr, base_vol_val)

            def get_conv(opt_df):
                if opt_df.empty: return 0.0
                tot_p, tot_m = 0, 0
                for _, row in opt_df.iterrows():
                    sym = str(row.get('contractSymbol', ''))
                    lp = round(float(row.get('lastPrice', 0)), 2)
                    if lp == 0: continue
                    
                    diff = 0.0
                    if sym in baseline_prices:
                        diff = round(lp - baseline_prices[sym], 2)
                    else:
                        gen_key = get_generic_key(sym)
                        if gen_key and gen_key in baseline_generic:
                            diff = round(lp - baseline_generic[gen_key], 2)
                            
                    if diff > 0.00: tot_p += 1 
                    elif diff < 0.00: tot_m += 1 

                act = tot_p + tot_m
                if act == 0: return 0.0
                return round((tot_p / act) * 100, 2) if tot_p >= tot_m else -round((tot_m / act) * 100, 2)
            
            final_list.append({
                'SYMS': s_name, 'OPEN_STATUS': open_status, 'V_PCR': v_pcr, 'O_PCR': o_pcr, 'V_CPR': v_cpr, 
                'LTP_CH': spot_data['ltp_ch'], 'CHG_%': spot_data['chg_pct'], 'LTP': spot_ltp,
                'VOL_ABS': round(vol_abs, 2), 'PCR_ABS': round(pcr_abs, 2), 
                'VOL_PCT': round(vol_pct, 2), 'PCR_PCT': round(pcr_pct, 2),
                'CE_CON': get_conv(calls), 'PE_CON': get_conv(puts)
            })

            new_csv_rows.append({'Date': date_str, 'Symbol': s_name, 'Time': time_str, 'LTP': spot_ltp, 'VOL PCR': v_pcr, 'OPT PCR': o_pcr, 'VOL CPR': v_cpr})

    if snapshot_changed and client:
        try:
            ss = client.open("US_F&O_Data")
            ws2 = ss.worksheet("Sheet2")
            ws2.update_cell(1, 2, json.dumps(snap_time))
        except: pass

    st.session_state.get_live_dump = live_ltp_data

    if new_csv_rows:
        new_df = pd.DataFrame(new_csv_rows)[['Date', 'Symbol', 'Time', 'LTP', 'VOL PCR', 'OPT PCR', 'VOL CPR']]
        if not os.path.isfile(HISTORY_FILE): new_df.to_csv(HISTORY_FILE, index=False)
        else: new_df.to_csv(HISTORY_FILE, mode='a', header=False, index=False)

    return final_list, scan_time_ist.timestamp()

# ==========================================
# 5. SIDEBAR & EOD SAVE
# ==========================================
st.sidebar.header("📊 Yahoo Finance Status")
st.sidebar.success("🟢 Connected to Market Data")
st.sidebar.markdown("---")
st.sidebar.header("💾 End Of Day (EOD) Save")

def save_eod_data():
    if 'get_live_dump' not in st.session_state:
        st.sidebar.error("⚠️ Error: Live Data abhi fetch nahi hua.")
        return False
        
    live_data = st.session_state.get_live_dump
    if not live_data:
        st.sidebar.error("⚠️ Error: Data khali hai! API data nahi de raha.")
        return False
        
    try:
        client = get_gspread_client()
        if not client: 
            st.sidebar.error("⚠️ Error: Google Sheet connect nahi hui.")
            return False
            
        ss = client.open("US_F&O_Data")
        ws2 = ss.worksheet("Sheet2")
        
        locked_live_data = {k: round(float(v), 2) for k, v in live_data.items()}
        json_str = json.dumps(locked_live_data)
        b64_str = base64.b64encode(json_str.encode('utf-8')).decode('utf-8')
        chunks = [b64_str[i:i+40000] for i in range(0, len(b64_str), 40000)]
        
        ws2.batch_clear(["A2:A100"])
        clist2 = ws2.range(f'A2:A{len(chunks)+1}')
        for i, cell in enumerate(clist2): cell.value = chunks[i]
        
        ws2.update_cell(1, 1, f"LAST_SAVED_DATE: {today_str}")
        ws2.update_cells(clist2)
        return True
    except Exception as e:
        st.sidebar.error(f"⚠️ Error: {e}")
        return False

if st.sidebar.button("Manual Save Data"):
    if save_eod_data(): 
        st.sidebar.success("✅ Sheet Saved Successfully!")

# ==========================================
# 6. APP RENDERING & MAGIC VIEWER
# ==========================================
cached_result, last_scan_timestamp = run_master_scan(today_str)

if cached_result is not None:
    st.session_state.cached_data = cached_result
    st.session_state.last_api_call = datetime.datetime.fromtimestamp(last_scan_timestamp, IST)
    
    if datetime.time(18, 0) <= now_ist.time() < datetime.time(19, 0):
        last_save = open(AUTO_SAVE_FILE, "r").read().strip() if os.path.exists(AUTO_SAVE_FILE) else ""
        if last_save != today_str:
            if save_eod_data(): open(AUTO_SAVE_FILE, "w").write(today_str)
else:
    if 'cached_data' not in st.session_state: st.session_state.cached_data = []

if len(st.session_state.cached_data) > 0:
    
    def style_indicators(val):
        if isinstance(val, str): 
            if "Gap Up" in val: return 'color: #00AA00; font-weight: bold; text-align: center;'
            if "Gap Down" in val: return 'color: #FF0000; font-weight: bold; text-align: center;'
            if "Same" in val: return 'color: #00BFFF; font-weight: bold; text-align: center;'
            return 'text-align: center;'
        if val > 0: return 'color: #00AA00; font-weight: bold; text-align: center;'
        elif val < 0: return 'color: #FF0000; font-weight: bold; text-align: center;'
        return 'color: #888888; font-weight: bold; text-align: center;'

    def style_pcr_columns(val):
        if isinstance(val, (int, float)):
            if val >= 1.0: return 'color: #00AA00; font-weight: bold; text-align: center;'
            elif val > 0 and val < 1.0: return 'color: #FF0000; font-weight: bold; text-align: center;'
        return 'text-align: center;'

    header_styles = [
        {'selector': 'th', 'props': [('background-color', 'darkblue'), ('color', 'white'), ('font-weight', 'bold'), ('text-align', 'center')]},
        {'selector': 'thead th', 'props': [('background-color', 'darkblue'), ('color', 'white'), ('font-weight', 'bold'), ('text-align', 'center')]}
    ]

    tab1, tab2 = st.tabs(["📊 Dashboard", "📈 TREND CHART"])
    
    with tab1:
        show_pct = st.toggle("📊 Show Checker Data in Percentage (%)", value=True)
        
        checker_fmt = '{:+.2f}%' if show_pct else '{:+.2f}'
        format_dict = {'VOL PCR': '{:.2f}', 'OPTION PCR': '{:.2f}', 'VOL CPR': '{:.2f}', 'LTP': '{:.2f}', 'LTP CHANGE': '{:.2f}', 'CHANGE%': '{:+.2f}%', 'VOL CHECKER': checker_fmt, 'PCR CHECKER': checker_fmt, 'CE_CONTRACT': '{:+.2f}%', 'PE_CONTRACT': '{:+.2f}%'}
        
        df = pd.DataFrame(st.session_state.cached_data)
        
        if not df.empty:
            df['Conv_Rank'] = df['CE_CON'].abs() + df['PE_CON'].abs()
            df = df.sort_values(by='Conv_Rank', ascending=False).drop(columns=['Conv_Rank']) 
            df['VOL CHECKER'] = df['VOL_PCT'] if show_pct else df['VOL_ABS']
            df['PCR CHECKER'] = df['PCR_PCT'] if show_pct else df['PCR_ABS']
            df = df.drop(columns=['VOL_ABS', 'PCR_ABS', 'VOL_PCT', 'PCR_PCT'])
            df = df.rename(columns={'SYMS': 'SYMBOL', 'OPEN_STATUS': 'OPENING', 'V_PCR': 'VOL PCR', 'O_PCR': 'OPTION PCR', 'V_CPR': 'VOL CPR', 'LTP_CH': 'LTP CHANGE', 'CHG_%': 'CHANGE%', 'LTP': 'LTP', 'CE_CON': 'CE_CONTRACT', 'PE_CON': 'PE_CONTRACT'})

            styled_df = (df.style.set_properties(**{'text-align': 'center'}).format(format_dict).set_table_styles(header_styles)
                         .map(style_indicators, subset=['OPENING', 'LTP CHANGE', 'CHANGE%', 'CE_CONTRACT', 'PE_CONTRACT', 'VOL CHECKER', 'PCR CHECKER'])
                         .map(style_pcr_columns, subset=['VOL PCR', 'OPTION PCR', 'VOL CPR']))

            st.dataframe(styled_df, use_container_width=True, height=800, hide_index=True)

    with tab2:
        st.markdown("### 📈 TREND CHART") 
        col_c1, col_c2 = st.columns([2, 2])
        with col_c1: sel_stock = st.selectbox("Select Stock for Trend:", raw_symbols, index=0, key="c_stock")
        with col_c2: 
            chart_mode = st.radio("SWITCH CHART VIEW:", ["Vol CPR", "Option PCR"], horizontal=True)

        if os.path.exists(HISTORY_FILE):
            try:
                hist_df = pd.read_csv(HISTORY_FILE)
                if not hist_df.empty and 'Date' in hist_df.columns:
                    df_sym = hist_df[(hist_df['Date'] == today_str) & (hist_df['Symbol'] == sel_stock)].copy()
                    if not df_sym.empty:
                        df_sym = df_sym.sort_values(by='Time')
                        df_sym['Datetime'] = pd.to_datetime(df_sym['Date'] + ' ' + df_sym['Time'])
                        
                        target_col = 'VOL CPR' if chart_mode == "Vol CPR" else 'OPT PCR'
                        line_color = "#FF4D4D" if chart_mode == "Vol CPR" else "#00BFFF" 
                        
                        fig = make_subplots(specs=[[{"secondary_y": True}]])
                        fig.add_trace(go.Scatter(x=df_sym['Datetime'], y=df_sym[target_col], name=f"{chart_mode}", line=dict(color=line_color, width=3, shape="spline"), mode="lines"), secondary_y=False)
                        fig.add_trace(go.Scatter(x=df_sym['Datetime'], y=df_sym['LTP'], name="Stock LTP", line=dict(color="#00CC66", width=3, shape="spline"), mode="lines"), secondary_y=True)

                        fig.update_layout(
                            template="plotly_white", hovermode="x unified", height=380, margin=dict(l=10, r=10, t=40, b=10), 
                            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", 
                            xaxis=dict(rangeslider=dict(visible=False), type="date", gridcolor="#E5E5E5", color="black"),
                            yaxis=dict(title=dict(text=f"{chart_mode} Scale", font=dict(color=line_color)), tickfont=dict(color=line_color), gridcolor="#E5E5E5", autorange=True),
                            yaxis2=dict(title=dict(text="LTP Price Scale", font=dict(color="#00CC66")), tickfont=dict(color="#00CC66"), showgrid=False, autorange=True),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(color="black"))
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else: st.info(f"⏳ Waiting for Market Data for {sel_stock}.")
                else: st.info("⏳ Market data hasn't started logging yet today.")
            except Exception as e: st.error(f"Chart Load Error: {e}")
        else: st.info("⏳ Chart History file is being prepared...")
else:
    st.error("⚠️ Data fetch me error aayi. Kripya check karein ki internet connection theek hai aur yfinance library install hai.")

# ==========================================
# 7. EXACT BOUNDARY AUTO-REFRESH LOGIC 🎯
# ==========================================
now_refresh = datetime.datetime.now(IST)
current_total_secs = now_refresh.minute * 60 + now_refresh.second

targets = [(m * 60 + 5) for m in range(0, 65, 5)]
secs_wait = 300
for t in targets:
    if t > current_total_secs:
        secs_wait = t - current_total_secs
        break

if secs_wait < 5: secs_wait = 5
st_autorefresh(interval=secs_wait * 1000, key="exact_boundary_timer")
