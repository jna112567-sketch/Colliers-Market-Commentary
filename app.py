import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import urllib.request
import xml.etree.ElementTree as ET
import os
import json
import urllib.parse
import plotly.io as pio
import requests
import time
from google import genai
from google.genai import types

def display_df_with_changes(df, is_percent=False):
    if df is None or df.empty or 'Quarter' not in df.columns:
        st.dataframe(df, width="stretch", hide_index=True)
        return

    df_calc = df.copy()
    numeric_cols = df_calc.select_dtypes(include=['number']).columns
    
    if is_percent:
        # Calculate Percentage Point Change: (New - Old) * 100
        df_qoq = df_calc[numeric_cols].diff() * 100
        df_yoy = df_calc[numeric_cols].diff(periods=4) * 100
        
        # Original data gets % format, changes get percentage points (%p)
        val_format = "{:.2%}" 
        change_format = "{:.2f} %p" 
    else:
        # Standard Percentage Change: (New - Old) / Old
        df_qoq = df_calc[numeric_cols].pct_change()
        df_yoy = df_calc[numeric_cols].pct_change(periods=4)
        
        # Original data gets commas, changes get converted from decimal to %
        val_format = "{:,.0f}"
        change_format = "{:.2%}" # Automatically multiplies by 100 and adds the % sign

    def color_negative_red(val):
        if isinstance(val, (int, float)) and pd.notna(val):
            color = 'red' if val < 0 else 'black'
            return f'color: {color}'
        return ''

    sub_tabs = st.tabs(["Original Data", "QoQ Change", "YoY Change"])
    
    with sub_tabs[0]:
        # Applied formatting to the original data table
        st.dataframe(df.style.format(val_format, subset=numeric_cols).map(color_negative_red, subset=numeric_cols), width="stretch", hide_index=True)
    
    with sub_tabs[1]:
        # Applied the new dynamic change_format
        st.dataframe(df_qoq.style.format(change_format, na_rep="-").map(color_negative_red), width="stretch", hide_index=True)
        
    with sub_tabs[2]:
        st.dataframe(df_yoy.style.format(change_format, na_rep="-").map(color_negative_red), width="stretch", hide_index=True)
def display_latest_metrics(df, title="Latest Quarter", format_type="number"):
    if df is None or df.empty or len(df) < 2:
        return
        
    latest_qtr = df['Quarter'].iloc[-1]
    
    with st.container(border=True):
        st.markdown(f"#### 📌 {title} Headline (as of {latest_qtr})")
        
        target_cols = [c for c in ["Overall", "CBD", "GBD", "YBD"] if c in df.columns]
        if not target_cols: return
            
        cols = st.columns(len(target_cols))
        
        for i, col in enumerate(target_cols):
            try:
                val_latest = float(df[col].iloc[-1])
                val_prev = float(df[col].iloc[-2])
            except:
                continue
                
            # Explicit formatting based on the type of metric
            if format_type == "percent":
                val_str = f"{val_latest:.2%}" # Forces 2 decimal percentage
                delta_val = (val_latest - val_prev) * 100
                delta_str = f"{delta_val:.2f} %p" # Percentage point change
            else:
                val_str = f"{val_latest:,.0f}"
                delta_val = (val_latest - val_prev) / abs(val_prev) if val_prev != 0 else 0
                delta_str = f"{delta_val:.2%}" 
                
            cols[i].metric(label=col, value=val_str, delta=delta_str)
API_KEY = st.secrets.get("ECOS_API_KEY")

@st.cache_data(ttl=3600)
def _get_ecos(table, cycle, start, end, i1, i2="?", i3="?"):
    if not API_KEY: return []
    url = f"http://ecos.bok.or.kr/api/StatisticSearch/{API_KEY}/json/kr/1/100/{table}/{cycle}/{start}/{end}/{i1}/{i2}/{i3}"
    try:
        res = requests.get(url, timeout=10).json()
        return res.get('StatisticSearch', {}).get('row', [])
    except:
        return []

@st.cache_data(ttl=3600)
def fetch_macro_core():
    start_q, end_q = "2018Q1", "2026Q4"
    start_m, end_m = "201801", "202612"
    
    # GDP
    df_gdp_raw = _get_ecos("200Y104", "Q", start_q, end_q, "1400")
    gdp_data = [{"Quarter": r["TIME"], "Real GDP (KRW Tr)": float(r["DATA_VALUE"]) / 1000} for r in df_gdp_raw]
    df_gdp = pd.DataFrame(gdp_data)
    if not df_gdp.empty:
        df_gdp["Real GDP Growth (YoY %)"] = (df_gdp["Real GDP (KRW Tr)"] / df_gdp["Real GDP (KRW Tr)"].shift(4) - 1) * 100
        df_gdp["Real GDP Growth (QoQ %)"] = (df_gdp["Real GDP (KRW Tr)"] / df_gdp["Real GDP (KRW Tr)"].shift(1) - 1) * 100
        
    # Trade
    df_exp_m = _get_ecos("901Y118", "M", start_m, end_m, "T002")
    df_imp_m = _get_ecos("901Y118", "M", start_m, end_m, "T004")
    trade_data = []
    for r in df_exp_m:
        q_str = r['TIME'][:4] + "Q" + str((int(r['TIME'][4:])-1)//3 + 1)
        trade_data.append({"Quarter": q_str, "Type": "Export (USD Bn)", "Val": float(r["DATA_VALUE"]) / 1000}) 
    for r in df_imp_m:
        q_str = r['TIME'][:4] + "Q" + str((int(r['TIME'][4:])-1)//3 + 1)
        trade_data.append({"Quarter": q_str, "Type": "Import (USD Bn)", "Val": float(r["DATA_VALUE"]) / 1000})
    df_trade = pd.DataFrame(trade_data)
    if not df_trade.empty:
        df_trade = df_trade.groupby(["Quarter", "Type"]).sum().reset_index().pivot(index="Quarter", columns="Type", values="Val").reset_index()
        if "Export (USD Bn)" in df_trade.columns and "Import (USD Bn)" in df_trade.columns:
            df_trade["Trade Balance (USD Bn)"] = df_trade["Export (USD Bn)"] - df_trade["Import (USD Bn)"]
        
    # CPI
    cpi_codes = {"0": "Total CPI", "A": "Food & Beverages", "B": "Alcohol & Tobacco", "C": "Clothing & Footwear",
                 "D": "Housing & Utilities", "E": "Furnishings", "F": "Health", "G": "Transport",
                 "H": "Communication", "I": "Recreation/Culture", "J": "Education", "K": "Restaurants/Hotels", "L": "Misc Goods/Services"}
    cpi_data = []
    for code, name in cpi_codes.items():
        res = _get_ecos("901Y009", "Q", start_q, end_q, code)
        for r in res:
            cpi_data.append({"Quarter": r["TIME"], "Metric": name, "Val": float(r["DATA_VALUE"])})
    df_cpi = pd.DataFrame(cpi_data)
    if not df_cpi.empty:
        df_cpi = df_cpi.pivot(index="Quarter", columns="Metric", values="Val").reset_index()
        if "Total CPI" in df_cpi.columns:
            df_cpi["Total CPI Growth (YoY %)"] = (df_cpi["Total CPI"] / df_cpi["Total CPI"].shift(4) - 1) * 100

    return df_gdp, df_trade, df_cpi

@st.cache_data(ttl=3600)
def fetch_macro_empl_forex():
    start_q, end_q = "2018Q1", "2026Q4"
    
    empl_data = []
    for r in _get_ecos("901Y027", "Q", start_q, end_q, "I61BA"):
        if "원계열" in r.get("ITEM_NAME2", "") or "원계열" in r.get("ITEM_NAME1", "") or "계절" not in str(r.values()):
             empl_data.append({"Quarter": r["TIME"], "Employed Pop (000s)": float(r["DATA_VALUE"])})
    df_empl1 = pd.DataFrame(empl_data).drop_duplicates(subset=["Quarter"])
    
    unemp_data = []
    for r in _get_ecos("901Y027", "Q", start_q, end_q, "I61BC/I28A"):
        unemp_data.append({"Quarter": r["TIME"], "Unemployment Rate (%)": float(r["DATA_VALUE"])})
    df_empl2 = pd.DataFrame(unemp_data).drop_duplicates(subset=["Quarter"])
    
    df_empl = pd.DataFrame()
    if not df_empl1.empty:
        df_empl = df_empl1
        if not df_empl2.empty:
            df_empl = df_empl.merge(df_empl2, on="Quarter", how="outer")
            
    forex_data = []
    # USD
    for r in _get_ecos("731Y006", "Q", start_q, end_q, "0000003", "0000100"): 
        forex_data.append({"Quarter": r["TIME"], "KRW/USD (Avg)": float(r["DATA_VALUE"])})
    # JPY
    for r in _get_ecos("731Y006", "Q", start_q, end_q, "0000006", "0000100"): 
        forex_data.append({"Quarter": r["TIME"], "KRW/100JPY (Avg)": float(r["DATA_VALUE"])})
    
    df_forex = pd.DataFrame(forex_data)
    if not df_forex.empty:
        df_forex = df_forex.groupby("Quarter").mean().reset_index() # groupby drops duplicates safely if they overlap mistakenly, but since columns vary, we should pivot or just take mean (wait, the columns are different!)
    
    # Correct Forex logic:
    df_f = pd.DataFrame(forex_data)
    if not df_f.empty:
        df_f1 = df_f.dropna(subset=["KRW/USD (Avg)"])[["Quarter", "KRW/USD (Avg)"]] if "KRW/USD (Avg)" in df_f.columns else pd.DataFrame(columns=["Quarter"])
        df_f2 = df_f.dropna(subset=["KRW/100JPY (Avg)"])[["Quarter", "KRW/100JPY (Avg)"]] if "KRW/100JPY (Avg)" in df_f.columns else pd.DataFrame(columns=["Quarter"])
        df_forex = df_f1.merge(df_f2, on="Quarter", how="outer")
    else:
        df_forex = pd.DataFrame()
        
    return df_empl, df_forex

@st.cache_data(ttl=3600)
def fetch_macro_rates():
    start_q, end_q = "2018Q1", "2026Q4"
    rates_data = []
    for r in _get_ecos("722Y001", "Q", start_q, end_q, "0101000"):
        rates_data.append({"Quarter": r["TIME"], "Type": "Base Rate (%)", "Val": float(r["DATA_VALUE"])})
    for r in _get_ecos("721Y001", "Q", start_q, end_q, "2010000"):
        rates_data.append({"Quarter": r["TIME"], "Type": "CD Rate 91-day (%)", "Val": float(r["DATA_VALUE"])})
    for r in _get_ecos("121Y006", "Q", start_q, end_q, "BECBLA03"):
        rates_data.append({"Quarter": r["TIME"], "Type": "Household Loans (%)", "Val": float(r["DATA_VALUE"])})
        
    df_rates = pd.DataFrame(rates_data)
    if not df_rates.empty:
        df_rates = df_rates.pivot(index="Quarter", columns="Type", values="Val").reset_index()
    else:
        df_rates = pd.DataFrame()
    return df_rates

def fetch_ecos_macro():
    """Aggregates all macro data for the AI Executive Summary prompt."""
    df_gdp, df_trade, df_cpi = fetch_macro_core()
    df_empl, df_forex = fetch_macro_empl_forex()
    df_rates = fetch_macro_rates()
    df = pd.DataFrame({"Quarter": []})
    for d in [df_gdp, df_trade, df_cpi, df_empl, df_forex, df_rates]:
        if d is not None and not d.empty:
            df = df.merge(d, on="Quarter", how="outer") if not df.empty else d
    if "Quarter" in df.columns:
        df = df[df["Quarter"] >= "2019Q1"].sort_values("Quarter").reset_index(drop=True)
    return df
def get_ai_market_report(data_package):
    """
    Generates a 7-paragraph Executive Commentary using ONE API call.
    Uses Google Search Grounding to research Korean news and policies.
    """
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            return "⚠️ GEMINI_API_KEY missing from secrets. Please configure it in .streamlit/secrets.toml."
        
        # Initialize the 2026 SDK Client
        client = genai.Client(api_key=api_key)
        
        # Consolidate all data into a single string for context
        full_context = "SEOUL OFFICE MARKET DATA (Q1 2026 Prototype):\n"
        for section, df in data_package.items():
            if df is not None:
                full_context += f"\n[{section} - Latest 5 Quarters]\n{df.tail(5).to_string()}\n"

        # The refined prototype prompt based on the Pinnacle Gangnam valuation
        prompt = f"""
        Role: Senior Research Lead, Colliers Seoul.
        Task: Generate a sophisticated, 7-paragraph Executive Market Commentary for the Seoul Office Sector. Synthesize the provided internal database figures with live web grounding to explain the "why" behind the latest trends.        
        INTERNAL DATA CONTEXT (Latest Database Export):
        {full_context}
        
        INSTRUCTIONS FOR ANALYSIS:
        1. Analyze the 'DATA CONTEXT' above. Identify the most recent quarter's figures for GDP, CPI, Vacancy Rates, Face Rents, Net Absorption, and Capital Values.
        2. Identify the trend (e.g., is vacancy rising or falling? Are rents surging or stabilizing?).
        
        LIVE SEARCH INSTRUCTIONS (Grounding):
        - Search the web in Korean (한국어) and English for the latest news from the Bank of Korea (BOK), Molit (국토교통부), and major financial outlets (e.g., Korea Economic Daily).
        - Investigate the real-world reasons behind the trends you found in Step 1. 
        - Look for recent corporate relocations, new supply completions (or delays), and current macroeconomic policy shifts in Seoul that explain the data.
        
        STRUCTURE & TONE REQUIREMENTS:
        - Provide exactly 7 concise paragraphs: 1. Macroeconomics, 2. Existing Supply, 3. Future Pipeline, 4. Vacancy Trends, 5. Net Absorption, 6. Rental Performance, and 7. Capital Markets.
        - Tone: Institutional, analytical, and authoritative. 
        - Style: One cohesive paragraph per topic. Use professional headers. No bullet points.
        """
        
        # Execute the Single API Call with Search Tool Enabled
        response = client.models.generate_content(
            model="gemini-2.5-flash", 
            contents=prompt,
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
                temperature=0.6 
            )
        )
        return response.text
        
    except Exception as e:
        return f"Error generating report: {e}"

# ---------------------------------------------------------
# GLOBAL DESIGN SETTINGS
# ---------------------------------------------------------
# 1. PAGE CONFIGURATION (Must be the very first Streamlit command!)
st.set_page_config(page_title="Seoul Office Market", layout="wide", page_icon="🏢")

# COLLIERS CSS INJECTION
colliers_css = """
<style>
/* Base Fonts and Colors */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Header/Title Styling */
h1 {
    color: #002B49 !important;
    font-weight: 700 !important;
    padding-bottom: 20px;
    border-bottom: 2px solid #00A3E0;
    margin-bottom: 30px;
}
h2, h3, h4 {
    color: #002B49 !important;
}

/* Tabs Styling */
div[data-baseweb="tab-list"] {
    gap: 8px;
}
div[data-baseweb="tab"] {
    background-color: #F2F2F2 !important;
    border-radius: 4px 4px 0 0 !important;
    padding: 10px 16px !important;
    border: none !important;
}
div[data-baseweb="tab"][aria-selected="true"] {
    background-color: #002B49 !important;
}
div[data-baseweb="tab"][aria-selected="true"] p {
    color: white !important;
    font-weight: 600;
}

/* Metrics and Expanders */
div[data-testid="stMetricValue"] {
    color: #002B49 !important;
    font-weight: 700 !important;
}
.streamlit-expanderHeader {
    color: #002B49 !important;
    font-weight: 600 !important;
    background-color: #F8F9F9;
    border-left: 4px solid #00A3E0;
}

/* Base button */
div.stButton > button {
    background-color: #00A3E0;
    color: white;
    border: none;
    border-radius: 4px;
    font-weight: 600;
}
div.stButton > button:hover {
    background-color: #002B49;
    color: white;
}
</style>
"""
st.markdown(colliers_css, unsafe_allow_html=True)

# PLOTLY GLOBALS
pio.templates["colliers"] = go.layout.Template(
    layout=go.Layout(
        colorway=["#002B49", "#00A3E0", "#E2231A", "#8CC63F", "#F26522", "#FFC425", "#63666A", "#1B587C", "#4AC0E0", "#981F28"]
    )
)
pio.templates.default = "plotly_white+colliers"

# PLOTLY EXPORT CONFIG (ALLOWS EDITING & 4K SCREENSHOTS)
CHART_CONFIG = {
    'displayModeBar': True, # Always show modebar
    'editable': True,       # Allows user to rename titles and axes before screenshot!
    'displaylogo': False,   # Hides Plotly logo for clearer enterprise screengrabs
    'toImageButtonOptions': {
        'format': 'png', 
        'filename': 'Colliers_Market_Chart',
        'height': 1080,
        'width': 1920,
        'scale': 2
    }
}


# ---------------------------------------------------------
# AUTHENTICATION
# ---------------------------------------------------------
def check_password():
    """Returns `True` if the user had the correct password."""
    required_password = st.secrets.get("APP_PASSWORD", None)
    
    # If no password is set in the secrets environment, we let them view it
    if not required_password:
        return True

    if st.session_state.get("password_correct", False):
        return True

    # Show Login UI
    st.markdown("<h2 style='text-align: center; color: #002B49; margin-top: 50px;'>🏢 Colliers Enterprise Dashboard</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Please securely log in to access internal market analytics.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.container(border=True):
            pwd = st.text_input("Enter Passcode:", type="password", key="pwd_input")
            if st.button("Unlock Dashboard", use_container_width=True):
                if pwd == required_password:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("Authentication Failed. Invalid Passcode.")
    return False

if not check_password():
    st.stop()  # Do not continue running the script

# 2. MAIN APP TITLE
st.title("🏢 Seoul Office Grade A Market Report")
# --- SIDEBAR REFRESH BUTTON ---
with st.sidebar:
    st.header("⚙️ Dashboard Controls")
    if st.button("🔄 Refresh Data from Excel"):
        st.cache_data.clear()
        st.success("Data refreshed successfully!")
        st.rerun()
# -----------------------------
REGION_COLORS = {
    "CBD": "#002B49",    # Colliers Dark Blue
    "GBD": "#00A3E0",    # Colliers Light Blue
    "YBD": "#8CC63F",    # Colliers Green
    "Overall": "#E2231A", # Colliers Red
    "Other": "#F26522",  # Colliers Orange
    "ETC": "#F26522"     # Colliers Orange Map fallback
}
# --- 3. DATABASE CONFIG ---
DATA_FILE = "seoul_office_data.xlsx"
if not os.path.exists(DATA_FILE):
    st.error(f"Data file '{DATA_FILE}' not found! Please ensure your .xlsx file is in the folder.")
    st.stop()

# ---------------------------------------------------------
# 2. HELPER FUNCTIONS
# ---------------------------------------------------------
@st.cache_data
def load_table(sheet_name):
    try:
        df = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
        
        # Prevent PyArrow serialization errors by ensuring Quarter and Year are strictly strings
        if 'Quarter' in df.columns:
            df['Quarter'] = df['Quarter'].astype(str)
        if 'Year' in df.columns:
            df['Year'] = df['Year'].astype(str)
            
        return df
    except Exception as e:
        return None
@st.cache_data(ttl=3600)
def fetch_dynamic_news(category, region, search_text, limit):
    # 1. The base requirement: It MUST be about commercial real estate
    base_query = '("오피스" OR "상업용 부동산" OR "빌딩")'
    
    # 2. Map the Category dropdown to specific search terms
    cat_keywords = {
        "M&A / Transactions": '("매각" OR "인수" OR "우선협상대상자" OR "펀드" OR "리츠" OR "자산운용")',
        "Macro / Economy": '("기준금리" OR "PF" OR "프로젝트파이낸싱" OR "환율" OR "한국은행")',
        "Development / Supply": '("개발" OR "인허가" OR "착공" OR "준공" OR "재건축" OR "리모델링")',
        "Leasing / Vacancy": '("공실" OR "임대차" OR "임대료" OR "사옥" OR "이전")'
    }
    
    # 3. Map the Region dropdown to specific search terms
    reg_keywords = {
        "CBD": '("도심" OR "종로" OR "을지로" OR "광화문" OR "시청" OR "서울역")',
        "GBD": '("강남" OR "테헤란로" OR "서초" OR "삼성역" OR "역삼" OR "잠실")',
        "YBD": '("여의도" OR "파크원" OR "IFC" OR "영등포")',
        "Other": '("성수" OR "판교" OR "마곡" OR "용산" OR "분당")'
    }
    
    # 4. Construct the dynamic query
    query_parts = [base_query]
    
    if region != "Overall":
        query_parts.append(reg_keywords.get(region, ""))
    else:
        query_parts.append('("서울")')
        
    if category != "All":
        query_parts.append(cat_keywords.get(category, ""))
        
    if search_text.strip():
        query_parts.append(f'"{search_text.strip()}"')
        
    final_query = " AND ".join(query_parts)
    encoded_query = urllib.parse.quote(final_query)
    
    url = f"https://news.google.com/rss/search?q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko"
    
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        
        news_list = []
        for item in root.findall('./channel/item')[:limit]:
            title = item.find('title').text
            link = item.find('link').text
            pubDate = item.find('pubDate').text
            source = item.find('source').text if item.find('source') is not None else "News"
            
            dt_obj = pd.to_datetime(pubDate)
            news_list.append({"title": title, "link": link, "dt": dt_obj, "source": source, "date": dt_obj.strftime("%Y-%m-%d %H:%M")})
            
        news_list.sort(key=lambda x: x['dt'], reverse=True)
        return news_list
    except Exception as e:
        return []

@st.cache_data
def fetch_seoul_geojson():
    # Public open-source GeoJSON containing Seoul district boundaries
    url = "https://raw.githubusercontent.com/southkorea/seoul-maps/master/kostat/2013/json/seoul_municipalities_geo_simple.json"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            return json.loads(response.read().decode('utf-8'))
    except Exception as e:
        return None
# ---------------------------------------------------------
# 3. DASHBOARD LOGIC
# ---------------------------------------------------------

# Updated Tab List (Added 'Summary' at index 0)
tabs = st.tabs([
    "📑 Executive Summary", "1. Macro: Core & Trade", "2. Macro: Empl & Forex", "3. Macro: Rates",
    "4. Supply", "5. Future", "6. Vacancy", "7. Absorption", "8. Rent", "9. Capital", "10. Transactions", "11. News"
])

# 0. Executive Summary Page
with tabs[0]:
    st.header("🏢 Q1 2026 Executive Market Commentary")
    st.caption("AI-Powered Synthesis: Internal Valuation Data + Live Web Grounding")

    if st.button("✨ Generate Commentary"):
        with st.spinner("Analyzing data and searching the Korean web for market context (this may take 15-30 seconds)..."):
            
            # Gather all local data
            package = {
                "Macroeconomics": fetch_ecos_macro(),
                "Existing Supply": load_table("existing_supply"),
                "Future Pipeline": load_table("future_pipeline"),
                "Vacancy": load_table("vacancy"),
                "Net Absorption": load_table("net_absorption"),
                "Rent Performance": load_table("rent"),
                "Capital Markets": load_table("cap_rate")
            }
            
            # Execute the single-call report
            full_report = get_ai_market_report(package)
            
            # Display results
            st.markdown("---")
            st.markdown(full_report)
            
            st.info("💡 **Methodology:** This commentary synthesizes your internal SQL data with real-time web searches of Korean financial news and BOK policy updates.")
    else:
        st.write("Click the button above to generate the grounded market commentary.")

# 1. Macro: Core & Trade
with tabs[1]:
    st.header("Macro: Core Indicators & Trade")
    df_gdp, df_trade, df_cpi = fetch_macro_core()
    
    if not df_gdp.empty:
        latest_gdp = df_gdp.iloc[-1]
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Real GDP (KRW Tr)", f"{latest_gdp.get('Real GDP (KRW Tr)', 0):.2f}")
        m2.metric("Real GDP Growth (YoY)", f"{latest_gdp.get('Real GDP Growth (YoY %)', 0):.2f}%")
        m3.metric("Real GDP Growth (QoQ)", f"{latest_gdp.get('Real GDP Growth (QoQ %)', 0):.2f}%")
        
        with st.expander("📈 View GDP Trend", expanded=True):
            fig_gdp = make_subplots(specs=[[{"secondary_y": True}]])
            fig_gdp.add_trace(go.Bar(x=df_gdp["Quarter"], y=df_gdp["Real GDP (KRW Tr)"], name="Real GDP (KRW Tr)", opacity=0.8, marker_color="#002B49"), secondary_y=False)
            if "Real GDP Growth (YoY %)" in df_gdp.columns:
                fig_gdp.add_trace(go.Scatter(x=df_gdp["Quarter"], y=df_gdp["Real GDP Growth (YoY %)"], mode="lines+markers", name="YoY (%)", line=dict(color="#E2231A")), secondary_y=True)
            if "Real GDP Growth (QoQ %)" in df_gdp.columns:
                fig_gdp.add_trace(go.Scatter(x=df_gdp["Quarter"], y=df_gdp["Real GDP Growth (QoQ %)"], mode="lines+markers", name="QoQ (%)", line=dict(color="#00A3E0")), secondary_y=True)
            fig_gdp.update_layout(title="GDP Growth Rates & Real GDP", hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig_gdp.update_yaxes(title_text="Trillion KRW", secondary_y=False)
            fig_gdp.update_yaxes(title_text="Growth (%)", secondary_y=True)
            st.plotly_chart(fig_gdp, width="stretch", config=CHART_CONFIG)
            st.dataframe(df_gdp.dropna().reset_index(drop=True), hide_index=True, use_container_width=True)

    if not df_trade.empty:
        with st.expander("🚢 View Trade Data (Aggregated Quarterly)", expanded=True):
            fig_trade = make_subplots(specs=[[{"secondary_y": True}]])
            if "Export (USD Bn)" in df_trade.columns:
                fig_trade.add_trace(go.Bar(x=df_trade["Quarter"], y=df_trade["Export (USD Bn)"], name="Export", marker_color="#002B49"), secondary_y=False)
            if "Import (USD Bn)" in df_trade.columns:
                fig_trade.add_trace(go.Bar(x=df_trade["Quarter"], y=df_trade["Import (USD Bn)"], name="Import", marker_color="#00A3E0"), secondary_y=False)
            if "Trade Balance (USD Bn)" in df_trade.columns:
                fig_trade.add_trace(go.Scatter(x=df_trade["Quarter"], y=df_trade["Trade Balance (USD Bn)"], name="Trade Balance", mode="lines+markers", line=dict(color="#E2231A", width=3)), secondary_y=True)
            fig_trade.update_layout(title="Export, Import and Trade Balance (USD Billion)", barmode="group", hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_trade, width="stretch", config=CHART_CONFIG)
            
    if not df_cpi.empty:
        with st.expander("🛒 View CPI breakdown", expanded=True):
            cols = [c for c in df_cpi.columns if c not in ("Quarter", "Total CPI Growth (YoY %)")]
            fig_cpi = px.line(df_cpi, x="Quarter", y=cols, title="CPI Index by Category", markers=True)
            st.plotly_chart(fig_cpi, width="stretch", config=CHART_CONFIG)
            st.dataframe(df_cpi, hide_index=True)

# 2. Macro: Empl & Forex
with tabs[2]:
    st.header("Macro: Employment & Forex")
    df_empl, df_forex = fetch_macro_empl_forex()
    
    if not df_empl.empty:
        with st.expander("👤 View Employment Statistics", expanded=True):
            fig_empl = make_subplots(specs=[[{"secondary_y": True}]])
            if "Employed Pop (000s)" in df_empl.columns:
                fig_empl.add_trace(go.Bar(x=df_empl["Quarter"], y=df_empl["Employed Pop (000s)"], name="Employed Pop (Thousands)", marker_color="#002B49", opacity=0.8), secondary_y=False)
            if "Unemployment Rate (%)" in df_empl.columns:
                fig_empl.add_trace(go.Scatter(x=df_empl["Quarter"], y=df_empl["Unemployment Rate (%)"], name="Unemployment Rate", mode="lines+markers", line=dict(color="#E2231A", width=3)), secondary_y=True)
            fig_empl.update_layout(title="Employment Statistics", hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig_empl.update_yaxes(title_text="Thousands", secondary_y=False)
            fig_empl.update_yaxes(title_text="Rate (%)", secondary_y=True)
            st.plotly_chart(fig_empl, width="stretch", config=CHART_CONFIG)
            
    if not df_forex.empty:
        with st.expander("💱 View Exchange Rates", expanded=True):
            cols = [c for c in df_forex.columns if c != "Quarter"]
            fig_fx = px.line(df_forex, x="Quarter", y=cols, title="Average Exchange Rates", markers=True)
            st.plotly_chart(fig_fx, width="stretch", config=CHART_CONFIG)

# 3. Macro: Rates
with tabs[3]:
    st.header("Macro: Interest Rates")
    df_rates = fetch_macro_rates()
    
    if not df_rates.empty:
        with st.expander("🏦 View Interest Rates Trend", expanded=True):
            cols = [c for c in df_rates.columns if c != "Quarter"]
            fig_rates = px.line(df_rates, x="Quarter", y=cols, title="Interest Rates (%)", markers=True)
            st.plotly_chart(fig_rates, width="stretch", config=CHART_CONFIG)
            st.dataframe(df_rates, hide_index=True)

# 2. Existing Supply
with tabs[4]:
    st.header("🏢 Existing Supply Distribution (Pyeong)")
    df_supply = load_table("existing_supply")
    
    if df_supply is not None and not df_supply.empty:
        # Clean numeric columns
        for c in ["CBD", "GBD", "YBD", "Overall"]:
            if c in df_supply.columns:
                df_supply[c] = pd.to_numeric(df_supply[c].astype(str).str.replace(',', ''), errors='coerce')
        
        # 📌 Headline Metric Box
        display_latest_metrics(df_supply, "Total Stock", format_type="number")
        
        # --- NEW FOLDABLE SECTION: Pie Chart & Table ---
        with st.expander("📊 View Regional Distribution (Chart & Data)", expanded=True):
            latest = df_supply.iloc[-1]
            pie_df = pd.DataFrame({
                "District": ["CBD", "GBD", "YBD"], 
                "Stock": [latest.get("CBD", 0), latest.get("GBD", 0), latest.get("YBD", 0)]
            })
            
            fig_pie = px.pie(
                pie_df, values='Stock', names='District', hole=0.4, 
                title="Total Stock Proportion",
                color='District', color_discrete_map=REGION_COLORS 
            )
            fig_pie.update_layout(height=500, margin=dict(t=40, b=0, l=0, r=0))
            st.plotly_chart(fig_pie, width='stretch', config=CHART_CONFIG)

            # Calculate percentages and format the numbers
            total_stock = pie_df["Stock"].sum()
            display_df = pie_df.copy()
            display_df["Share (%)"] = (display_df["Stock"] / total_stock).apply(lambda x: f"{x:.1%}")
            display_df["Stock"] = display_df["Stock"].apply(lambda x: f"{x:,.0f}")
            
            # Use columns to keep the table acting like a "small box" in the center
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.markdown("<h5 style='text-align: center;'>🔢 Regional Distribution Breakdown</h5>", unsafe_allow_html=True)
                st.dataframe(display_df, width='stretch', hide_index=True)

    st.markdown("---")
    
    # --- UNDERNEATH: Expandable Historical Supply Section ---
    with st.expander("📈 View Historical Supply Trend & Data", expanded=False):
        df_hist_supply = load_table("historical_supply")
        
        if df_hist_supply is not None and not df_hist_supply.empty:
            y_col = [c for c in df_hist_supply.columns if c != "Quarter"][0]
            
            zoom_in = st.toggle("🔍 Zoom in on trend (Cut Y-Axis bottom)")
            
            fig_hist = px.bar(
                df_hist_supply, x="Quarter", y=y_col, 
                title=f"Historical Supply Trend ({y_col})",
                color_discrete_sequence=["#3498DB"] 
            )
            
            if zoom_in:
                min_y = df_hist_supply[y_col].min()
                max_y = df_hist_supply[y_col].max()
                padding = (max_y - min_y) * 0.1 if max_y != min_y else max_y * 0.1
                fig_hist.update_yaxes(range=[max(0, min_y - padding), max_y + padding])
                
            st.plotly_chart(fig_hist, width='stretch', config=CHART_CONFIG)
            
            st.markdown("**Raw Data**")
            st.dataframe(df_hist_supply, width='stretch', hide_index=True)
            
        else:
            st.info("💡 To view the historical trend chart, please add a sheet named 'historical_supply' to your Excel file with 'Quarter' and your single variable data.")
# 3. Future Supply Pipeline (Pyeong)
with tabs[5]:
    st.header("🚀 Future Supply Pipeline (Pyeong)")
    df_future = load_table("future_supply")
    
    if df_future is not None and not df_future.empty:
        # Find the raw GFA column dynamically
        raw_gfa_col = next((c for c in df_future.columns if "GFA" in c or "pyeong" in c), None)
        
        if raw_gfa_col:
            df_future['GFA_Numeric'] = pd.to_numeric(df_future[raw_gfa_col].astype(str).str.replace(',', ''), errors='coerce')
            
            # NEW: AI Summary Box
        raw_gfa_col = next((c for c in df_future.columns if "GFA" in c or "pyeong" in c), None)
        
        if raw_gfa_col:
            # Create a clean numeric version for calculations/graphing
            df_future['GFA_Numeric'] = pd.to_numeric(df_future[raw_gfa_col].astype(str).str.replace(',', ''), errors='coerce')
            
            # Frame the Chart
            with st.expander("📊 View Upcoming Supply Pipeline Chart", expanded=True):
                fig_fut = px.bar(
                    df_future, x="Year", y="GFA_Numeric", color="Submarket", barmode="group", 
                    title="Upcoming Supply by Year (Pyeong)",
                    color_discrete_map=REGION_COLORS # Apply global colors
                )
                fig_fut.update_yaxes(title_text="GFA (Pyeong)")
                st.plotly_chart(fig_fut, width="stretch", config=CHART_CONFIG)
            
            # Frame the Table
            with st.expander("📄 View Future Supply Details", expanded=False):
                df_display = df_future.drop(columns=['GFA_Numeric']).rename(columns={raw_gfa_col: "Estimated GFA (Pyeong)"})
                st.dataframe(df_display, width="stretch", hide_index=True)

# 4. Vacancy 
with tabs[6]:
    st.header("Vacancy Rate Trend")
    df_vac = load_table("vacancy")
    if df_vac is not None and not df_vac.empty:
        display_latest_metrics(df_vac, "Vacancy Rate", format_type="percent")
      
        with st.expander("📈 View Vacancy Rate Trends", expanded=True):
            expected_cols = ["CBD", "GBD", "YBD", "Overall"]
            available_cols = [col for col in expected_cols if col in df_vac.columns]
            extra_cols = [c for c in df_vac.columns if c not in ["Quarter", "Indicator"] and c not in expected_cols]
            
            fig = px.line(df_vac, x="Quarter", y=available_cols + extra_cols, markers=True, color_discrete_map=REGION_COLORS)
            fig.update_layout(yaxis_tickformat='.1%')
            st.plotly_chart(fig, width="stretch", config=CHART_CONFIG)
        
        with st.expander("📄 View Vacancy Data Table", expanded=False):
            st.subheader("Data Table")
            display_df_with_changes(df_vac, is_percent=True)

# 5. Net Absorption 
with tabs[7]:
    st.header("Net Absorption(Pyeong)")
    df_abs = load_table("net_absorption")
    if df_abs is not None and not df_abs.empty:
        for col in ["CBD", "GBD", "YBD", "Overall"]:
            if col in df_abs.columns:
                df_abs[col] = df_abs[col].astype(str).str.replace(',', '').apply(pd.to_numeric, errors='coerce')
        
        display_latest_metrics(df_abs, "Net Absorption")
        
        with st.expander("📊 View Net Absorption Chart", expanded=True):
            expected_cols = ["CBD", "GBD", "YBD", "Overall"]
            available_cols = [col for col in expected_cols if col in df_abs.columns]
            extra_cols = [c for c in df_abs.columns if c not in ["Quarter", "Indicator"] and c not in expected_cols]
            
            fig_abs = px.bar(
                df_abs, x="Quarter", y=available_cols + extra_cols, barmode='group',
                title="Quarterly Net Absorption (Pyeong) by Submarket",
                color_discrete_map=REGION_COLORS # Matching colors
            )
            st.plotly_chart(fig_abs, width="stretch", config=CHART_CONFIG)
        
        # Frame the table
        with st.expander("📄 View Absorption Data Table", expanded=False):
            st.subheader("Data Table")
            display_df_with_changes(df_abs)

# 6. Rent Performance 
with tabs[8]:
    st.header("Rent Performance Trend")
    df_rent = load_table("rent")
    if df_rent is not None and not df_rent.empty:
        display_latest_metrics(df_rent, "Rent Performance")
        
        # Frame the chart
        with st.expander("📈 View Rent Performance Trends", expanded=True):
            expected_cols = ["CBD", "GBD", "YBD", "Overall"]
            available_cols = [col for col in expected_cols if col in df_rent.columns]
            extra_cols = [c for c in df_rent.columns if c not in ["Quarter", "Indicator"] and c not in expected_cols]
            
            fig2 = px.line(
                df_rent, x="Quarter", y=available_cols + extra_cols, markers=True,
                color_discrete_map=REGION_COLORS # Matching colors
            )
            st.plotly_chart(fig2, width="stretch", config=CHART_CONFIG)
        
        # Frame the table
        with st.expander("📄 View Rent Data Table", expanded=False):
            st.subheader("Data Table")
            display_df_with_changes(df_rent, is_percent=False)

# 7. Capital Markets
with tabs[9]:
    st.header("Capital Markets Analysis")
    
    # Pre-load the data so the summary can read it
    df_cv = load_table("capital_value")
    df_rate = load_table("cap_rate")
            
    col1, col2 = st.columns(2)
    
    with col1:
        # Note: I removed the duplicate 'load_table' line here since we loaded it above!
        if df_cv is not None and not df_cv.empty:
            for c in ["CBD", "GBD", "YBD", "Overall"]:
                if c in df_cv.columns:
                    df_cv[c] = df_cv[c].astype(str).str.replace(',', '').apply(pd.to_numeric, errors='coerce')
            
            display_latest_metrics(df_cv, "Capital Value", format_type="number")
            
            # Frame the left chart
            with st.expander("📈 CV Data Chart", expanded=True):
                fig_cv = px.line(df_cv, x="Quarter", y=["CBD", "GBD", "YBD", "Overall"], markers=True, title="Capital Value (KRW/P)", color_discrete_map=REGION_COLORS)
                st.plotly_chart(fig_cv, width="stretch", config=CHART_CONFIG)
            
            # Frame the left table
            with st.expander("📄 CV Data Table", expanded=False):
                display_df_with_changes(df_cv)

    with col2:
        df_rate = load_table("cap_rate")
        if df_rate is not None and not df_rate.empty:
            for c in ["CBD", "GBD", "YBD", "Overall"]:
                if c in df_rate.columns:
                    df_rate[c] = df_rate[c].astype(str).str.replace('%', '').apply(pd.to_numeric, errors='coerce')
            
            display_latest_metrics(df_rate, "Cap Rate", format_type="percent")
            
            # Frame the right chart
            with st.expander("📈 Cap Rate Chart", expanded=True):
                fig_rate = px.line(df_rate, x="Quarter", y=["CBD", "GBD", "YBD", "Overall"], markers=True, title="Cap Rate (%)", color_discrete_map=REGION_COLORS)
                st.plotly_chart(fig_rate, width="stretch", config=CHART_CONFIG)
            
            # Frame the right table
            with st.expander("📄 Cap Rate Table", expanded=False):
                display_df_with_changes(df_rate, is_percent=True)

# 8. Transactions
with tabs[10]:
    st.header("Major Transactions Analysis")
    df_cap_raw = load_table("capital_markets")
    
    if df_cap_raw is not None and not df_cap_raw.empty:
        df_cap_raw['Consideration_Num'] = df_cap_raw['Consideration'].astype(str).str.replace(',', '').apply(pd.to_numeric, errors='coerce')
        
        # --- FOLDABLE MAP VISUALIZATION ---
        with st.expander("📍 View Transaction Map & Submarket Regions", expanded=False):
            if "Latitude" in df_cap_raw.columns and "Longitude" in df_cap_raw.columns:
                df_cap_raw["Latitude"] = pd.to_numeric(df_cap_raw["Latitude"], errors="coerce")
                df_cap_raw["Longitude"] = pd.to_numeric(df_cap_raw["Longitude"], errors="coerce")
                map_df = df_cap_raw.dropna(subset=["Latitude", "Longitude", "Consideration_Num"])
                
                if not map_df.empty:
                    hover_cols = {
                        "Quarter": True, "Consideration": True, "Latitude": False, "Longitude": False
                    }
                    for extra_col in [
                        "TransactedGFApy", "UnitRatebyKRWpy", "UnitRatebyKRWsq. m.", 
                        "Cap Rate", "DealStructure", "Buyer", "Seller", "Investor Type", "InvestorType"
                    ]:
                        if extra_col in map_df.columns:
                            hover_cols[extra_col] = True

                    # 1. Base Scatter Map
                    fig_map = px.scatter_map(
                        map_df, lat="Latitude", lon="Longitude", 
                        color="Subdistrict", size="Consideration_Num", 
                        hover_name="Property", 
                        hover_data=hover_cols,
                        zoom=11, center={"lat": 37.54, "lon": 126.98}, 
                        map_style="carto-positron", 
                        height=800,
                        color_discrete_map=REGION_COLORS,
                        size_max=25  # 👈 NEW: Forces the biggest deals to be drawn much larger!
                    )
                    
                    # 2. Add the Shaded Background Regions
                    seoul_geo = fetch_seoul_geojson()
                    map_layers = [] 
                    
                    if seoul_geo:
                        cbd_feat = [f for f in seoul_geo['features'] if f['properties']['name'] in ['종로구', '중구', '용산구', '서대문구']]
                        if cbd_feat:
                            map_layers.append({"sourcetype": "geojson", "source": {"type": "FeatureCollection", "features": cbd_feat}, "type": "fill", "color": "rgba(231, 76, 60, 0.2)"})
                        
                        gbd_feat = [f for f in seoul_geo['features'] if f['properties']['name'] in ['강남구', '서초구', '송파구']]
                        if gbd_feat:
                            map_layers.append({"sourcetype": "geojson", "source": {"type": "FeatureCollection", "features": gbd_feat}, "type": "fill", "color": "rgba(52, 152, 219, 0.2)"})
                        
                        ybd_feat = [f for f in seoul_geo['features'] if f['properties']['name'] in ['영등포구']]
                        if ybd_feat:
                            map_layers.append({"sourcetype": "geojson", "source": {"type": "FeatureCollection", "features": ybd_feat}, "type": "fill", "color": "rgba(46, 204, 113, 0.2)"})

                    fig_map.update_layout(map_layers=map_layers, margin={"r":0,"t":0,"l":0,"b":0})
                    
                    # 👈 NEW: Lowered opacity so overlapping bubbles look better, and lowered sizemin
                    fig_map.update_traces(marker=dict(opacity=0.75, sizemin=4)) 
                    
                    st.plotly_chart(fig_map, width='stretch', config=CHART_CONFIG)
                else:
                    st.info("No valid latitude/longitude coordinates found to map.")
        
        # --- FOLDABLE TIMELINE SCATTER CHART ---
        with st.expander("📈 View Major Transactions Timeline", expanded=True):
            scatter_hover = ["Property", "TransactedGFApy"]
            for extra_col in [
                "UnitRatebyKRWpy", "UnitRatebyKRWsq. m.", "Cap Rate", 
                "DealStructure", "Buyer", "Seller", "Investor Type", "InvestorType"
            ]:
                if extra_col in df_cap_raw.columns:
                    scatter_hover.append(extra_col)
                    
            fig_scatter = px.scatter(
                df_cap_raw.dropna(subset=['Consideration_Num']), 
                x="Quarter", y="Consideration_Num", color="Subdistrict", 
                size="Consideration_Num", hover_data=scatter_hover,
                title="Transaction Value by Quarter (Hover for Details)",
                color_discrete_map={"CBD": "#e74c3c", "GBD": "#3498db", "YBD": "#2ecc71", "ETC": "#f1c40f"},
                labels={
                    "TransactedGFApy": "Transacted GFA (Pyeong)", 
                    "Consideration_Num": "Consideration (KRW)",
                    "UnitRatebyKRWpy": "Unit Rate (KRW/Pyeong)",
                    "UnitRatebyKRWsq. m.": "Unit Rate (KRW/sqm)"
                } 
            )
            st.plotly_chart(fig_scatter, width='stretch', config=CHART_CONFIG)
        
        # --- FOLDABLE RAW DATA TABLE ---
        with st.expander("📄 View Raw Transaction Data", expanded=True):
            st.dataframe(df_cap_raw.drop(columns=['Consideration_Num']), width='stretch', hide_index=True)

# 9. News Tab
with tabs[11]:
    st.header("📰 Live Market News")
    
    # 1. UI Controls
    search_query = st.text_input("🔍 Search Exact Keywords (e.g., '삼성SDS', '타워8'):", "")
    
    filter_col1, filter_col2, filter_col3 = st.columns([2, 2, 1])
    with filter_col1:
        selected_cat = st.selectbox("Category:", ["All", "M&A / Transactions", "Macro / Economy", "Development / Supply", "Leasing / Vacancy"])
    with filter_col2:
        selected_reg = st.selectbox("Region:", ["Overall", "CBD", "GBD", "YBD", "Other"])
    with filter_col3:
        article_limit = st.selectbox("Articles to fetch:", [10, 20, 50, 100], index=1)

    st.markdown("---")
    
    # 2. Fetch data directly based on the UI controls
    with st.spinner("Fetching live news from search engine..."):
        display_news = fetch_dynamic_news(selected_cat, selected_reg, search_query, article_limit)
    
    # 3. Display the Results
    if display_news:
        st.caption(f"Showing the top **{len(display_news)}** most recent articles matching your criteria.")
        col1, col2 = st.columns(2)
        for i, n in enumerate(display_news):
            with (col1 if i % 2 == 0 else col2):
                with st.container(border=True):
                    st.markdown(f"**[{n['title']}]({n['link']})**")
                    st.caption(f"📅 {n['date']} | 🏢 {n['source']}")
    else:
        st.warning("No recent news found for this exact combination. Try broadening your search or changing the region!")

