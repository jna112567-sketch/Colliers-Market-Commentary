import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
import urllib.request
import xml.etree.ElementTree as ET
import os
import json
import urllib.parse
import plotly.io as pio
import requests
from google import genai
import time
from google.api_core import exceptions

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
        st.markdown(f"#### 📌 {title} Headline ({latest_qtr})")
        
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
def fetch_ecos_macro():
    """Fetches quarterly Macro data (2019-2026) with corrected item codes."""
    try:
        API_KEY = st.secrets.get("ECOS_API_KEY", "GUP36MNVBH5Y1PO2AS9S")
    except Exception:
        API_KEY = "GUP36MNVBH5Y1PO2AS9S"

    # Updated Indicators: Unemployment now uses the direct I61BC/I28A path
    indicators = {
        "Real GDP Growth (%)": ("200Y102", "Q", "10111"),
        "Unemployment (%)": ("901Y027", "Q", "I61BC/I28A"), # Removed "0/" prefix
        "CPI Index": ("901Y009", "Q", "0")
    }

    all_data = []
    start_q, end_q = "2018Q1", "2026Q4"

    try:
        for label, (table, cycle, item) in indicators.items():
            url = f"http://ecos.bok.or.kr/api/StatisticSearch/{API_KEY}/json/kr/1/100/{table}/{cycle}/{start_q}/{end_q}/{item}"
            response = requests.get(url)
            data = response.json()
            
            if "StatisticSearch" in data and "row" in data["StatisticSearch"]:
                for r in data["StatisticSearch"]["row"]:
                    all_data.append({
                        "Quarter": r["TIME"],
                        "Indicator": label,
                        "Value": float(r["DATA_VALUE"])
                    })
        
        if not all_data: return None

        df_raw = pd.DataFrame(all_data)
        df = df_raw.pivot(index="Quarter", columns="Indicator", values="Value").reset_index()
        
        # --- CALCULATIONS ---
        if "CPI Index" in df.columns:
            # Only calculating the Year-on-Year Growth Rate now
            df["CPI Growth Rate (%)"] = (df["CPI Index"] / df["CPI Index"].shift(4) - 1) * 100
        
        # Safety initialization for required display columns
        required_cols = ["Quarter", "CPI Index", "CPI Growth Rate (%)", "Real GDP Growth (%)", "Unemployment (%)"]
        for col in required_cols:
            if col not in df.columns:
                df[col] = 0.0
        
        # Filter for 2019 onwards
        df = df[df["Quarter"] >= "2019Q1"].sort_values("Quarter").reset_index(drop=True)
        return df[required_cols]

    except Exception as e:
        st.error(f"ECOS API Error: {e}")
        return None
def get_ai_summary(tab_name, df_context):
    """Generates a summary with a retry loop to handle 429 Rate Limits."""
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        return "⚠️ API Key missing."
        
    client = genai.Client(api_key=api_key)
    data_summary = df_context.tail(5).to_string() 
    
    prompt = f"Write a 100-200 word summary for '{tab_name}' based on: {data_summary}..."

    # --- RETRY LOGIC (Exponential Backoff) ---
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt
            )
            return response.text
        except Exception as e:
            # If we hit a 429 error, wait and try again
            if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
                wait_time = (attempt + 1) * 5  # Wait 5s, then 10s, then 15s
                st.warning(f"Rate limit hit for {tab_name}. Retrying in {wait_time}s...")
                time.sleep(wait_time)
            else:
                return f"Error: {e}"
    
    return "❌ Failed to generate summary after multiple retries due to rate limits."

# ---------------------------------------------------------
# 1. APP CONFIGURATION & STYLING
# ---------------------------------------------------------
st.set_page_config(page_title="Seoul Office Market Dashboard", layout="wide")

st.markdown("""
<style>
    /* 1. Add breathing room to the top and sides of the page */
    .block-container { 
        padding-top: 2rem !important; 
        padding-bottom: 2rem !important; 
        max-width: 95% !important; 
    }
    
    /* 2. Style the tabs structurally (letting Streamlit handle the colors natively) */
    .stTabs [data-baseweb="tab-list"] { 
        gap: 8px; 
    }
    .stTabs [data-baseweb="tab"] { 
        border-radius: 6px 6px 0px 0px; 
        padding: 10px 20px; 
        border: 1px solid rgba(128, 128, 128, 0.2); 
        border-bottom: none; 
    }
    
    /* 3. Cards (Containers) with soft adaptive drop-shadows */
    [data-testid="stVerticalBlockBorderWrapper"] { 
        border-radius: 10px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
        padding: 0.5rem; 
    }
</style>
""", unsafe_allow_html=True)

st.title("🏢 Seoul Office Grade A Market Report")

DB_NAME = "seoul_office.db"
if not os.path.exists(DB_NAME):
    st.error(f"Database '{DB_NAME}' not found! Please ensure your .db file is in the folder.")
    st.stop()

# ---------------------------------------------------------
# GLOBAL DESIGN SETTINGS
# ---------------------------------------------------------
REGION_COLORS = {
    "CBD": "#E74C3C",    # Red
    "GBD": "#3498DB",    # Blue
    "YBD": "#2ECC71",    # Green
    "Overall": "#8E44AD", # Purple
    "Other": "#F39C12",  # Orange
    "ETC": "#F39C12"     # Orange (for map fallback)
}
if not os.path.exists(DB_NAME):
    st.error(f"Database '{DB_NAME}' not found! Please ensure your .db file is in the folder.")
    st.stop()

# ---------------------------------------------------------
# 2. HELPER FUNCTIONS
# ---------------------------------------------------------
@st.cache_data
def load_table(table_name):
    try:
        conn = sqlite3.connect(DB_NAME)
        df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
        conn.close()
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
    "📑 Executive Summary", "1. Macro", "2. Supply", "3. Future", "4. Vacancy", 
    "5. Absorption", "6. Rent", "7. Capital", "8. Transactions", "9. News", "10. 🛠️ Admin"
])

# 0. Executive Summary Page
with tabs[0]:
    st.header("🏢 Seoul Office Market: AI Executive Summary")
    
    if st.button("✨ Generate / Refresh Market Summary"):
        with st.spinner("AI is analyzing market trends (spacing requests to avoid rate limits)..."):
            
            # --- 1. Macro Analysis ---
            df_macro = fetch_ecos_macro()
            if df_macro is not None:
                with st.expander("🇰🇷 Macroeconomic Outlook", expanded=True):
                    st.write(get_ai_summary("Macroeconomic Trends", df_macro))
                time.sleep(2) # ⏸️ Pause for 2 seconds
            
            # --- 2. Vacancy Analysis ---
            df_vac = load_table("vacancy")
            if df_vac is not None:
                with st.expander("📉 Vacancy & Occupancy Behavior", expanded=True):
                    st.write(get_ai_summary("Vacancy Rates", df_vac))
                time.sleep(2) # ⏸️ Pause for 2 seconds

            # --- 3. Rent Analysis ---
            df_rent = load_table("rent")
            if df_rent is not None:
                with st.expander("💰 Rental Performance Analysis", expanded=True):
                    st.write(get_ai_summary("Rent Performance", df_rent))
                time.sleep(2) # ⏸️ Pause for 2 seconds
                    
            # --- 4. Capital Markets Analysis ---
            df_cap = load_table("cap_rate")
            if df_cap is not None:
                with st.expander("🏛️ Capital Markets & Yields", expanded=True):
                    st.write(get_ai_summary("Cap Rates and Capital Value", df_cap))
            
            st.success("✅ Summary generation complete!")
    else:
        st.info("Click the button above to generate the report. It takes about 15-20 seconds to complete due to rate limiting safety.")

# 1. Macro
with tabs[1]:
    st.header("🇰🇷 Macroeconomic Overview (2019-2026)")
    
    df_macro_live = fetch_ecos_macro()

    if df_macro_live is not None:
        latest = df_macro_live.iloc[-1]
        
        # 1. Headline Metrics - We keep these visible for a quick summary
        m1, m2, m3 = st.columns(3)
        m1.metric("Real GDP Growth", f"{latest.get('Real GDP Growth (%)', 0):.2f}%")
        m2.metric("Inflation (YoY)", f"{latest.get('CPI Growth Rate (%)', 0):.2f}%")
        m3.metric("Unemployment", f"{latest.get('Unemployment (%)', 0):.1f}%")

        # 2. FOLDABLE CHART: Set to expanded=True so it's visible on load
        with st.expander("📈 View Macroeconomic Trends Chart", expanded=True):
            chart_cols = ["Real GDP Growth (%)", "Unemployment (%)", "CPI Growth Rate (%)"]
            fig = px.line(
                df_macro_live, 
                x="Quarter", 
                y=[c for c in chart_cols if c in df_macro_live.columns],
                markers=True,
                title="Growth, Unemployment, and Inflation Trends",
                color_discrete_map={
                    "Real GDP Growth (%)": "#3498DB",
                    "Unemployment (%)": "#E74C3C",
                    "CPI Growth Rate (%)": "#F1C40F"
                }
            )
            fig.update_layout(yaxis_title="Percentage (%)", hovermode="x unified")
            st.plotly_chart(fig, width='stretch')

        # 3. FOLDABLE TABLE: Set to expanded=False to save space
        with st.expander("📄 View Detailed Statistics Table", expanded=False):
            st.dataframe(
                df_macro_live.style.format({
                    "CPI Index": "{:.1f}",
                    "CPI Growth Rate (%)": "{:.2f}%",
                    "Real GDP Growth (%)": "{:.2f}%",
                    "Unemployment (%)": "{:.1f}%"
                }), 
                width="stretch", hide_index=True
            )
    else:
        st.warning("Data fetch failed. Ensure your ECOS API key is valid.")

# 2. Existing Supply
with tabs[2]:
    st.header("🏢 Existing Supply Distribution (Pyeong)")
    df_supply = load_table("existing_supply")
    
    if df_supply is not None and not df_supply.empty:
        # Clean numeric columns
        for c in ["CBD", "GBD", "YBD", "Overall"]:
            if c in df_supply.columns:
                df_supply[c] = pd.to_numeric(df_supply[c].astype(str).str.replace(',', ''), errors='coerce')
        
        # 📌 Headline Metric Box
        display_latest_metrics(df_supply, "Total Stock", format_type="number")
        
        # Top Box: Pie Chart
        # Foldable Data Table
        with st.expander("📊 View Detailed Supply Data", expanded=True):
            st.dataframe(df_supply, width="stretch", hide_index=True) 
            # Create Pie Chart using ONLY the latest row
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
            fig_pie.update_layout(margin=dict(t=40, b=0, l=0, r=0))
            st.plotly_chart(fig_pie, width="stretch")

# 3. Future Supply Pipeline (Pyeong)
with tabs[3]:
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
                st.plotly_chart(fig_fut, width="stretch")
            
            # Frame the Table
            with st.expander("📄 View Future Supply Details", expanded=False):
                df_display = df_future.drop(columns=['GFA_Numeric']).rename(columns={raw_gfa_col: "Estimated GFA (Pyeong)"})
                st.dataframe(df_display, width="stretch", hide_index=True)

# 4. Vacancy 
with tabs[4]:
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
            st.plotly_chart(fig, width="stretch")
        
        with st.expander("📄 View Vacancy Data Table", expanded=False):
            st.subheader("Data Table")
            display_df_with_changes(df_vac, is_percent=True)

# 5. Net Absorption 
with tabs[5]:
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
            st.plotly_chart(fig_abs, width="stretch")
        
        # Frame the table
        with st.expander("📄 View Absorption Data Table", expanded=False):
            st.subheader("Data Table")
            display_df_with_changes(df_abs)

# 6. Rent Performance 
with tabs[6]:
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
            st.plotly_chart(fig2, width="stretch")
        
        # Frame the table
        with st.expander("📄 View Rent Data Table", expanded=False):
            st.subheader("Data Table")
            display_df_with_changes(df_rent, is_percent=False)

# 7. Capital Markets
with tabs[7]:
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
                st.plotly_chart(fig_cv, width="stretch")
            
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
                st.plotly_chart(fig_rate, width="stretch")
            
            # Frame the right table
            with st.expander("📄 Cap Rate Table", expanded=False):
                display_df_with_changes(df_rate, is_percent=True)

# 8. Transactions
with tabs[8]:
    st.header("Major Transactions Analysis")
    df_cap_raw = load_table("capital_markets")
    
    if df_cap_raw is not None and not df_cap_raw.empty:
        df_cap_raw['Consideration_Num'] = df_cap_raw['Consideration'].astype(str).str.replace(',', '').apply(pd.to_numeric, errors='coerce')
        
        # --- NEW MAP VISUALIZATION (WITH SHADED REGIONS) ---
      # --- NEW MAP VISUALIZATION (WITH SHADED REGIONS) ---
        st.subheader("📍 Transaction Map & Submarket Regions")
        if "Latitude" in df_cap_raw.columns and "Longitude" in df_cap_raw.columns:
            df_cap_raw["Latitude"] = pd.to_numeric(df_cap_raw["Latitude"], errors="coerce")
            df_cap_raw["Longitude"] = pd.to_numeric(df_cap_raw["Longitude"], errors="coerce")
            map_df = df_cap_raw.dropna(subset=["Latitude", "Longitude", "Consideration_Num"])
            
            if not map_df.empty:
                
                # 1. Base Scatter Map
                fig_map = px.scatter_map(
                    map_df, lat="Latitude", lon="Longitude", 
                    color="Subdistrict", size="Consideration_Num", 
                    hover_name="Property", 
                    hover_data={"Quarter": True, "Consideration": True, "Latitude": False, "Longitude": False},
                    zoom=11, center={"lat": 37.54, "lon": 126.98}, 
                    map_style="carto-positron", # <--- Fixed to the dark map layer! (Change to "carto-positron" if you want the light map)
                    height=800,
                    color_discrete_map=REGION_COLORS
                )
                
                # 2. Add the Shaded Background Regions
                seoul_geo = fetch_seoul_geojson()
                map_layers = [] # Updated variable name for clarity
                
                if seoul_geo:
                    # CBD: 종로구, 중구, 용산구, 서대문구 (Red)
                    cbd_feat = [f for f in seoul_geo['features'] if f['properties']['name'] in ['종로구', '중구', '용산구', '서대문구']]
                    if cbd_feat:
                        map_layers.append({"sourcetype": "geojson", "source": {"type": "FeatureCollection", "features": cbd_feat}, "type": "fill", "color": "rgba(231, 76, 60, 0.2)"})
                    
                    # GBD: 강남구, 서초구, 송파구 (Blue)
                    gbd_feat = [f for f in seoul_geo['features'] if f['properties']['name'] in ['강남구', '서초구', '송파구']]
                    if gbd_feat:
                        map_layers.append({"sourcetype": "geojson", "source": {"type": "FeatureCollection", "features": gbd_feat}, "type": "fill", "color": "rgba(52, 152, 219, 0.2)"})
                    
                    # YBD: 영등포구 (Green)
                    ybd_feat = [f for f in seoul_geo['features'] if f['properties']['name'] in ['영등포구']]
                    if ybd_feat:
                        map_layers.append({"sourcetype": "geojson", "source": {"type": "FeatureCollection", "features": ybd_feat}, "type": "fill", "color": "rgba(46, 204, 113, 0.2)"})

                # Apply layers and styling (Updated mapbox_layers to map_layers)
                fig_map.update_layout(map_layers=map_layers, margin={"r":0,"t":0,"l":0,"b":0})
                fig_map.update_traces(marker=dict(opacity=0.9, sizemin=7)) 
                
                # Updated use_container_width here too!
                st.plotly_chart(fig_map, width="stretch")
            else:
                st.info("No valid latitude/longitude coordinates found to map.")
        
        st.markdown("---")
        
        st.subheader("Major Transactions Timeline")
        fig_scatter = px.scatter(
            df_cap_raw.dropna(subset=['Consideration_Num']), 
            x="Quarter", y="Consideration_Num", color="Subdistrict", 
            size="Consideration_Num", hover_data=["Property", "TransactedGFApy"],
            title="Transaction Value by Quarter (Hover for GFA in Pyeong)", # Updated title
            color_discrete_map={"CBD": "#e74c3c", "GBD": "#3498db", "YBD": "#2ecc71", "ETC": "#f1c40f"},
            labels={"TransactedGFApy": "Transacted GFA (Pyeong)", "Consideration_Num": "Consideration (KRW)"} # Cleans up the hover box text
        )
        st.plotly_chart(fig_scatter, width="stretch")
        
        with st.expander("View Raw Transaction Data"):
            st.dataframe(df_cap_raw.drop(columns=['Consideration_Num']), width='stretch', hide_index=True)

# 9. News Tab
with tabs[9]:
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

# 10. Admin Panel
with tabs[10]:
    st.header("🛠️ Database Admin Panel")
    st.subheader("1. Edit Data (Rows)")
    
    # Inside Tab 9 (Admin)
    table_to_edit = st.selectbox(
        "Select Table to Edit:", 
        [
            "vacancy", "rent", "macro", "existing_supply", "future_supply", 
            "net_absorption", "capital_value", "cap_rate", "capital_markets"
        ]
    )
    
    conn = sqlite3.connect(DB_NAME)
    df_current = pd.read_sql(f"SELECT * FROM {table_to_edit}", conn)
    conn.close()
    
    edited_df = st.data_editor(df_current, num_rows="dynamic", width="stretch", height=350)
    
    if st.button(f"💾 Save Data Changes to '{table_to_edit}'"):
        try:
            conn = sqlite3.connect(DB_NAME)
            edited_df.to_sql(table_to_edit, conn, if_exists="replace", index=False)
            conn.close()
            st.cache_data.clear()
            st.success(f"✅ Data saved successfully!")
            st.rerun()
        except Exception as e: st.error(f"❌ Error saving data: {e}")

    st.markdown("---")
    st.subheader("2. Modify Database Structure (Columns)")
    safe_columns = [c for c in df_current.columns if c.lower() not in ['quarter', 'indicator', 'year']]
    struct_tabs = st.tabs(["➕ Add Column", "✏️ Rename Column", "🗑️ Delete Column"])
    
    with struct_tabs[0]:
        with st.form("add_column_form"):
            new_col_name = st.text_input("New Column Name")
            new_col_type = st.selectbox("Data Type", ["REAL (Numbers / Decimals)", "TEXT (Words / Dates)"])
            if st.form_submit_button("🏗️ Build New Column"):
                if new_col_name and new_col_name not in df_current.columns:
                    try:
                        clean_name = "".join(e for e in new_col_name if e.isalnum() or e == '_')
                        sql_type = "REAL" if "REAL" in new_col_type else "TEXT"
                        conn = sqlite3.connect(DB_NAME)
                        conn.execute(f'ALTER TABLE {table_to_edit} ADD COLUMN "{clean_name}" {sql_type}')
                        conn.commit(); conn.close(); st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(e)
                else: st.warning("Invalid or duplicate name.")

    with struct_tabs[1]:
        with st.form("rename_column_form"):
            col_to_rename = st.selectbox("Select Column to Rename:", safe_columns)
            new_name = st.text_input("Type New Name:")
            if st.form_submit_button("✏️ Rename Column"):
                if new_name and col_to_rename:
                    try:
                        clean_name = "".join(e for e in new_name if e.isalnum() or e == '_')
                        conn = sqlite3.connect(DB_NAME)
                        conn.execute(f'ALTER TABLE {table_to_edit} RENAME COLUMN "{col_to_rename}" TO "{clean_name}"')
                        conn.commit(); conn.close(); st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(e)

    with struct_tabs[2]:
        with st.form("delete_column_form"):
            col_to_delete = st.selectbox("Select Column to Delete:", safe_columns)
            confirm_delete = st.checkbox(f"Confirm deletion of '{col_to_delete}'.")
            if st.form_submit_button("🗑️ Delete Column") and col_to_delete and confirm_delete:
                try:
                    conn = sqlite3.connect(DB_NAME)
                    conn.execute(f'ALTER TABLE {table_to_edit} DROP COLUMN "{col_to_delete}"')
                    conn.commit(); conn.close(); st.cache_data.clear(); st.rerun()
                except Exception as e: st.error(e)