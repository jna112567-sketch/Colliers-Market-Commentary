import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.request
import xml.etree.ElementTree as ET

# ---------------------------------------------------------
# 1. APP CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(page_title="Seoul Office Market Dashboard", layout="wide")
st.title("🏢 Seoul Office Grade A Market Report")

st.sidebar.header("Upload Database")
uploaded_file = st.sidebar.file_uploader("Upload Mapletree Workbook (Excel)", type=["xlsx", "xls"])

# ---------------------------------------------------------
# 2. HELPER FUNCTIONS FOR DATA CLEANING
# ---------------------------------------------------------
@st.cache_data
def load_excel_data(file):
    xls = pd.ExcelFile(file)
    sheet_dict = {}
    for sheet_name in xls.sheet_names:
        sheet_dict[sheet_name] = pd.read_excel(file, sheet_name=sheet_name)
    return sheet_dict

def format_general_table(df, rows=20):
    """Smart Formatter: Finds the real table header, cleans data, and fixes duplicate column names!"""
    clean_df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    if clean_df.empty: return clean_df
    
    # Scan the top 10 rows to find the row with the most columns (the true header)
    best_row_idx = 0
    max_non_nulls = 0
    for i in range(min(10, len(clean_df))):
        non_nulls = clean_df.iloc[i].notna().sum()
        if non_nulls > max_non_nulls:
            max_non_nulls = non_nulls
            best_row_idx = i
            
    # Grab the raw header names
    raw_columns = clean_df.iloc[best_row_idx].fillna("").astype(str).tolist()
    
    # Fix duplicates: Streamlit (PyArrow) crashes if multiple columns have the same name (like blank spaces)
    seen = {}
    unique_cols = []
    for col in raw_columns:
        col = col.strip()
        if col == "": 
            col = "Unnamed" # Give blank columns a default name
            
        if col in seen:
            seen[col] += 1
            unique_cols.append(f"{col} {seen[col]}") # E.g., Unnamed 1, Unnamed 2
        else:
            seen[col] = 0
            unique_cols.append(col)
            
    # Set the deduplicated names as the official columns
    clean_df.columns = unique_cols
    clean_df = clean_df.iloc[best_row_idx + 1:]
    
    return clean_df.head(rows).fillna("")

def extract_macro_table(df):
    try:
        ind_row = None
        for r in range(10):
            for c in range(10):
                if str(df.iloc[r, c]).strip() == 'Indicator':
                    ind_row = r
                    break
            if ind_row is not None: break
        if ind_row is None: return None
        header = df.iloc[ind_row].values
        macro_df = df.iloc[ind_row+1 : ind_row+10].copy()
        macro_df.columns = header
        macro_df = macro_df[macro_df['Indicator'].notna()]
        return macro_df.dropna(axis=1, how='all')
    except: return None

def extract_timeseries_table(df):
    try:
        first_col = df.iloc[:, 0].fillna("").astype(str).str.strip()
        cbd_indexes = first_col[first_col == 'CBD'].index
        if len(cbd_indexes) == 0: return None
        all_data = []
        for idx in cbd_indexes:
            header_idx = idx - 1 
            clean_df = df.iloc[header_idx : idx + 5].copy()
            clean_df.columns = clean_df.iloc[0].fillna("").astype(str).str.strip()
            clean_df = clean_df.iloc[1:] 
            clean_df.set_index(clean_df.columns[0], inplace=True)
            clean_df.index = clean_df.index.fillna("").astype(str).str.strip()
            valid_rows = [r for r in clean_df.index if any(sub in r for sub in ['CBD', 'GBD', 'YBD', 'Overall', 'Seoul'])]
            clean_df = clean_df.loc[valid_rows].dropna(axis=1, how='all')
            ts_data = clean_df.T
            ts_data.index.name = "Quarter"
            ts_data = ts_data.apply(pd.to_numeric, errors='coerce')
            valid_idx = (ts_data.index.fillna("").astype(str).str.strip() != "") & \
                        (ts_data.index.fillna("").astype(str).str.len() < 20)
            ts_data = ts_data[valid_idx]
            if not ts_data.empty: all_data.append(ts_data)
        if not all_data: return None
        combined = pd.concat(all_data)
        rename_dict = {c: 'Overall' for c in combined.columns if 'Seoul' in c or 'Overall' in c}
        combined.rename(columns=rename_dict, inplace=True)
        return combined[~combined.index.duplicated(keep='last')].reset_index()
    except: return None

# --- LIVE NEWS FETCHER ---
@st.cache_data(ttl=3600) # Caches the news for 1 hour so it doesn't slow down your app!
def fetch_korean_office_news():
    """Fetches real-time news from Google News RSS."""
    # Search query: "서울 오피스 빌딩" (Seoul Office Building)
    url = "https://news.google.com/rss/search?q=%EC%84%9C%EC%9A%B8+%EC%98%A4%ED%94%BC%EC%8A%A4+%EB%B9%8C%EB%94%A9&hl=ko&gl=KR&ceid=KR:ko"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        news_items = []
        # Grab the top 5 articles
        for item in root.findall('./channel/item')[:5]:
            title = item.find('title').text
            link = item.find('link').text
            pubDate = item.find('pubDate').text
            date_str = pubDate[5:16] # Clean up the date format
            news_items.append(f"### 📰 [{date_str}] **[{title}]({link})**")
        return "\n\n---\n\n".join(news_items)
    except Exception as e:
        return f"뉴스를 불러오는 데 실패했습니다. (Error: {e})"

# ---------------------------------------------------------
# 3. DASHBOARD LOGIC
# ---------------------------------------------------------
tabs = st.tabs(["1. Macro", "2. Supply", "3. Future", "4. Vacancy", "5. Absorption", "6. Rent", "7. Capital", "8. Outlook", "9. News"])

if uploaded_file is not None:
    data = load_excel_data(uploaded_file)
    
    if data:
        # --- Tab 1: Macroeconomics ---
        with tabs[0]:
            st.header("Macroeconomics")
            head_sheet = next((s for s in data.keys() if 'Headline' in s), None)
            if head_sheet:
                macro_table = extract_macro_table(data[head_sheet])
                if macro_table is not None: st.dataframe(macro_table.astype(str), width='stretch')

        # --- Tab 2: Existing Supply ---
        with tabs[1]:
            st.header("Existing Supply")
            supply_sheet = next((s for s in data.keys() if 'Existing' in s), None)
            if supply_sheet: st.dataframe(format_general_table(data[supply_sheet]), width='stretch')

        # --- Tab 3: Future Supply ---
        with tabs[2]:
            st.header("Future Supply")
            future_sheet = next((s for s in data.keys() if 'Future' in s), None)
            if future_sheet: st.dataframe(format_general_table(data[future_sheet]), width='stretch')

        # --- Tab 4: Vacancy ---
        with tabs[3]:
            st.header("Vacancy Rate Trend")
            vac_sheet = next((s for s in data.keys() if 'Vacancy' in s), None)
            if vac_sheet:
                clean_vac = extract_timeseries_table(data[vac_sheet])
                if clean_vac is not None:
                    expected_cols = ["CBD", "GBD", "YBD", "Overall"]
                    available_cols = [col for col in expected_cols if col in clean_vac.columns]
                    clean_vac = clean_vac.dropna(subset=available_cols, how='all') 
                    fig = px.line(clean_vac, x="Quarter", y=available_cols, markers=True)
                    fig.update_layout(yaxis_tickformat='.1%')
                    st.plotly_chart(fig, width="stretch")

        # --- Tab 5: Net Absorption ---
        with tabs[4]:
            st.header("Net Absorption")
            abs_sheet = next((s for s in data.keys() if 'Absorption' in s), None)
            if abs_sheet: st.dataframe(format_general_table(data[abs_sheet]), width='stretch')

        # --- Tab 6: Rent Performance ---
        with tabs[5]:
            st.header("Rent Performance Trend")
            rent_sheet = next((s for s in data.keys() if 'Rent' in s), None)
            if rent_sheet:
                clean_rent = extract_timeseries_table(data[rent_sheet])
                if clean_rent is not None:
                    expected_cols = ["CBD", "GBD", "YBD", "Overall"]
                    available_cols = [col for col in expected_cols if col in clean_rent.columns]
                    clean_rent = clean_rent.dropna(subset=available_cols, how='all')
                    fig2 = px.line(clean_rent, x="Quarter", y=available_cols, markers=True)
                    st.plotly_chart(fig2, width="stretch")

        # --- Tab 7: Capital Markets ---
        with tabs[6]:
            st.header("Capital Markets (Yield & Capital Value)")
            cap_sheet = next((s for s in data.keys() if 'CV' in s or 'Yield' in s), None)
            if cap_sheet: st.dataframe(format_general_table(data[cap_sheet]), width='stretch')

        # --- Tab 8: Outlook ---
        with tabs[7]:
            st.header("Market Outlook")
            outlook_sheet = next((s for s in data.keys() if 'Outlook' in s), None)
            if outlook_sheet:
                try: outlook_text = str(data[outlook_sheet].iloc[0, 1])
                except: outlook_text = "No commentary found."
                st.text_area("Copy and Paste this Commentary:", value=outlook_text, height=300)

else:
    st.info("👈 Please upload your Mapletree Workbook in the sidebar.")

# --- Tab 9: News Source ---
with tabs[8]:
    st.header("📰 서울 오피스 시장 실시간 뉴스 (Live News)")
    st.write("Google News RSS 피드를 통해 자동으로 수집된 최신 기사입니다.")
    st.markdown(fetch_korean_office_news())