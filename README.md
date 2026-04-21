# Colliers Market Commentary Dashboard

An enterprise-grade, highly interactive Streamlit dashboard designed to streamline real estate market analysis, automate commentary generation, and manage market statistics (Macro, Office, Logistics) seamlessly.

## 🚀 Key Features

### 1. 📑 Automated Executive Summary (AI + Web Grounding)
- Uses the **Google Gemini API** to digest historical market data alongside live web searches (Bank of Korea, Ministry of Economy and Finance, and general market news).
- Automatically synthesizes a structured, institutional-grade market commentary explaining the "why" behind the numbers, saving hours of manual drafting.

### 2. 📈 Time-Series Market Forecasting
- Integrates a local forecasting model using **Holt's Linear Exponential Smoothing** (`statsmodels`).
- Automatically analyzes historical trends to predict the next quarter's Vacancy Rates, Net Absorption, and Face Rents.
- These projections are plotted distinctly on the visual charts and fed back into the AI to help generate forward-looking executive summaries.

### 3. 🤖 Offline Market Data Extractor
- A lightweight, **100% offline** PDF extraction engine built with `pdfplumber` and Regex pattern matching.
- **Workflow**: Upload multiple consultancy reports or leasing flyers (PDFs) for a specific quarter. The engine instantly extracts the estimated *Overall Vacancy Rate* and *Average Face Rent* across all documents.
- Provides a unified UI to review these metrics and a 1-click button to securely commit them into the local Excel database.
- *Privacy First*: Since this uses no APIs or external servers, your sensitive internal reports never leave your machine.

### 4. 🏢 Comprehensive Visualizations
- Detailed tabs tracking **Existing Supply**, **Future Pipeline**, **Vacancy**, **Absorption**, **Rent**, and **Capital Markets**.
- Includes interactive, enterprise-formatted **Plotly** charts (with high-res export options) and foldable data tables for quick deep-dives.
- **Transactions Map**: A specialized, interactive scatter map plotting major capital market transactions across subdistricts (CBD, GBD, YBD) using geo-coordinates.

### 5. 📰 Live Dynamic News Feed
- Integrates a dynamic RSS feed tailored to the specific asset class (Office, Logistics, or Macro).
- Filter recent news articles by category (M&A, Leasing, Development) or region (CBD, GBD, etc.) to stay updated on market shifts.

---

## 🛠 Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone <repository_url>
   cd Colliers-Market-Commentary
   ```

2. **Install the dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure Secrets:**
   Create a `.streamlit/secrets.toml` file in the root directory and add the following:
   ```toml
   APP_PASSWORD = "your_secure_dashboard_password"
   ECOS_API_KEY = "your_bank_of_korea_api_key"
   GEMINI_API_KEY = "your_google_gemini_api_key"
   ```

4. **Prepare the Data Sources:**
   Ensure the following Excel files exist in the root directory (formatted with the required sheets like `vacancy`, `rent`, `capital_markets`, etc.):
   - `seoul_office_data.xlsx`
   - `seoul_logistics_data.xlsx`

5. **Run the Dashboard:**
   ```bash
   streamlit run app.py
   ```

## 🔐 Security
The dashboard includes an authentication layer. Users must enter the `APP_PASSWORD` defined in the secrets to access the internal analytics and AI tools.