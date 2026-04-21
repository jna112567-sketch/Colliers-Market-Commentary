import pandas as pd
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import re
import warnings
warnings.filterwarnings("ignore")

def parse_quarter(q_str):
    """Parse 'Q1 2019' or '1Q 2019' into a sortable tuple (2019, 1)."""
    q_str = str(q_str).strip().upper()
    match = re.search(r'(?:Q(\d)\s*(\d{4})|(\d)Q\s*(\d{4}))', q_str)
    if match:
        if match.group(1):
            return int(match.group(2)), int(match.group(1))
        else:
            return int(match.group(4)), int(match.group(3))
    # Fallback if no match
    return 9999, 99

def format_next_quarter(last_q_tuple):
    """Return next quarter string, e.g., 'Q2 2024'."""
    year, q = last_q_tuple
    next_q = q + 1
    next_year = year
    if next_q > 4:
        next_q = 1
        next_year += 1
    return f"Q{next_q} {next_year}"

class MarketForecaster:
    @staticmethod
    def forecast_next_quarter(df, value_col):
        """
        Takes a dataframe with 'Quarter' and a specific value_col.
        Returns the predicted value for the next quarter.
        """
        if df is None or df.empty or value_col not in df.columns:
            return None
        
        # Clean and sort data
        df_clean = df.copy()
        df_clean['Q_Tuple'] = df_clean['Quarter'].apply(parse_quarter)
        df_clean = df_clean.sort_values('Q_Tuple')
        
        # Extract values
        y = pd.to_numeric(df_clean[value_col].astype(str).str.replace(',', '').str.replace('%', ''), errors='coerce').dropna().values
        
        if len(y) < 4:
            # Not enough data for exponential smoothing, use naive approach (last value)
            return y[-1] if len(y) > 0 else None
            
        try:
            # Simple Exponential Smoothing (Holt's Linear) since data length is short
            model = ExponentialSmoothing(y, trend='add', seasonal=None, initialization_method="estimated")
            fit_model = model.fit()
            forecast = fit_model.forecast(1)[0]
            return float(forecast)
        except Exception as e:
            # Fallback to moving average if model fails
            return float(np.mean(y[-3:]))
            
    @staticmethod
    def add_forecast_to_df(df, cols_to_forecast):
        """
        Takes a dataframe and appends one row representing the forecasted Q+1.
        Returns the appended dataframe.
        """
        if df is None or df.empty or 'Quarter' not in df.columns:
            return df
            
        df_clean = df.copy()
        df_clean['Q_Tuple'] = df_clean['Quarter'].apply(parse_quarter)
        df_clean = df_clean.sort_values('Q_Tuple')
        
        last_tuple = df_clean['Q_Tuple'].iloc[-1]
        next_q_str = format_next_quarter(last_tuple)
        
        forecast_row = {'Quarter': next_q_str, 'Is_Forecast': True}
        
        for col in cols_to_forecast:
            if col in df.columns:
                val = MarketForecaster.forecast_next_quarter(df_clean, col)
                if val is not None:
                    forecast_row[col] = val
                    
        df_forecast = pd.DataFrame([forecast_row])
        df_clean['Is_Forecast'] = False
        
        res = pd.concat([df_clean, df_forecast], ignore_index=True)
        res = res.drop(columns=['Q_Tuple'])
        return res


