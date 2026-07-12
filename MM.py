import yfinance as yf
import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import ColorScaleRule

print("--- NSE Market Monitor (Total Market + A/D Breadth) ---")

# 1. Fetch the Nifty Total Market Index list
print("Fetching Nifty Total Market universe list...")
try:
    url = "https://archives.nseindia.com/content/indices/ind_niftytotalmarket_list.csv"
    nifty_total_df = pd.read_csv(url, storage_options={'User-Agent': 'Mozilla/5.0'})
    tickers = [str(symbol) + ".NS" for symbol in nifty_total_df['Symbol']]
    print(f"Successfully loaded {len(tickers)} stocks.")
except Exception as e:
    print(f"Error fetching live list: {e}. Falling back to standard list.")
    tickers = ["RELIANCE.NS", "TCS.NS", "INFY.NS", "HDFCBANK.NS", "ICICIBANK.NS", "SBIN.NS"]

# 2. Download historical data
print("Downloading 2 years of historical data from Yahoo Finance (this may take 1-2 minutes)...")
data = yf.download(tickers, period="2y", multi_level_index=False)['Close']
data = data.dropna(how='all', axis=1)

today_close = data.iloc[-1]

# 3. Calculate Breadth Metrics (4% moves & A/D Line)
daily_returns = data.pct_change() * 100
today_returns = daily_returns.iloc[-1]

# Primary Breadth
up_4_percent = int((today_returns >= 4.0).sum())
down_4_percent = int((today_returns <= -4.0).sum())

# Advance/Decline Line
advances = int((today_returns > 0).sum())
declines = int((today_returns < 0).sum())

# Calculate A/D Ratio (Handle division by zero safely if no stocks decline)
ad_ratio = round((advances / declines), 2) if declines > 0 else float(advances)

# 4. Calculate Moving Average Percentages
sma_200_series = data.rolling(window=200).mean().iloc[-1]
valid_200_count = sma_200_series.notna().sum() 
pct_above_200 = round(((today_close > sma_200_series).sum() / valid_200_count) * 100, 2) if valid_200_count > 0 else 0.0

sma_50_series = data.rolling(window=50).mean().iloc[-1]
valid_50_count = sma_50_series.notna().sum()
pct_above_50 = round(((today_close > sma_50_series).sum() / valid_50_count) * 100, 2) if valid_50_count > 0 else 0.0

ema_20_series = data.ewm(span=20, adjust=False).mean().iloc[-1]
valid_20_count = ema_20_series.notna().sum()
pct_above_20 = round(((today_close > ema_20_series).sum() / valid_20_count) * 100, 2) if valid_20_count > 0 else 0.0

ema_10_series = data.ewm(span=10, adjust=False).mean().iloc[-1]
valid_10_count = ema_10_series.notna().sum()
pct_above_10 = round(((today_close > ema_10_series).sum() / valid_10_count) * 100, 2) if valid_10_count > 0 else 0.0

date_str = datetime.now().strftime("%Y-%m-%d")

# 5. Excel File Operations
file_name = "NSE_Market_Monitor.xlsx"

if not os.path.exists(file_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "Market Monitor"
    # Updated 10-Column Header
    ws.append(["Date", "Advances", "Declines", "A/D Ratio", "Up 4% Today", "Down 4% Today", "> 200 SMA (%)", "> 50 SMA (%)", "> 20 EMA (%)", "> 10 EMA (%)"])
else:
    wb = load_workbook(file_name)
    ws = wb.active

# Append today's data row
ws.append([date_str, advances, declines, ad_ratio, up_4_percent, down_4_percent, pct_above_200, pct_above_50, pct_above_20, pct_above_10])

# 6. Apply Conditional Formatting (Color Scales)
ws.conditional_formatting._cf_rules = {}
max_row = ws.max_row

# Reusable Color Scales
green_scale = ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='63BE7B')
red_scale = ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='F8696B')
ma_scale = ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=50, mid_color='FFFFFF', end_type='num', end_value=100, end_color='63BE7B')
# A/D Ratio Scale: < 1.0 is Red, 1.0 is White, > 2.0 is Green
ad_scale = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='num', mid_value=1.0, mid_color='FFFFFF', end_type='max', end_color='63BE7B')

# Apply Rules to matching columns
ws.conditional_formatting.add(f"B2:B{max_row}", green_scale) # Advances
ws.conditional_formatting.add(f"C2:C{max_row}", red_scale)   # Declines
ws.conditional_formatting.add(f"D2:D{max_row}", ad_scale)    # A/D Ratio
ws.conditional_formatting.add(f"E2:E{max_row}", green_scale) # Up 4%
ws.conditional_formatting.add(f"F2:F{max_row}", red_scale)   # Down 4%
ws.conditional_formatting.add(f"G2:J{max_row}", ma_scale)    # Moving Averages (200 down to 10)

# Save the final styled Excel file
wb.save(file_name)

print(f"\n--- SUCCESS! ---")
print(f"Added entry for {date_str} to {file_name}")
print(f"Advances: {advances} | Declines: {declines} | A/D Ratio: {ad_ratio}")
print(f"Up 4%: {up_4_percent} | Down 4%: {down_4_percent}")