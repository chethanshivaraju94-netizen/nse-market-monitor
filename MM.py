import yfinance as yf
import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule

print("--- NSE Market Monitor (Endless Archive Engine) ---")
file_name = "NSE_Market_Monitor.xlsx"

# 1. Define the exact layout requested
headers = [
    "Date", "Up 4% Today", "Down 4% Today", "5 Day Ratio", "10 Day Ratio", 
    "Advances", "Declines", "A/D Ratio", 
    "> 200 SMA (%)", "> 50 SMA (%)", "> 20 EMA (%)", "> 10 EMA (%)"
]

# 2. Fetch the Nifty Total Market Index list
print("Fetching Nifty Total Market universe list...")
try:
    url = "https://archives.nseindia.com/content/indices/ind_niftytotalmarket_list.csv"
    nifty_total_df = pd.read_csv(url, storage_options={'User-Agent': 'Mozilla/5.0'})
    tickers = [str(symbol) + ".NS" for symbol in nifty_total_df['Symbol']]
except Exception as e:
    print(f"Error fetching live list: {e}. Falling back to standard list.")
    tickers = ["RELIANCE.NS", "TCS.NS", "INFY.NS", "HDFCBANK.NS", "ICICIBANK.NS", "SBIN.NS"]

# 3. Download data (We download 1 year to ensure today's 200 SMA is mathematically perfect)
print("Downloading data from Yahoo Finance (this may take 1-2 minutes)...")
data = yf.download(tickers, period="2y", multi_level_index=False)['Close']
data = data.dropna(how='all', axis=1)

print("Calculating daily breadth and technical matrices...")
daily_returns_df = data.pct_change() * 100
sma_200_df = data.rolling(window=200).mean()
sma_50_df = data.rolling(window=50).mean()
ema_20_df = data.ewm(span=20, adjust=False).mean()
ema_10_df = data.ewm(span=10, adjust=False).mean()

# 4. Generate the breadth math for the most recent 65 days
# (This ensures accurate rolling 5/10 math and handles the format upgrade seamlessly)
history_data = []
lookback_days = 65 
if len(data) < lookback_days: lookback_days = len(data) - 1

for i in range(-lookback_days, 0):
    date_str = data.index[i].strftime("%Y-%m-%d")
    day_close = data.iloc[i]
    day_returns = daily_returns_df.iloc[i]
    
    up_4 = int((day_returns >= 4.0).sum())
    down_4 = int((day_returns <= -4.0).sum())
    
    advances = int((day_returns > 0).sum())
    declines = int((day_returns < 0).sum())
    ad_ratio = round((advances / declines), 2) if declines > 0 else float(advances)
    
    sma_200 = sma_200_df.iloc[i]
    valid_200 = sma_200.notna().sum() 
    pct_200 = round(((day_close > sma_200).sum() / valid_200) * 100, 2) if valid_200 > 0 else 0.0

    sma_50 = sma_50_df.iloc[i]
    valid_50 = sma_50.notna().sum()
    pct_50 = round(((day_close > sma_50).sum() / valid_50) * 100, 2) if valid_50 > 0 else 0.0

    ema_20 = ema_20_df.iloc[i]
    valid_20 = ema_20.notna().sum()
    pct_20 = round(((day_close > ema_20).sum() / valid_20) * 100, 2) if valid_20 > 0 else 0.0

    ema_10 = ema_10_df.iloc[i]
    valid_10 = ema_10.notna().sum()
    pct_10 = round(((day_close > ema_10).sum() / valid_10) * 100, 2) if valid_10 > 0 else 0.0
    
    history_data.append({
        "Date": date_str, "Up 4% Today": up_4, "Down 4% Today": down_4, 
        "Advances": advances, "Declines": declines, "A/D Ratio": ad_ratio, 
        "> 200 SMA (%)": pct_200, "> 50 SMA (%)": pct_50, "> 20 EMA (%)": pct_20, "> 10 EMA (%)": pct_10
    })

df_recent = pd.DataFrame(history_data)

# Process Rolling Ratios on the chronological data
df_recent['5 Day Ratio'] = (df_recent['Up 4% Today'].rolling(5).sum() / df_recent['Down 4% Today'].rolling(5).sum().replace(0, 1)).round(2)
df_recent['10 Day Ratio'] = (df_recent['Up 4% Today'].rolling(10).sum() / df_recent['Down 4% Today'].rolling(10).sum().replace(0, 1)).round(2)

# Sort newest to top
df_recent = df_recent.sort_values(by="Date", ascending=False)

# 5. The "Smart Merge" - Stack new data on top of the endless historical archive
print("Merging with historical archive...")
if os.path.exists(file_name):
    try:
        df_archive = pd.read_excel(file_name)
        # Combine data and use drop_duplicates to automatically overwrite weekend/manual runs with clean data!
        df_combined = pd.concat([df_recent, df_archive], ignore_index=True)
        df_combined = df_combined.drop_duplicates(subset=['Date'], keep='first')
        
        # Ensure the final output matches your requested layout
        for col in headers:
            if col not in df_combined.columns:
                df_combined[col] = "" # Handles old files missing the ratio columns gracefully
        df_final = df_combined[headers]
    except Exception as e:
        print(f"Archive read error (normal if upgrading format): {e}")
        df_final = df_recent[headers]
else:
    df_final = df_recent[headers]

# Final safety sort (Newest date firmly at the top)
df_final = df_final.sort_values(by="Date", ascending=False)
df_final = df_final.fillna("") # Clean any Pandas NaN artifacts

# 6. Rebuild Excel and Apply Conditional Formatting
if os.path.exists(file_name):
    os.remove(file_name)

wb = Workbook()
ws = wb.active
ws.title = "Market Monitor"

ws.append(headers)

for index, row in df_final.iterrows():
    # Clean up the date string format for the Excel sheet
    date_val = str(row['Date'])[:10] if str(row['Date']) != "" else ""
    ws.append([
        date_val, row['Up 4% Today'], row['Down 4% Today'], row['5 Day Ratio'], row['10 Day Ratio'],
        row['Advances'], row['Declines'], row['A/D Ratio'], 
        row['> 200 SMA (%)'], row['> 50 SMA (%)'], row['> 20 EMA (%)'], row['> 10 EMA (%)']
    ])

max_row = ws.max_row
ws.conditional_formatting._cf_rules = {}

# Layout mapped formatting
green_scale = ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='63BE7B')
red_scale = ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='F8696B')
ma_scale = ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=50, mid_color='FFFFFF', end_type='num', end_value=100, end_color='63BE7B')
ratio_scale = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='num', mid_value=1.0, mid_color='FFFFFF', end_type='max', end_color='63BE7B')

ws.conditional_formatting.add(f"B2:B{max_row}", green_scale) # Up 4%
ws.conditional_formatting.add(f"C2:C{max_row}", red_scale)   # Down 4%
ws.conditional_formatting.add(f"D2:E{max_row}", ratio_scale) # 5 & 10 Day Ratios
ws.conditional_formatting.add(f"F2:F{max_row}", green_scale) # Advances
ws.conditional_formatting.add(f"G2:G{max_row}", red_scale)   # Declines
ws.conditional_formatting.add(f"H2:H{max_row}", ratio_scale) # A/D Ratio
ws.conditional_formatting.add(f"I2:L{max_row}", ma_scale)    # Moving Averages

wb.save(file_name)
print(f"\n--- SUCCESS! Endlessly growing file saved perfectly to {file_name}. ---")
