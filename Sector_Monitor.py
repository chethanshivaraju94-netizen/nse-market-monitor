import pandas as pd
import numpy as np
from tvDatafeed import TvDatafeed, Interval
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.formatting.rule import ColorScaleRule
import warnings

# Suppress the unauthenticated session warning from tvDatafeed
warnings.filterwarnings("ignore")

print("--- NSE Sector Rotation Engine Initiated (TradingView Data) ---")

# 1. Initialize TradingView Connection
tv = TvDatafeed()

file_name = "NSE_Sector_Monitor.xlsx"
benchmark_ticker = "CNX500" # Nifty 500

# Updated dictionary using strict TradingView NSE symbol conventions
tickers = {
    "IT": "CNXIT",
    "Bank": "BANKNIFTY",
    "Auto": "CNXAUTO",
    "FMCG": "CNXFMCG",
    "Pharma": "CNXPHARMA",
    "Metal": "CNXMETAL",
    "Energy": "CNXENERGY",
    "Realty": "CNXREALTY",
    "Fin Services": "CNXFIN",
    "Infrastructure": "CNXINFRA",
    "Consumption": "CNXCONSUM",
    "Commodities": "CNXCMDT",
    "PSE": "CNXPSE",
    "MNC": "CNXMNC",
    "Media": "CNXMEDIA",
    "PSU Bank": "CNXPSUBANK",
    "Healthcare": "NIFTY_HEALTHCARE",
    "Defence": "MODEFENCE"
}

def fetch_data(symbol, n_bars=300):
    try:
        # Requesting daily interval bars directly from the NSE exchange logic
        df = tv.get_hist(symbol=symbol, exchange='NSE', interval=Interval.in_daily, n_bars=n_bars)
        if df is not None and not df.empty:
            return df[['close']]
        return None
    except Exception as e:
        print(f"Error fetching {symbol}: {e}")
        return None

# 2. Fetch Benchmark
print(f"Fetching Benchmark: {benchmark_ticker}")
bench_df = fetch_data(benchmark_ticker)
if bench_df is None:
    raise ValueError("Failed to fetch benchmark data. Exiting.")

results = []

# 3. Process Sector Data & Calculate Indicators
for sector, symbol in tickers.items():
    print(f"Fetching Sector: {sector}")
    sector_df = fetch_data(symbol)
    if sector_df is None:
        continue
        
    # Align sector dates perfectly with the benchmark
    df = pd.merge(sector_df, bench_df, left_index=True, right_index=True, suffixes=('', '_bench'))
    
    # Calculate Base RS Line
    df['RS_Line'] = df['close'] / df['close_bench']
    
    # Calculate Institutional Momentum (Lookbacks)
    df['5D_RS'] = df['RS_Line'].pct_change(periods=5) * 100
    df['21D_RS'] = df['RS_Line'].pct_change(periods=21) * 100
    df['65D_RS'] = df['RS_Line'].pct_change(periods=65) * 100
    
    # Standard Exponential and Simple Moving Averages for Absolute Trend
    df['10_EMA'] = df['close'].ewm(span=10, adjust=False).mean()
    df['20_EMA'] = df['close'].ewm(span=20, adjust=False).mean()
    df['50_SMA'] = df['close'].rolling(window=50).mean()
    df['200_SMA'] = df['close'].rolling(window=200).mean()
    
    # Macro RS Trend (> 50 SMA of the RS Line)
    df['RS_50_SMA'] = df['RS_Line'].rolling(window=50).mean()
    
    # VCP Structure: % Off 52-Week RS High
    df['RS_252_High'] = df['RS_Line'].rolling(window=252).max()
    df['Pct_Off_RS_High'] = ((df['RS_Line'] - df['RS_252_High']) / df['RS_252_High']) * 100
    
    # Skip indices that don't have enough historical data to generate the SMAs
    if len(df) < 70:
        continue
        
    # Extract structural prints for the current and prior week
    latest = df.iloc[-1]
    prev = df.iloc[-2]
    prev_week = df.iloc[-6]
    
    close_val = latest['close']
    pct_chg = ((latest['close'] - prev['close']) / prev['close']) * 100
    
    rs_trend = "Up" if latest['RS_Line'] > latest['RS_50_SMA'] else "Down"
    
    results.append({
        'Sector': sector,
        'Close': round(close_val, 2),
        '% Chg': round(pct_chg, 2),
        '5D RS %': round(latest['5D_RS'], 2),
        '21D RS %': round(latest['21D_RS'], 2),
        '65D RS %': round(latest['65D_RS'], 2),
        'RS Trend (>50 SMA)': rs_trend,
        '% Off RS High': round(latest['Pct_Off_RS_High'], 2),
        '> 10 EMA': "Yes" if close_val > latest['10_EMA'] else "No",
        '> 20 EMA': "Yes" if close_val > latest['20_EMA'] else "No",
        '> 50 SMA': "Yes" if close_val > latest['50_SMA'] else "No",
        '> 200 SMA': "Yes" if close_val > latest['200_SMA'] else "No",
        '65D_Raw': latest['65D_RS'],
        'Prev_65D_Raw': prev_week['65D_RS']
    })

final_df = pd.DataFrame(results)

# 4. Rank Engine & Delta Calculation
final_df['65D RS Rank'] = final_df['65D_Raw'].rank(ascending=False, method='min').astype(int)
final_df['Prev_Rank'] = final_df['Prev_65D_Raw'].rank(ascending=False, method='min').astype(int)
final_df['1-Week Rank Delta'] = final_df['Prev_Rank'] - final_df['65D RS Rank']

# Clean up temporary arrays and sort structurally
final_df = final_df.drop(columns=['65D_Raw', 'Prev_65D_Raw', 'Prev_Rank'])
final_df = final_df.sort_values(by='65D RS Rank')

# Standardize output column architecture
cols = ['Sector', '65D RS Rank', '1-Week Rank Delta', 'Close', '% Chg', '5D RS %', '21D RS %', '65D RS %', 'RS Trend (>50 SMA)', '% Off RS High', '> 10 EMA', '> 20 EMA', '> 50 SMA', '> 200 SMA']
final_df = final_df[cols]

# 5. Dashboard Generation and Excel Styling
final_df.to_excel(file_name, index=False, sheet_name="Sector Momentum")
wb = openpyxl.load_workbook(file_name)
ws = wb["Sector Momentum"]

green_fill = PatternFill(start_color="85E085", end_color="85E085", fill_type="solid")
red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.alignment = Alignment(horizontal='center')
        
        # Color the absolute trend blocks
        if cell.column == cols.index('RS Trend (>50 SMA)') + 1:
            if cell.value == "Up":
                cell.fill = green_fill
            elif cell.value == "Down":
                cell.fill = red_fill
                
        # Color the boolean MA filters
        ma_cols = ['> 10 EMA', '> 20 EMA', '> 50 SMA', '> 200 SMA']
        if cols[cell.column - 1] in ma_cols:
            if cell.value == "Yes":
                cell.fill = green_fill
            elif cell.value == "No":
                cell.fill = red_fill

# Apply heatmap logic to the relative percentage columns
color_scale = ColorScaleRule(start_type='min', start_color='FF9999', mid_type='num', mid_value=0, mid_color='FFFFFF', end_type='max', end_color='85E085')
numeric_cols = ['% Chg', '5D RS %', '21D RS %', '65D RS %', '% Off RS High']
for col_name in numeric_cols:
    col_idx = cols.index(col_name) + 1
    col_letter = openpyxl.utils.get_column_letter(col_idx)
    ws.conditional_formatting.add(f'{col_letter}2:{col_letter}{ws.max_row}', color_scale)

wb.save(file_name)
print("--- Dashboard Exported Successfully ---")
