import yfinance as yf
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule

print("--- NSE Sector Rotation Engine Initiated ---")

file_name = "NSE_Sector_Monitor.xlsx"
benchmark_ticker = "^CRSLDX" # Nifty 500

tickers = {
    "IT": "^CNXIT",
    "Bank": "^NSEBANK",
    "Auto": "^CNXAUTO",
    "FMCG": "^CNXFMCG",
    "Pharma": "^CNXPHARMA",
    "Metal": "^CNXMETAL",
    "Energy": "^CNXENERGY",
    "Realty": "^CNXREALTY",
    "Fin Services": "^CNXFIN",
    "Infrastructure": "^CNXINFRA",
    "Consumption": "^CNXCONSUM",
    "Commodities": "^CNXCMDT",
    "PSE": "^CNXPSE",
    "MNC": "^CNXMNC",
    "Media": "^CNXMEDIA",
    "PSU Bank": "^CNXPSUBANK",
    "Consumer Durables": "VOLTAS.NS", # Proxy heavyweight for Durables
    "Healthcare": "HEALTHY.NS",       # Aditya BSL Nifty Healthcare ETF
    "Defence": "MODEFENCE.NS"
}

# 1. Download Data
print("Downloading 2-Year Market Data...")
all_tickers = list(tickers.values()) + [benchmark_ticker]
raw_data = yf.download(all_tickers, period="2y", group_by='ticker')

# Check for multi-index formatting from yfinance
if isinstance(raw_data.columns, pd.MultiIndex):
    close_data = pd.DataFrame()
    for t in all_tickers:
        if t in raw_data:
            close_data[t] = raw_data[t]['Close']
else:
    close_data = raw_data['Close']

close_data = close_data.ffill().dropna(how='all')
benchmark_close = close_data[benchmark_ticker]

# 2. Calculate Sector Matrices
print("Calculating Relative Strength and Moving Averages...")
metrics = []
historical_ranks = pd.DataFrame(index=close_data.index, columns=tickers.keys())

# Calculate daily ranks for Tab 2
for date_idx in range(65, len(close_data)):
    daily_rs_returns = {}
    for sector_name, symbol in tickers.items():
        if symbol not in close_data.columns: continue
        
        sector_series = close_data[symbol].iloc[:date_idx+1]
        bench_series = benchmark_close.iloc[:date_idx+1]
        
        # RS Line = Sector Price / Benchmark Price
        rs_line = sector_series / bench_series
        
        # 65-Day RS Momentum
        if len(rs_line) > 65:
            rs_65d = ((rs_line.iloc[-1] - rs_line.iloc[-66]) / rs_line.iloc[-66]) * 100
            daily_rs_returns[sector_name] = rs_65d
            
    # Rank them for this specific day (Highest RS % = Rank 1)
    if daily_rs_returns:
        sorted_ranks = sorted(daily_rs_returns.items(), key=lambda x: x[1], reverse=True)
        for rank, (sec, val) in enumerate(sorted_ranks, 1):
            historical_ranks.at[close_data.index[date_idx], sec] = rank

# Today's Calculations for Tab 1
for sector_name, symbol in tickers.items():
    if symbol not in close_data.columns: 
        print(f"Warning: {sector_name} ({symbol}) data unavailable.")
        continue
        
    sec_close = close_data[symbol]
    rs_line = sec_close / benchmark_close
    
    # Base Price Metrics
    chg_1d = ((sec_close.iloc[-1] - sec_close.iloc[-2]) / sec_close.iloc[-2]) * 100
    
    # Moving Averages
    ema_10 = sec_close.ewm(span=10, adjust=False).mean().iloc[-1]
    ema_20 = sec_close.ewm(span=20, adjust=False).mean().iloc[-1]
    sma_50 = sec_close.rolling(50).mean().iloc[-1]
    sma_200 = sec_close.rolling(200).mean().iloc[-1]
    
    gt_10 = "Yes" if sec_close.iloc[-1] > ema_10 else "No"
    gt_20 = "Yes" if sec_close.iloc[-1] > ema_20 else "No"
    gt_50 = "Yes" if sec_close.iloc[-1] > sma_50 else "No"
    gt_200 = "Yes" if sec_close.iloc[-1] > sma_200 else "No"
    
    # RS Metrics
    rs_5d = ((rs_line.iloc[-1] - rs_line.iloc[-6]) / rs_line.iloc[-6]) * 100 if len(rs_line) > 5 else 0
    rs_21d = ((rs_line.iloc[-1] - rs_line.iloc[-22]) / rs_line.iloc[-22]) * 100 if len(rs_line) > 21 else 0
    rs_65d = ((rs_line.iloc[-1] - rs_line.iloc[-66]) / rs_line.iloc[-66]) * 100 if len(rs_line) > 65 else 0
    
    rs_sma_50 = rs_line.rolling(50).mean().iloc[-1]
    rs_trend = "Up" if rs_line.iloc[-1] > rs_sma_50 else "Down"
    
    rs_252_max = rs_line.rolling(252).max().iloc[-1]
    rs_52w_high = "Yes" if rs_line.iloc[-1] >= rs_252_max else ""
    
    # Rank Delta
    current_rank = historical_ranks[sector_name].iloc[-1]
    past_rank_1w = historical_ranks[sector_name].iloc[-6] # 5 trading days ago
    
    rank_delta = 0
    if pd.notna(current_rank) and pd.notna(past_rank_1w):
        # If past rank was 10, current is 4. Delta is +6
        rank_delta = int(past_rank_1w) - int(current_rank)

    metrics.append({
        "Sector": sector_name,
        "65D RS Rank": current_rank,
        "1-Week Rank Delta": f"+{rank_delta}" if rank_delta > 0 else str(rank_delta),
        "Close": round(sec_close.iloc[-1], 2),
        "% Chg": round(chg_1d, 2),
        "5D RS %": round(rs_5d, 2),
        "21D RS %": round(rs_21d, 2),
        "65D RS %": round(rs_65d, 2),
        "RS Trend (>50 SMA)": rs_trend,
        "52W RS High": rs_52w_high,
        "> 10 EMA": gt_10,
        "> 20 EMA": gt_20,
        "> 50 SMA": gt_50,
        "> 200 SMA": gt_200
    })

df_heatmap = pd.DataFrame(metrics).sort_values(by="65D RS Rank", ascending=True)

# 3. Build Excel File
print("Writing Engine to Excel...")
wb = Workbook()

# Tab 1: Cross-Sectional Heatmap
ws1 = wb.active
ws1.title = "Heatmap"
headers_t1 = df_heatmap.columns.tolist()
ws1.append(headers_t1)

for r in df_heatmap.itertuples(index=False):
    ws1.append(list(r))

# Tab 2: Historical Ranks Tracker
ws2 = wb.create_sheet(title="Rotation Tracker")
# Get last 65 trading days
hist_tracker = historical_ranks.dropna(how='all').tail(65).sort_index(ascending=False)
headers_t2 = ["Date"] + list(hist_tracker.columns)
ws2.append(headers_t2)

for date, row in hist_tracker.iterrows():
    r_data = [date.strftime("%Y-%m-%d")] + [row[col] for col in hist_tracker.columns]
    ws2.append(r_data)

# 4. Apply UI Formatting
green_fill = PatternFill(start_color="63BE7B", end_color="63BE7B", fill_type="solid")
red_fill = PatternFill(start_color="F8696B", end_color="F8696B", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")

# Tab 1 Formatting
for col in ws1.columns:
    col_let = col[0].column_letter
    ws1.column_dimensions[col_let].width = 15
    for cell in col: cell.alignment = center_align

# Conditional formatting for Time-Frame Matrices (Columns F, G, H - 5D, 21D, 65D RS)
rs_scale = ColorScaleRule(start_type='num', start_value=-10, start_color='F8696B', mid_type='num', mid_value=0, mid_color='FFFFFF', end_type='num', end_value=10, end_color='63BE7B')
ws1.conditional_formatting.add(f"F2:H{ws1.max_row}", rs_scale)

# Conditional Formatting for Binary Indicators
ws1.conditional_formatting.add(f"I2:I{ws1.max_row}", CellIsRule(operator='equal', formula=['"Up"'], fill=green_fill))
ws1.conditional_formatting.add(f"I2:I{ws1.max_row}", CellIsRule(operator='equal', formula=['"Down"'], fill=red_fill))

ws1.conditional_formatting.add(f"J2:J{ws1.max_row}", CellIsRule(operator='equal', formula=['"Yes"'], fill=green_fill))

for col_let in ['K', 'L', 'M', 'N']: # Moving Averages
    ws1.conditional_formatting.add(f"{col_let}2:{col_let}{ws1.max_row}", CellIsRule(operator='equal', formula=['"Yes"'], fill=green_fill))
    ws1.conditional_formatting.add(f"{col_let}2:{col_let}{ws1.max_row}", CellIsRule(operator='equal', formula=['"No"'], fill=red_fill))

# Tab 2 Formatting
for col in ws2.columns:
    col_let = col[0].column_letter
    ws2.column_dimensions[col_let].width = 12
    for cell in col: cell.alignment = center_align

rank_scale = ColorScaleRule(start_type='num', start_value=1, start_color='63BE7B', mid_type='num', mid_value=10, mid_color='FFFFFF', end_type='num', end_value=19, end_color='F8696B')

# BUG FIX: Get last column letter safely to avoid generator subscript error
last_col_letter = ws2.cell(row=1, column=ws2.max_column).column_letter
ws2.conditional_formatting.add(f"B2:{last_col_letter}{ws2.max_row}", rank_scale)

wb.save(file_name)
print(f"--- SUCCESS! Engine saved to {file_name} ---")
