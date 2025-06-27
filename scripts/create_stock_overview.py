import sys
import os
import json
import numpy as np
import pandas as pd
import nasdaqdatalink as ndl
import xlwings as xw
import yfinance as yf

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.formatting_helpers import format_metrics

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
API_KEY_PATH = os.path.join(BASE_DIR, 'api_key.json')
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')

with open(API_KEY_PATH) as f:
    config = json.load(f)
    API_KEY = config['api_key']

with open(CONFIG_PATH, 'r') as f:
    config = json.load(f)

METRIC_GROUPS = config['metric_groups']
COLORS = config['colors']

DARK_GREEN = COLORS['DARK_GREEN']
MED_GREEN = COLORS['MED_GREEN']
LIGHT_GREEN = COLORS['LIGHT_GREEN']
DARK_RED = COLORS['DARK_RED']
MED_RED = COLORS['MED_RED']
LIGHT_RED = COLORS['LIGHT_RED']
YELLOW = COLORS['YELLOW']

MARKET_RETURN = 0.08

ndl.ApiConfig.api_key = API_KEY

def calculate_rolling_cagr(values):
    cagr = pd.Series(index=values.index)
    
    for i in range(3, len(values)):
        end_value = values.iloc[i]
        start_value = values.iloc[i-3]
        cagr.iloc[i] = (end_value / start_value) ** (1/3) - 1
                
    return cagr


def fetch_beta_and_rf(ticker):
    stock = yf.Ticker(ticker)
    beta = stock.info.get('beta', 1.0)
    treasury = yf.Ticker('^TNX')  # 10Y Treasury yield
    rf = treasury.history(period='1d')['Close'].iloc[-1] / 100

    if np.isnan(beta) or np.isnan(rf):
        raise ValueError(f"Failed to fetch beta or risk-free rate for {ticker}")
    
    return beta, rf


def compute_wacc(market_cap, debt, interest_exp, tax_exp, ebt, ticker):
    beta, rf = fetch_beta_and_rf(ticker)
    cost_of_equity = rf + beta * (MARKET_RETURN - rf)
    cost_of_debt = 0.0 if debt == 0 else interest_exp / debt

    tax_rate = 0.0 if ebt == 0 else tax_exp / ebt
    total_capital = market_cap + debt
    weight_equity = market_cap / total_capital
    weight_debt = debt / total_capital

    return (weight_equity * cost_of_equity) + (weight_debt * cost_of_debt * (1 - tax_rate))


def grab_fundamental_data(ticker):
    """
    Fetches and calculates comprehensive fundamental analysis metrics for a given ticker.
    Returns historical financial metrics across multiple periods for analysis.
    """
    metrics = {}
    data = ndl.get_table('SHARADAR/SF1', ticker=ticker, paginate=True)

    data = data[data['dimension'] == 'ART']  # As Reported, Trailing Twelve Months (TTM)
    ltm = data.iloc[0:1]
    data = data[data['fiscalperiod'].str.contains('Q4')].copy()

    if data['fiscalperiod'].duplicated().any():
        data = data.sort_values('calendardate', ascending=False).drop_duplicates('fiscalperiod').sort_index()

    data['calendardate'] = pd.to_datetime(data['calendardate'])
    data['year'] = data['calendardate'].dt.year
    ltm['calendardate'] = pd.to_datetime(ltm['calendardate'])
    ltm['year'] = 'LTM'

    data = data.sort_values('year').reset_index(drop=True)
    data = pd.concat([data, ltm])

    sf2_data = ndl.get_table('SHARADAR/SF2', ticker=ticker, paginate=True)
    sf2_data['transactiondate'] = pd.to_datetime(sf2_data['transactiondate'])
    insider_buys = sf2_data[sf2_data['transactioncode'] == 'P']

    def count_insider_buys(end_date):
        if end_date == data['calendardate'].max():
            end_date = pd.to_datetime('today')
        start_date = end_date - pd.DateOffset(months=12)

        count = insider_buys[
            (insider_buys['transactiondate'] > start_date) & 
            (insider_buys['transactiondate'] <= end_date)
        ].shape[0]
        return count
    
    sep_data = ndl.get_table('SHARADAR/SEP', ticker=ticker, paginate=True)
    sep_data['date'] = pd.to_datetime(sep_data['date'])
    latest_share_price = sep_data.sort_values('date').iloc[-1]['close']
    current_shares_outstanding = data['sharesbas'].iloc[-1] * data['sharefactor'].iloc[-1]
    current_market_cap = latest_share_price * current_shares_outstanding

    # Valuation Metrics
    metrics['TEV'] = data['ev'] / 1_000_000
    metrics['Mkt Cap'] = data['marketcap'] / 1_000_000
    metrics['TEV/EBITDA'] = data['ev'] / (data['ebitda'] / data['fxusd'])
    metrics['TEV/Rev'] = data['ev'] / (data['revenue'] / data['fxusd'])
    metrics['TEV/FCF'] = data['ev'] / (data['fcf'] / data['fxusd'])
    metrics['P/E'] = data['marketcap'] / (data['netinc'] / data['fxusd'])
    metrics['P/B'] = data['marketcap'] / ((data['equity']) / data['fxusd'])
    metrics['EPS'] = data['eps'] / data['fxusd']

    ltm_debt = data['debt'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_cash = data['cashneq'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_revenue = data['revenue'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_fcf = data['fcf'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_netinc = data['netinc'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_equity = data['equity'].iloc[-1] / data['fxusd'].iloc[-1]

    ltm_interest_exp = data['intexp'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_tax_exp = data['taxexp'].iloc[-1] / data['fxusd'].iloc[-1]
    ltm_ebt = data['ebt'].iloc[-1] / data['fxusd'].iloc[-1]

    if np.isnan(data['ebitda'].iloc[-1]): # LTM EBITDA is nan for Chinese stocks
        ltm_ebitda = data['ebitda'].iloc[-2] / data['fxusd'].iloc[-1]
    else:
        ltm_ebitda = data['ebitda'].iloc[-1] / data['fxusd'].iloc[-1]

    new_ev = current_market_cap + ltm_debt - ltm_cash

    # Keep up to date valuation metrics as market cap changes past latest earnings
    metrics['TEV'].iat[-1] = new_ev / 1_000_000
    metrics['Mkt Cap'].iat[-1] = current_market_cap / 1_000_000
    metrics['TEV/EBITDA'].iat[-1] = new_ev / ltm_ebitda
    metrics['TEV/Rev'].iat[-1] = new_ev / ltm_revenue
    metrics['TEV/FCF'].iat[-1] = new_ev / ltm_fcf
    metrics['P/E'].iat[-1] = current_market_cap / ltm_netinc
    metrics['P/B'].iat[-1] = current_market_cap / ltm_equity

    # Income Statement
    metrics['Rev'] = (data['revenue'] / data['fxusd']) / 1_000_000
    metrics['Rev 3YCAGR'] = calculate_rolling_cagr(data['revenue'] / data['fxusd'])
    metrics['GP'] = (data['gp'] / data['fxusd']) / 1_000_000
    metrics['Net Inc'] = (data['netinc'] / data['fxusd']) / 1_000_000
    metrics['Op Inc'] = (data['opinc'] / data['fxusd']) / 1_000_000
    metrics['EBITDA'] = (data['ebitda'] / data['fxusd']) / 1_000_000

    # Operating Expense Detail
    metrics['R&D'] = (data['rnd'] / data['fxusd']) / 1_000_000
    metrics['SG&A'] = (data['sgna'] / data['fxusd']) / 1_000_000
    metrics['D&A'] = (data['depamor'] / data['fxusd']) / 1_000_000
    metrics['SBC'] = (data['sbcomp'] / data['fxusd']) / 1_000_000

    # Operating Expense Ratios (with validation)
    safe_revenue = data['revenue'].replace(0, np.nan)
    metrics['R&D/Rev'] = data['rnd'] / safe_revenue
    metrics['SG&A/Rev'] = data['sgna'] / safe_revenue
    metrics['SBC/Rev'] = data['sbcomp'] / safe_revenue

    # Cash Flow
    metrics['CFO'] = (data['ncfo'] / data['fxusd']) / 1_000_000
    metrics['FCF'] = (data['fcf'] / data['fxusd']) / 1_000_000
    metrics['Op Exp'] = (data['opex'] / data['fxusd']) / 1_000_000
    metrics['CapEx'] = (data['capex'] / data['fxusd']) / 1_000_000
    metrics['Int Exp'] = (data['intexp'] / data['fxusd']) / 1_000_000

    # Cash Flow Analysis (with validation)
    safe_netinc = data['netinc'].replace(0, np.nan)
    metrics['NI to CFO'] = data['ncfo'] / safe_netinc
    metrics['SBC Add-back'] = (data['sbcomp'] / data['fxusd']) / 1_000_000
    
    # Calculate working capital change
    current_wc = (data['assetsc'] - data['liabilitiesc']) / data['fxusd']
    prev_wc = current_wc.shift(1)
    metrics['WC Change'] = (prev_wc - current_wc) / 1_000_000  # Negative means cash outflow

    # Margins
    metrics['GP Marg'] = data['grossmargin']
    metrics['EBITDA Marg'] = data['ebitdamargin']
    metrics['Net Marg'] = data['netmargin']
    metrics['Op Marg'] = data['opinc'] / safe_revenue
    metrics['FCF Marg'] = data['fcf'] / safe_revenue

    # Shareholder Yield
    metrics['Div Yield'] = data['divyield']
    metrics['BB Yield'] = (data['sharesbas'].shift(1) - data['sharesbas']) / data['sharesbas'].shift(1)
    metrics['Ins Buys'] = data['calendardate'].apply(count_insider_buys)

    # Balance Sheet
    metrics['Equity'] = (data['equity'] / data['fxusd']) / 1_000_000
    metrics['Debt'] = (data['debt'] / data['fxusd']) / 1_000_000
    metrics['Assets'] = (data['assets'] / data['fxusd']) / 1_000_000
    metrics['Liab'] = (data['liabilities'] / data['fxusd']) / 1_000_000
    metrics['Cash & ST Inv'] = (data['cashneq'] + data['investmentsc']) / 1_000_000
    metrics['Net Cash'] = (data['cashneq'] + data['investmentsc'] - data['debt']) / 1_000_000
    metrics['TBV'] = ((data['assets'] - data['intangibles'] - data['liabilities']) / data['fxusd']) / 1_000_000

    # Asset Quality
    metrics['Receivables'] = (data['receivables'] / data['fxusd']) / 1_000_000
    metrics['Inventory'] = (data['inventory'] / data['fxusd']) / 1_000_000
    metrics['PPE Net'] = (data['ppnenet'] / data['fxusd']) / 1_000_000
    metrics['Intangibles'] = (data['intangibles'] / data['fxusd']) / 1_000_000
    metrics['Payables'] = (data['payables'] / data['fxusd']) / 1_000_000
    metrics['Def Revenue'] = (data['deferredrev'] / data['fxusd']) / 1_000_000

    # Working Capital Analysis (with data validation)
    safe_cor = data['cor'].replace(0, np.nan)
    safe_receivables = data['receivables'].replace(0, np.nan)
    safe_inventory = data['inventory'].replace(0, np.nan)

    metrics['DSO'] = (data['receivables'] / safe_revenue) * 365  # Days Sales Outstanding
    metrics['DIO'] = (data['inventory'] / safe_cor) * 365  # Days Inventory Outstanding
    metrics['DPO'] = (data['payables'] / safe_cor) * 365  # Days Payable Outstanding
    metrics['Cash Cycle'] = metrics['DSO'] + metrics['DIO'] - metrics['DPO']

    # Solvency
    metrics['D/E'] = data['debt'] / data['equity']
    metrics['Debt/EBITDA'] = data['debt'] / data['ebitda']
    metrics['Cash Ratio'] = data['cashneq'] / data['liabilitiesc']
    metrics['Cash/Debt'] = data['cashneq'] / data['debt']
    metrics['Int Cov'] = data['ebit'] / data['intexp']

    # Liquidity
    metrics['Curr Ratio'] = data['currentratio']
    metrics['Quick Ratio'] = (data['assetsc'] - data['inventory']) / data['liabilitiesc']

    # Efficiency (with validation for new turnover ratios)
    safe_working_capital = (data['assetsc'] - data['liabilitiesc']).replace(0, np.nan)
    
    metrics['WC Turn'] = safe_revenue / safe_working_capital
    metrics['Asset Turn'] = data['assetturnover']
    metrics['Recv Turn'] = safe_revenue / safe_receivables  # Revenue / Receivables
    metrics['Inv Turn'] = safe_cor / safe_inventory  # COGS / Inventory

    # Profitability
    metrics['ROA'] = data['roa']
    metrics['ROE'] = data['roe']
    metrics['ROIC'] = data['roic']

    # Create DataFrame and apply data cleaning
    metrics_df = pd.DataFrame(metrics)
    
    # Replace infinite values and extreme outliers
    metrics_df = metrics_df.replace([np.inf, -np.inf], np.nan)

    # Cap extreme values for display purposes (optional but recommended)
    for col in ['DSO', 'DIO', 'DPO', 'Cash Cycle']:
        if col in metrics_df.columns:
            # Cap at 999 days for display (negative values allowed for Cash Cycle)
            if col == 'Cash Cycle':
                metrics_df[col] = metrics_df[col].clip(lower=-999, upper=999)
            else:
                metrics_df[col] = metrics_df[col].clip(upper=999)

    # Cap turnover ratios to reasonable ranges
    for col in ['Recv Turn', 'Inv Turn', 'WC Turn']:
        if col in metrics_df.columns:
            metrics_df[col] = metrics_df[col].clip(upper=100)  # Cap at 100x turnover

    # Cap NI to CFO ratio to reasonable range
    if 'NI to CFO' in metrics_df.columns:
        metrics_df['NI to CFO'] = metrics_df['NI to CFO'].clip(lower=-10, upper=10)

    metrics_df.index = data['year']

    wacc = compute_wacc(
        market_cap=current_market_cap,
        debt=ltm_debt,
        interest_exp=ltm_interest_exp,
        tax_exp=ltm_tax_exp,
        ebt=ltm_ebt,
        ticker=ticker
    )

    return metrics_df.round(2), wacc


def apply_conditional_formatting(sheet, metrics_df, start_row, start_col):
    # Work with transposed data to match Excel layout
    transposed_metrics = metrics_df.transpose()
    
    for row_idx, metric_name in enumerate(transposed_metrics.index):    
        row_values = transposed_metrics.loc[metric_name].values
        
        current_row = start_row + 1 + row_idx  # +1 to skip header row
        data_range = sheet.range((current_row, start_col + 2),  # +2 to skip category and metric name columns
                               (current_row, start_col + 2 + len(row_values) - 1))
        
        format_metrics(data_range, row_values, metric_name)


def write_dcf_to_excel(sheet, start_col, wacc, fcf_row_num, years):
    dcf_start_row = 10
    dcf_start_col = start_col + len(years) + 3

    sheet.cells(dcf_start_row, dcf_start_col + 1).value = "DF"
    sheet.cells(dcf_start_row, dcf_start_col + 2).value = "10Y GR"
    sheet.cells(dcf_start_row, dcf_start_col + 3).value = "Perp GR"

    wacc_cell = sheet.cells(dcf_start_row + 1, dcf_start_col + 1)
    wacc_cell.value = wacc
    wacc_cell.api.NumberFormat = "0.0000"
    sheet.cells(dcf_start_row + 1, dcf_start_col + 2).value = 1.1
    sheet.cells(dcf_start_row + 1, dcf_start_col + 3).value = 1.03

    sheet.range((dcf_start_row, dcf_start_col + 1), (dcf_start_row + 1, dcf_start_col + 1)).color = (255, 116, 116)
    sheet.range((dcf_start_row, dcf_start_col + 2), (dcf_start_row + 1, dcf_start_col + 2)).color = (146, 208, 80)
    sheet.range((dcf_start_row, dcf_start_col + 3), (dcf_start_row + 1, dcf_start_col + 3)).color = (255, 255, 0)

    if fcf_row_num:
        discount_factor_cell = sheet.cells(dcf_start_row + 1, dcf_start_col + 1).address
        gr_10y_cell = sheet.cells(dcf_start_row + 1, dcf_start_col + 2).address
        perp_gr_cell = sheet.cells(dcf_start_row + 1, dcf_start_col + 3).address

        # 10Y FCF Extrapolation
        for i in range(10):
            current_col = dcf_start_col + i
            prev_col_addr = sheet.cells(fcf_row_num, current_col - 1).address
            formula = f"={prev_col_addr}*({gr_10y_cell})"
            sheet.cells(fcf_row_num, current_col).formula = formula
            sheet.cells(fcf_row_num, current_col).api.NumberFormat = "#,##0"

        # 40Y Perpetual Growth FCF Extrapolation
        for i in range(40):
            current_col = dcf_start_col + 10 + i
            prev_col_addr = sheet.cells(fcf_row_num, current_col - 1).address
            formula = f"={prev_col_addr}*({perp_gr_cell})"
            sheet.cells(fcf_row_num, current_col).formula = formula
            sheet.cells(fcf_row_num, current_col).api.NumberFormat = "#,##0"

        # DCF Calculation
        dcf_label_cell = sheet.cells(fcf_row_num + 2, dcf_start_col + 1)
        dcf_value_cell = sheet.cells(fcf_row_num + 3, dcf_start_col + 1)

        dcf_label_cell.value = "NPV"

        # Construct NPV formula
        npv_start_cell = sheet.cells(fcf_row_num, dcf_start_col).address
        npv_end_cell = sheet.cells(fcf_row_num, dcf_start_col + 49).address

        npv_formula = f"=NPV({discount_factor_cell}, {npv_start_cell}:{npv_end_cell})"
        dcf_value_cell.formula = npv_formula
        dcf_value_cell.api.NumberFormat = "#,##0"
        sheet.range((fcf_row_num + 2, dcf_start_col + 1), (fcf_row_num + 3, dcf_start_col + 1)).color = (77, 147, 217)
        sheet.api.Columns(dcf_start_col + 1).AutoFit()


def write_to_excel(sheet, metrics, wacc, start_row=4, start_col=5):
    years = sorted([idx for idx in metrics.index if idx != 'LTM'])
    max_years_for_data = 15
    if len(years) > max_years_for_data:
        years = years[-max_years_for_data:]
        metrics = metrics.loc[years + ['LTM']]

    transposed_metrics = metrics.transpose()

    headers = ["Category", "Metric"] + [str(year) for year in years] + ["LTM"]
    for col_num, header in enumerate(headers, start=start_col):
        sheet.cells(start_row, col_num).value = header

    current_row = start_row + 1

    for i, group in enumerate(METRIC_GROUPS):
        group_start_row = current_row
        first_metric = True
        
        for metric_name, description in group['metrics'].items():
            if metric_name in transposed_metrics.index:
                # Write category name (only for first metric in group)
                if first_metric:
                    category_cell = sheet.cells(current_row, start_col)
                    category_cell.value = group['name']
                    first_metric = False
                
                metric_cell = sheet.cells(current_row, start_col + 1)
                metric_cell.value = metric_name
                
                if metric_cell.api.Comment is not None:
                    metric_cell.api.Comment.Delete()
                metric_cell.api.AddComment(description)
                metric_cell.api.Comment.Visible = False
                
                for col_num, value in enumerate(transposed_metrics.loc[metric_name], start=start_col + 2):
                    cell = sheet.cells(current_row, col_num)
                    cell.value = value
                    
                    percentage_metrics = [
                        'GP Marg', 'EBITDA Marg', 'Net Marg', 'Op Marg', 'FCF Marg',
                        'Div Yield', 'BB Yield', 'Rev 3YCAGR',
                        'R&D/Rev', 'SG&A/Rev', 'SBC/Rev',
                        'ROA', 'ROE', 'ROIC'
                    ]
                    
                    decimal_metrics = [
                        'Curr Ratio', 'Quick Ratio', 'D/E', 'Debt/EBITDA', 'Cash Ratio', 'Cash/Debt',
                        'Int Cov', 'WC Turn', 'Asset Turn', 'Recv Turn', 'Inv Turn', 'EPS', 'NI to CFO',
                        'TEV/EBITDA', 'TEV/Rev', 'TEV/FCF', 'P/E', 'P/B'
                    ]
                    
                    whole_number_metrics = ['DSO', 'DIO', 'DPO', 'Cash Cycle', 'Ins Buys']
                    
                    # Apply formatting based on explicit categories
                    if metric_name in percentage_metrics:
                        cell.api.NumberFormat = "0%"
                    elif metric_name in decimal_metrics:
                        cell.api.NumberFormat = "0.00"
                    elif metric_name in whole_number_metrics:
                        cell.api.NumberFormat = "0"
                    elif isinstance(value, (int, float)) and pd.notna(value):
                        cell.api.NumberFormat = "#,##0"
                    
                current_row += 1

        if i < len(METRIC_GROUPS) - 1:
            border_range = sheet.range(
                sheet.cells(group_start_row, start_col),
                sheet.cells(current_row - 1, start_col + len(years) + 2)
            )
            border_range.api.Borders(9).Weight = 2

    table_range = sheet.range(
        sheet.cells(start_row, start_col),
        sheet.cells(current_row - 1, start_col + len(years) + 2)
    )

    for i, row in enumerate(table_range.rows):
        row_number = start_row + i

        if row_number == start_row:
            row.color = (180, 180, 180)  # Dark grey
        elif row_number == start_row + 1:
            pass
        elif (row_number - start_row) % 2 == 0:
            row.color = (217, 217, 217)  # Light grey
    
    for col in range(start_col, start_col + len(years) + 2):
        sheet.api.Columns(col).AutoFit()

    apply_conditional_formatting(sheet, metrics, start_row, start_col)

    sheet.range((1, 1), (1, sheet.api.Columns.Count)).color = (185, 216, 72)

    sheet.range((2, 1), (3, sheet.api.Columns.Count)).color = (0, 201, 192)

    fcf_row_num = None
    for i in range(start_row + 1, current_row):
        if sheet.cells(i, start_col + 1).value == 'FCF':
            fcf_row_num = i
            break
    write_dcf_to_excel(sheet, start_col, wacc, fcf_row_num, years)


def api_test():
    data, wacc = grab_fundamental_data('BABA')
    print(data)
    print(f"WACC: {wacc:.2%}")

def main():
    _ = sys.argv[1] # spreadsheet path
    ticker = sys.argv[2]
    metrics, wacc = grab_fundamental_data(ticker)

    wb = xw.books.active
    sheet = wb.sheets.active
    write_to_excel(sheet, metrics, wacc, start_row=4, start_col=5)
    header_cell = sheet.cells(1, 5)
    header_cell.value = f"{ticker} Overview"
    header_cell.api.Font.Size = 20


main()