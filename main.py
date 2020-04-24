import pandas as pd
import math

ORIGINAL_CAPITAL = 1000000
cash = ORIGINAL_CAPITAL

holdings_cols = ["Ticker", "Stock Price", "Stock Number", "Year Month"]
holdings = pd.DataFrame(columns=holdings_cols)

capital_at_time = pd.DataFrame(columns=['Capital', 'Month'])

def year_month_int(y, m):
    return int(f'{y}{m}' if m > 9 else f'{y}0{m}')

def owned_tickers():
    return list(holdings['Ticker'])

def last_stock_price(ticker, y, q):
    quarter = q - 1
    holding_info = pd.DataFrame()
    for year in reversed(range(2007, y+1)):
        while quarter > 0:
            try:
                old_q_df = df[str(year)].loc[df[str(year)]["Fiscal Year"] == year].loc[df[str(year)]["Fiscal Quarter"] == quarter]
                holding_info = old_q_df.loc[old_q_df["Ticker Symbol"] == ticker]
                if holding_info.shape[0] > 0 and not holding_info['Stock Price'].isnull().any():
                    return holding_info
            except pd.core.indexing.IndexingError:
                print("not here")
            quarter -= 1
        if holding_info.shape[0] > 0 and not holding_info['Stock Price'].isnull().any():
            return holding_info
        quarter = 4
    raise KeyError

def cur_capital(current_quarter_df):
    capital = cash
    for i, holding in holdings.iterrows():
        holding_info = current_quarter_df.loc[current_quarter_df["Ticker Symbol"] == holding["Ticker"]]

        # GET LAST STOCK PRICE IF STOCK NO LONGER IN INFO
        if holding_info.shape[0] == 0:
            holding_info = last_stock_price(holding["Ticker"], 2009, 4)

        capital += float(holding_info['Stock Price'])*float(holding['Stock Number'])
    return capital

# read in excel file
df = pd.read_excel('data.xlsx', ['2007', '2008', '2009'])
# print(df["2007"].keys())

for y in range(2007, 2010):
    df[str(y)]["Earnings Yield"] = df[str(y)]["Operating Income After Depreciation - Quarterly"]/\
                                   df[str(y)]["Common/Ordinary Equity - Total"]
    df[str(y)]["ROC"] = df[str(y)]["Operating Income After Depreciation - Quarterly"]/\
                        (df[str(y)]["Property Plant and Equipment - Total (Net)"]+df[str(y)]["Working Capital (Balance Sheet)"])

months_since_start = 0
for y in range(2007, 2010):
    for m in range(12):
        q = m//3+1
        print(f"{y}, {m} (Q{q}):")

        year_df=df[str(y)]
        quarter_df = year_df.loc[year_df["Fiscal Year"] == y].loc[year_df["Fiscal Quarter"] == q].loc[year_df["Stock Price"] > 0].dropna(subset=["Earnings Yield", "ROC"])
        quarter_df = quarter_df[~quarter_df['Ticker Symbol'].isin(owned_tickers())]
        quarter_df["Earnings Yield Rank"] = quarter_df["Earnings Yield"].rank()
        quarter_df["ROC Rank"] = quarter_df["ROC"].rank()
        quarter_df["Rank"] = quarter_df["Earnings Yield Rank"] + quarter_df["ROC Rank"]
        quarter_df = quarter_df.sort_values(by=["Rank"], ascending=False)

        # SELLING
        to_sell = holdings.loc[holdings["Year Month"] <= year_month_int(y-1,m)]
        if (to_sell.shape[0] > 1):
            print('Selling:')
            print(to_sell)

        for i, holding in to_sell.iterrows():
            holding_info = quarter_df.loc[quarter_df["Ticker Symbol"] == holding["Ticker"]]

            # GET LAST STOCK PRICE IF STOCK NO LONGER IN INFO
            if holding_info.shape[0] == 0:
                holding_info = last_stock_price(holding["Ticker"], y, q)

            cash += float(holding_info['Stock Price'])*float(holding['Stock Number'])
            holdings = holdings.loc[holdings['Ticker'] != holding['Ticker']]

        # BUYING
        if months_since_start < 12:
            money_to_spend = ORIGINAL_CAPITAL/12
        else:
            money_to_spend = cash

        first = quarter_df.iloc[0]
        first_amount = round((money_to_spend/2)/first["Stock Price"])
        money_to_spend -= first_amount*first["Stock Price"]

        second = quarter_df.iloc[1]
        second_amount = math.floor(money_to_spend/second["Stock Price"])

        cash -= first_amount*first["Stock Price"] + second_amount*second["Stock Price"]
        new_df = pd.DataFrame([
            [first["Ticker Symbol"], first["Stock Price"], first_amount, year_month_int(y,m)],
            [second["Ticker Symbol"], second["Stock Price"], second_amount, year_month_int(y,m)]
        ], columns=holdings_cols)
        print('Buying:')
        print(new_df)

        holdings=holdings.append(new_df, ignore_index=True)
        months_since_start+=1

        print(f"Portfolio:")
        print(f"Cash: ${round(cash)}")
        print(f"Number of holdings: {holdings.shape[0]}")
        current_capital = round(cur_capital(quarter_df))
        print(f"Capital: ${current_capital}\n")
        current_capital_df = pd.DataFrame([[current_capital, year_month_int(y,m)]], columns=['Capital', 'Month'])
        capital_at_time=capital_at_time.append(current_capital_df, ignore_index=True)

print('\n')
total_end_capital = cash
for i, holding in holdings.iterrows():
    holding_info = quarter_df.loc[quarter_df["Ticker Symbol"] == holding["Ticker"]]

    # GET LAST STOCK PRICE IF STOCK NO LONGER IN INFO
    if holding_info.shape[0] == 0:
        holding_info = last_stock_price(holding["Ticker"], 2009, 4)

    total_end_capital += float(holding_info['Stock Price'])*float(holding['Stock Number'])
print(f'Total end capital: ${round(total_end_capital)}')
print(f'Returns: {round(total_end_capital/ORIGINAL_CAPITAL*100)-100}%')
