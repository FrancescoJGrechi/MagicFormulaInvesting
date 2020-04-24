import pandas as pd

ORIGINAL_CAPITAL = 1000000
cash = ORIGINAL_CAPITAL

holdings_cols = ["Ticker", "Stock Price", "Stock Number", "Year Month"]
holdings = pd.DataFrame(columns=holdings_cols)
print(holdings)

# read in excel file
df = pd.read_excel('data.xlsx', ['2007', '2008', '2009'])
df2007 = df["2007"]
print(df2007.keys())

for y in range(2007, 2010):
    df[str(y)]["Earnings Yield"] = df[str(y)]["Operating Income After Depreciation - Quarterly"]/\
                                   df[str(y)]["Common/Ordinary Equity - Total"]
    df[str(y)]["ROC"] = df[str(y)]["Operating Income After Depreciation - Quarterly"]/\
                        (df[str(y)]["Property Plant and Equipment - Total (Net)"]+df[str(y)]["Working Capital (Balance Sheet)"])

months_since_start = 0
for y in range(2007, 2010):
    for m in range(12):
        q = m//3+1
        year_df=df[str(y)]
        quarter_df = year_df.loc[year_df["Fiscal Year"] == y].loc[year_df["Fiscal Quarter"] == q].dropna(subset=["Earnings Yield", "ROC"])
        quarter_df["Earnings Yield Rank"] = quarter_df["Earnings Yield"].rank()
        quarter_df["ROC Rank"] = quarter_df["ROC"].rank()
        quarter_df["Rank"] = quarter_df["Earnings Yield Rank"] + quarter_df["ROC Rank"]
        quarter_df = quarter_df.sort_values(by=["Rank"], ascending=False)

        # SELLING
        to_sell = holdings.loc[holdings["Year Month"] <= int(f"{y-1}{m}")]
        for i, holding in to_sell.iterrows():
            holding_info = quarter_df.loc[quarter_df["Ticker Symbol"] == holding["Ticker"]]

            if y == 2007 and m == 10:
                print(holding_info['Stock Price'])
                print(holding['Stock Number'])

            cash += float(holding_info['Stock Price'])*float(holding['Stock Number'])
        holdings = pd.concat([holdings, to_sell]).drop_duplicates(keep=False)

        # BUYING
        if months_since_start < 12:
            money_to_spend = ORIGINAL_CAPITAL/24
        else:
            money_to_spend = cash/2

        first = quarter_df.iloc[2*m]
        first_amount = money_to_spend//first["Stock Price"]
        second = quarter_df.iloc[2*m+1]
        second_amount = money_to_spend//second["Stock Price"]

        cash -= first_amount*first["Stock Price"] + second_amount*second["Stock Price"]
        new_df = pd.DataFrame([
            [first["Ticker Symbol"], first["Stock Price"], first_amount, int(f"{y}{m}")],
            [second["Ticker Symbol"], second["Stock Price"], second_amount, int(f"{y}{m}")]
        ], columns=holdings_cols)
        holdings=holdings.append(new_df, ignore_index=True)
        months_since_start+=1

        print(f"{y}, {m} - Portfolio:")
        print(f"Cash: ${cash}")
        print(f"Number of holdings: {holdings.shape}")
        #print(holdings)
        print()