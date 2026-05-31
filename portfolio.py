import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import sqlite3
import datetime

portfolio = {}
num_stocks = int(input("How many stocks do u want to track? "))
for i in range(num_stocks):
    ticker = input(f"Enter stock ticker {i+1} (eg. AAPL):").upper()
    shares = int(input(f"Enter the number of shares of {ticker} you own:"))
    portfolio[ticker] = shares
tickers = list(portfolio.keys())
data = yf.download(tickers, period = '1y' , auto_adjust = True)['Close']
sp500 = yf.download('^GSPC', period = '1y' , auto_adjust = True ) ['Close']
shares = pd.Series(portfolio)
portfolio_value = data.multiply(shares).sum(axis=1)
daily_returns = portfolio_value.pct_change().dropna()
total_return = total_return = ((portfolio_value.iloc[-1] / portfolio_value.iloc[0]) - 1) * 100
sp500_return = sp500_return = ((sp500.squeeze().iloc[-1] / sp500.squeeze().iloc[0]) - 1) * 100
#Squeeze is used to turn a df into a series , bcz sp500 is a stock and we need it to be a series
#Compares todays price ,yesterdays price and gives percentage 'Your portfolio grew N% this year
volatality = daily_returns.std() * np.sqrt(252) *100
#std gives standard dev , sqrt bcz risk adds up and grows by thr sqrt
sharpe = (daily_returns.mean()/daily_returns.std())*np.sqrt(252)
#Lower Ratio is bad , higher than good and 2-3 Very Good n Exceptional
print(f"\n PORTFOLIO SUMMARY")
print(f"Total Return:  {total_return:.2f}%")
print(f"Volatality:    {volatality:.2f}%")
print(f"Sharpe Ratio:  {sharpe:.2f}")
print(f"Current Value: ${portfolio_value.iloc[-1]:,.2f}")
#:.2f means2 decimal places , :.,2f adds commas for thousands and 2 decimal spots
print(f"S&P 500 Return: {sp500_return:.2f}%")
print(f"Difference:     {total_return - sp500_return:.2f}%")
if total_return > sp500_return:
    print("Your portfolio outperformed the S&P 500!")
else:
    print("Your portfolio underperformed the S&P 500.")
fig, (ax1,ax2) = plt.subplots(1,2 , figsize=(14,5))
ax1.plot(portfolio_value, color = 'steelblue', linewidth = 2)
ax1.plot(sp500/sp500.iloc[0]*portfolio_value.iloc[0],color = 'red' , linewidth = 2,label = 'S&P 500')
ax1.legend()
ax1.set_title( ' Portfolio Value Over Time')
ax1.set_xlabel('Date')
ax1.set_ylabel('Value($)')
ax1.grid(True , alpha = 0.3)
#ax1 and ax2 used because we have 2 subplots , .set used because just plt doesnt work with subplots
current_prices = data.iloc[-1]
values = {t: portfolio[t] * current_prices[t] for t in tickers}
# t is the stock , portfolio[t] is basically how amny shares i own of that stock , same applies to the rest of the things , for t in tickers basically does this process for every stock in my list
ax2.pie(values.values(), labels = values.keys() , autopct = '%1.1f%%' , startangle = 90)
#autopct is a format string that specifies how to display the percentage values on the pie chart. '%1.1f%%' means to display the percentage with one decimal place followed by a percent sign. startangle = 90 means that the first slice of the pie chart will start at the 90-degree angle (the top of the circle) and the slices will be drawn in a counterclockwise direction from there.
ax2.set_title('Current Portfolio Distribution')
plt.tight_layout()
plt.savefig('portfolio_output.png',dpi=150)
#dpi is dots per inch , higher the dpi better the quality of the image but also increases the file size
plt.show()
print("\n CHART SAVED AS portfolio_output.png")
#code to export the data to excel
wb = openpyxl.Workbook()
ws= wb.active
ws.title = 'Portfolio Summary'
ws.append(['Stock','Shares','Current Values($)'])
for t in tickers:
    ws.append([t, portfolio[t], round(portfolio[t]) * float(current_prices[t]),2])
ws.append([])
ws.append(['Total Return (%)', round(total_return,2)])
ws.append(['S&P 500 Return (%)', round(sp500_return,2)])
ws.append(['Difference (%)', round(total_return - sp500_return,2)])
ws.append(['Volatality (%)', round(volatality,2)])
ws.append(['Sharpe Ratio', round(sharpe,2)])
wb.save('portfolio_summary.xlsx')
print("Portfolio summary exported to Portfolio_summary.xlsx")
today = datetime.date.today().strftime('%Y-%m-%d')
#code to export data to sqlite database
conn = sqlite3.connect('portfolio.db')
cursor = conn.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS portfolio_history(date text , total_return real ,sp500_return real , difference real , volatality real , sharpe real , portfolio_value real)''')
cursor.execute('''INSERT INTO portfolio_history (date , total_return , sp500_return , difference, volatality , sharpe , portfolio_value)VALUES(?,?,?,?,?,?,?)''',(today, round(total_return,2),round(sp500_return,2),round(total_return - sp500_return,2),round(volatality,2),round(sharpe,2),round(portfolio_value.iloc[-1],2)))
conn.commit()
conn.close()
print('portfolio summary data saved to portfolio.db')



