# VBA-challenge
In this assignment given by Birmingham University Data Boot Camp a generated stock market data has been analyzed using VB scripting. 

-Alphabetical_testing.xlxs has been used for developing codes and making tests since, its data was smaller (Total data=153.890 , sheets A,B,C,D,E,F). Then codes have been run in Multiple_year_stock_data.xlxs (Total data=2.268.003, sheets 2018-2019-2020).

-VB codes loop through each sheet. From the data, below information has been identified in separate columns and tables.

Table 1:

Ticker Symbol: Column I. Each different ticker symbol is written here.

Yearly Change: Column J. The difference between open and close values at the beginning day and last day of the year. 

Percent Change: Column K,  yearly change is divided by opening value. Percent change column has been % formatted.

Yearly Change and Percent Change columns have been formatted with green and red for positive and negative changes.

Total Stock Volume: Column L, Total Stock Volumes for each ticker symbol.

Table: 2

Ticker and Value of:

Greatest % increase (O2) : Biggest yearly change as percentage

Greatest % decrease (O3) : Smallest yearly change as percentage

Greatest Total Volume (O4) :  Biggest of Total Stock Volumes

