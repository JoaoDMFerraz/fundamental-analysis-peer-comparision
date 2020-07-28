# Fundamental-Analysis: peer-comparision
This project was created to analyse the financial statements of public companies for a mexican pension fund.

This is a python file that will extract data from yahoo finance about public companies and insert it in a excel file created and named by the user in the same path of the python file.

After typing the ticker symbol of the company the following information is inserted on excel:
- Annual Sales (trailling twelve month)
- Annual Sales (last 10-k)
- Sales Growth (trailling twelve month)
- Sales Growth (36 month)
- Total Debt
- Total Assets
- Debt to Asset ratio
- Debt-to-Ebitda ratio
- (Capex + Acquisitions)/ Revenues (I can't do this one)
- Operating Margin (trailling twelve month)
- Float
- Beta (5Y Monthly)
- PE ratio (trailling twelve month)
- Price/Book ratio (mrq)
- EV/Asset
- EV/Ebitda
- Dividend Yield (%)

Below these data, there are some verification camps on column A on excel that can be deleted afterwards, but are relevant to make sure everything is right:
- Company name is written to ensure the correct ticker symbol was typed
- Since Debt/Ebitda and EV/Assets takes values from 2 and 3 different tables respectively, thus it is important to make sure the columns in each of the tables have the same date
- Sales (ttm), Sales (last 10-K) and Sales (36M) will let you know that the data from sales equals to the same date for every company. Ideally every line should have equal values. There is an exception in these last 3 lines: some companies end its fiscal year in the middle of the year (like microsoft) thus it will appear different dates for different companies. If that is ok, then just make sure that the date shown is really the fiscal date on the SEC website for example. 

A sample output is provided
