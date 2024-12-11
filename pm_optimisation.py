#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import numpy as np
import pandas as pd
from pypfopt import EfficientFrontier, risk_models, expected_returns, objective_functions
import xlwings as xw
import yfinance as yf


# In[ ]:


def optimise_python():
    wb = xw.Book("a.xlsm")   #step1-connect to excel file
    sheet = wb.sheets['Sheet1']
    
    x1 = sheet.range('B4').value #step2- taking tickers from excel as input
    tickers = x1.split(",")
    
    data = yf.download(tickers, start='2019-01-01', end='2025-01-01')['Adj Close'] #step3- data,returns,risk
    returns = expected_returns.mean_historical_return(data)
    cov_matrix = risk_models.sample_cov(data)

    max_return = max(returns)
    min_return = min(returns)
    
    ef = EfficientFrontier(returns, cov_matrix) #step4- efficient frontier
    ef.add_objective(objective_functions.L2_reg, gamma=0.3)
    weights = ef.max_sharpe()
    cleaned_weights = ef.clean_weights()

    optimized_return = round(ef.portfolio_performance()[0]*100,2)
    optimized_risk = round(ef.portfolio_performance()[1]*100,2)
    optimized_sharpe = round(ef.portfolio_performance()[2],2)

    optimal_weights = [] #step5- convert ef weights for export
    for stock,weight in cleaned_weights.items():
        optimal_weights.append([stock,round(weight*100,2)])
    data2 = pd.DataFrame(optimal_weights)
############################################################################### step5- transferring output data back to excel
    sheet.range('A12:C27').value = ""
    sheet['A12'].value = data2

    sheet.range('A12:C12').value = ""
    sheet.range('A12:A27').value = ""

    sheet['B12'].value = "stock"
    sheet['C12'].value = "Optimal Weight" 

    sheet['B7'].value = optimized_sharpe # sharpe
    sheet['B8'].value = optimized_return # return
    sheet['B9'].value = optimized_risk # risk
################################################################################
    wb.save() #saving excel file

optimise_python()
# In[ ]:




