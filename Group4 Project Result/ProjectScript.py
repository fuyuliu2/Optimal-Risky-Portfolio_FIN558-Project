# -*- coding: utf-8 -*-
"""
Created on Sun Nov 12 21:17:59 2023

@author: 
    Zhuoping Fan
    Yuehan Yu
    Fuyu Liu
"""
#%% Package Import
import pandas as pd
import numpy as np

#%% Part A. Data Preparation.

# 1. Open the spreadsheet of the constituent companies returns in python to make a pivot table of return
# (sp500_rtns). Convert dates in datetime format.
sp500_rtns = pd.read_excel('ProjectData.xlsx', index_col=0, parse_dates=True, sheet_name='Constituents monthly return')
sp500_rtns = sp500_rtns.pivot_table(index='MthCalDt', columns='PERMNO', values='MthRet')
# Remove stocks that do not have complete data.
nonna = sp500_rtns.columns[sp500_rtns.count() == len(sp500_rtns)]
sp500_rtns = sp500_rtns[nonna]

# 2. Compute average monthly return of S&P500 constituent companies from the pivot table.
avmth_return = sp500_rtns.mean()
avmth_return

# 3. Construct variance and covariance matrix of returns of the constituent companies from the pivot table.
varmatrix = sp500_rtns.var()
varmatrix
covarmatrix = sp500_rtns.cov()
covarmatrix

# 4. Save the number of constituents to a variable ‘N’.
N = len(sp500_rtns.columns)
N

#%% Part B. Custom functions.

# 5. Write a function that returns the expected return of a portfolio: F_PortRtn(r_i , w_i)
# a. Input parameters are average returns of stocks and their weights in a portfolio.
# b. It should return the expected return of a portfolio.
def F_PortRtn(r_i, w_i):
    return np.dot(r_i, w_i)

# 6. Write a function that computes the standard deviation of a portfolio: F_PortStd(cov, w_i)
# a. Two input parameters: 1) cov: variance-covariance matrix of stock returns, 2) weights of stocks in a portfolio.
# b. Hint: matrix multiplication is easier than using ‘for’ or ‘while’ loop.
def F_PortStd(cov, w_i):
    return np.sqrt(np.dot(w_i, np.dot(cov, w_i)))

# 7. Write a function that returns the Sharpe ratio of a portfolio: F_Sharpe(r_p, r_f, s_p). 𝐸[𝑟𝑝 − 𝑟𝑓]/𝜎𝑝
# a. Three input parameters: 1) r_p: return of the portfolio (Output from F_PortRtn), 2) r_f: risk free
# asset return, 3) s_p: standard deviation of a portfolio (Output from F_PortStd).
# b. For the numerator of the Sharpe ratio, it should compute the excess returns of each month to find the average.
# c. Denominator is the standard deviation of the portfolio returns.
def F_Sharpe(r_p, r_f, s_p):
    excess_returns = r_p - r_f
    return np.mean(excess_returns) / s_p

#%% Part C. Simulation

# 8. Generate 1,000 sets of ‘N’ random numbers following uniform distribution between 0 and 1. Set the seed
# of the random number generator as the lowest student ID number of your group members’. In each set,
# divide random numbers by the total of the random numbers in the set. Now, you have random weights of stocks.
np.random.seed(659824337)  
random_weights = np.random.uniform(0, 1, (1000, N))
random_weights /= random_weights.sum(axis=1)[:, np.newaxis]

# 9. Using ‘for’ or ‘while’ loop, compute Sharpe ratio of each simulated portfolio.
# a. Maintain good records of portfolio weight combination and Sharpe ratios.
# b. You need to use functions in Part B.
# c. You must have 1,000 Sharpe ratios.
sharpe_ratios = []
rf = pd.read_excel('ProjectData.xlsx', index_col=0, parse_dates=True, sheet_name='30-days TB monthly return')

for i in range(1000):
    w = random_weights[i]
    r = F_PortRtn(avmth_return, w)
    std = F_PortStd(covarmatrix, w)
    sr = F_Sharpe(r, rf.mean(), std)
    sharpe_ratios.append(sr)

# 10. Find the maximum Sharpe ratio and the corresponding weights of constituent stocks.
maxsr = np.argmax(sharpe_ratios)
max_sharpe_ratio = sharpe_ratios[maxsr]
max_sharpe_ratio
optimal_weights = random_weights[maxsr]
optimal_weights

# 11. Report the following in a Separate MS Excel spreadsheet.
# a. Optimal portfolio’s constituent weights. Provide their “permno” and names.
optimal_portfolio = pd.DataFrame({'Permno': sp500_rtns.columns, 'Weight': optimal_weights, 'Sharpe Ratio':max_sharpe_ratio})
optimal_portfolio.to_excel('optimal_portfolio.xlsx', index=False)





