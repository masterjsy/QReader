from openpyxl.workbook.workbook import Workbook
import pandas as pd
import yfinance as yf
import datetime as dt
import os
import requests
from bs4 import BeautifulSoup
import logging
import concurrent.futures
import time
import tkinter as tk
import tkinter.filedialog as fd

# Function start ################################################################

# Read data from Market Watch and return Operating Income and Net Income 
def getDataQuarter1YearAgo(stock):
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"}
    url="https://www.marketwatch.com/investing/stock/{}/financials/income/quarter"

    res = requests.get(url.format(stock), headers=headers)
    soup = BeautifulSoup(res.text, 'lxml')
    # print(soup)

    operating_income = 0.0
    net_income = 0.0

    # Get Operating Income
    if(soup.find("div", text="Gross Income") and soup.find("div", text="SG&A Expense")):
        # Find Gross Income
        cell = soup.find("div", text="Gross Income")
        cell = cell.parent.find_next_sibling("td")
        cell = cell.find("div", attrs={"class":"cell__content"})
        str_gross_income = cell.get_text()
        # Find SG&A Expense
        cell = soup.find("div", text="SG&A Expense")
        cell = cell.parent.find_next_sibling("td")
        cell = cell.find("div", attrs={"class":"cell__content"})
        str_expense = cell.get_text()
        # Calculate operating income
        operating_income = convertStrToFloat(str_gross_income) - convertStrToFloat(str_expense)
    elif (soup.find("div", text="Operating Income After Interest Expense")):
        cell = soup.find("div", text="Operating Income After Interest Expense")
        cell = cell.parent.find_next_sibling("td")
        cell = cell.find("div", attrs={"class":"cell__content"})
        operating_income = convertStrToFloat(cell.get_text())
    else:
        logging.debug("Cannot Get Operating Income. " + stock)

    # Get Net Income
    if(soup.find("div", text="Net Income")):
        cell = soup.find("div", text="Net Income")
        cell = cell.parent.find_next_sibling("td")
        cell = cell.find("div", attrs={"class":"cell__content"})
        str_net_income = cell.get_text()
        net_income = convertStrToFloat(str_net_income)
    else:
        logging.debug("Cannot Get Net Income. " + stock)

    return operating_income, net_income


# Convert String (17B, 15M, 1K...) to float
def convertStrToFloat(str):
    m = {'K':3, 'M':6, 'B':9, 'T':12, 'k':3, 'm':6, 'b':9, 't':12,}
    multiplier = 1

    if(str.startswith('(')):
        multiplier = -1
        str = str.replace('(', '').replace(')','')

    if(str == '-'):
        temp = 0
    elif((str[-1]).isalpha()):
        temp = float(str[:-1]) * 10**m[str[-1]]
    else:
        temp = float(str)

    temp *= multiplier

    return temp


def getExcelHeader():
    row = ["Ticker", "MarketCap", "industry", "country", "EV", "ebitda", "earningsQuarterlyGrowth", "revenueQuarterlyGrowth", "earningsGrowth", "revenueGrowth", 
    "returnOnAssets", "debtToEquity", "returnOnEquity", "totalCash", "totalDebt", "bookValue", "priceToBook", "earningsQuarterlyGrowth", "priceToSalesTrailing12Months", 
    "pegRatio", "freeCashflow", "MRQ Revenue", "MRQ Earnings", "MRY Revenue", "MRY Earnings", "MRQ Gross Profit", "MRQ Net Income", "MRQ Ebit", 
    "MRQ Operating Income", "MRQ Total Revenue", "MRQ Net Income From Continuing Ops", "MRQ Interest Expense", "MRQ-1 Gross Profit", "MRQ-1 Net Income", 
    "MRQ-1 Ebit", "MRQ-1 Operating Income", "MRQ-1 Total Revenue", "MRQ-1 Net Income From Continuing Ops", "MRY Gross Profit", "MRY Net Income", "MRY Ebit", 
    "MRY Operating Income", "MRY Total Revenue", "MRY Net Income From Continuing Ops", "MRY-1 Gross Profit", "MRY-1 Net Income", "MRY-1 Ebit", "MRY-1 Operating Income", 
    "MRY-1 Total Revenue", "MRY-1 Net Income From Continuing Ops", "MRQ Total Cash From Operating Activities", "MRQ Change In Cash", "MRQ Change To Netincome", 
    "MRQ Change To Capital Expenditures", "MRY Total Cash From Operating Activities", "MRY Change In Cash", "MRY Change To Netincome", "MRQ Total Liabilities", 
    "MRQ Total Assets", "MRQ Cash", "MRQ Total Current Liabilities", "MRQ Short Term Debt", "MRQ Total Current Assets", "MRQ Long Term Debt", 
    "MRQ Net Tangible Assets", "MRQ Total Stockholder Equity", "MRY Total Liab", "MRY Total Assets", "MRY Cash", "MRY Total Current Liabilities", 
    "MRY Short Term Debt", "MRY Total Current Assets", "MRY Long Term Debt", "MRY Net Tangible Assets", "MRY Total Stockholder Equity", "MRY-1 Total Liab", 
    "MRY-1 Total Assets", "MRY-1 Cash", "MRY-1 Total Current Liabilities", "MRY-1 Short Term Debt", "MRY-1 Total Current Assets", "MRY-1 Long Term Debt", 
    "MRY-1 Net Tangible Assets", "MRY-1 Total Stockholder Equity", "MRQ-4 Operating Income", "MRQ-4 Total Net Income"]
    return row

def getStockRowData(stock):
    logging.debug("Process start: " + stock)
    rowData = []
    rowData.append(stock)

    try:
        ticker = yf.Ticker(stock)
        # df = ticker.history(period="2y")

        appendToList(rowData, ticker.info, 'marketCap')
        appendToList(rowData, ticker.info, 'industry')
        appendToList(rowData, ticker.info, 'country')
        appendToList(rowData, ticker.info, 'enterpriseValue')
        appendToList(rowData, ticker.info, 'ebitda')
        appendToList(rowData, ticker.info, 'earningsQuarterlyGrowth')
        appendToList(rowData, ticker.info, 'revenueQuarterlyGrowth')
        appendToList(rowData, ticker.info, 'earningsGrowth')
        appendToList(rowData, ticker.info, 'revenueGrowth')
        appendToList(rowData, ticker.info, 'returnOnAssets')
        appendToList(rowData, ticker.info, 'debtToEquity')
        appendToList(rowData, ticker.info, 'returnOnEquity')
        appendToList(rowData, ticker.info, 'totalCash')
        appendToList(rowData, ticker.info, 'totalDebt')
        appendToList(rowData, ticker.info, 'bookValue')
        appendToList(rowData, ticker.info, 'priceToBook')
        appendToList(rowData, ticker.info, 'earningsQuarterlyGrowth')
        appendToList(rowData, ticker.info, 'priceToSalesTrailing12Months')
        appendToList(rowData, ticker.info, 'pegRatio')
        appendToList(rowData, ticker.info, 'freeCashflow')
        appendValueToList(rowData, ticker.quarterly_earnings['Revenue'][-1])
        appendValueToList(rowData, ticker.quarterly_earnings['Earnings'][-1])
        appendValueToList(rowData, ticker.earnings.iloc[-1]['Revenue'])
        appendValueToList(rowData, ticker.earnings.iloc[-1]['Earnings'])

        quaterlyFinancials = ticker.quarterly_financials[ticker.quarterly_financials.columns[0]]
        appendToList(rowData, quaterlyFinancials, 'Gross Profit')
        appendToList(rowData, quaterlyFinancials, 'Net Income')
        appendToList(rowData, quaterlyFinancials, 'Ebit')
        appendToList(rowData, quaterlyFinancials, 'Operating Income')
        appendToList(rowData, quaterlyFinancials, 'Total Revenue')
        appendToList(rowData, quaterlyFinancials, 'Net Income From Continuing Ops')
        appendToList(rowData, quaterlyFinancials, 'Interest Expense')

        quaterlyFinancials = ticker.quarterly_financials[ticker.quarterly_financials.columns[1]]
        appendToList(rowData, quaterlyFinancials, 'Gross Profit')
        appendToList(rowData, quaterlyFinancials, 'Net Income')
        appendToList(rowData, quaterlyFinancials, 'Ebit')
        appendToList(rowData, quaterlyFinancials, 'Operating Income')
        appendToList(rowData, quaterlyFinancials, 'Total Revenue')
        appendToList(rowData, quaterlyFinancials, 'Net Income From Continuing Ops')

        yearlyFinancials = ticker.financials[ticker.financials.columns[0]]
        appendToList(rowData, yearlyFinancials, 'Gross Profit')
        appendToList(rowData, yearlyFinancials, 'Net Income')
        appendToList(rowData, yearlyFinancials, 'Ebit')
        appendToList(rowData, yearlyFinancials, 'Operating Income')
        appendToList(rowData, yearlyFinancials, 'Total Revenue')
        appendToList(rowData, yearlyFinancials, 'Net Income From Continuing Ops')

        yearlyFinancials = ticker.financials[ticker.financials.columns[1]]
        appendToList(rowData, yearlyFinancials, 'Gross Profit')
        appendToList(rowData, yearlyFinancials, 'Net Income')
        appendToList(rowData, yearlyFinancials, 'Ebit')
        appendToList(rowData, yearlyFinancials, 'Operating Income')
        appendToList(rowData, yearlyFinancials, 'Total Revenue')
        appendToList(rowData, yearlyFinancials, 'Net Income From Continuing Ops')

        quarterlyCashflow = ticker.quarterly_cashflow[ticker.quarterly_cashflow.columns[0]]
        appendToList(rowData, quarterlyCashflow, 'Total Cash From Operating Activities')
        appendToList(rowData, quarterlyCashflow, 'Change In Cash')
        appendToList(rowData, quarterlyCashflow, 'Change To Netincome')
        appendToList(rowData, quarterlyCashflow, 'Capital Expenditures')

        yearlyCashflow = ticker.cashflow[ticker.cashflow.columns[0]]
        appendToList(rowData, yearlyCashflow, 'Total Cash From Operating Activities')
        appendToList(rowData, yearlyCashflow, 'Change In Cash')
        appendToList(rowData, yearlyCashflow, 'Change To Netincome')

        quaterlyBalancesheet = ticker.quarterly_balancesheet[ticker.quarterly_balancesheet.columns[0]]
        appendToList(rowData, quaterlyBalancesheet, 'Total Liab')
        appendToList(rowData, quaterlyBalancesheet, 'Total Assets')
        appendToList(rowData, quaterlyBalancesheet, 'Cash')
        appendToList(rowData, quaterlyBalancesheet, 'Total Current Liabilities')
        appendToList(rowData, quaterlyBalancesheet, 'Short Long Term Debt')
        appendToList(rowData, quaterlyBalancesheet, 'Total Current Assets')
        appendToList(rowData, quaterlyBalancesheet, 'Long Term Debt')
        appendToList(rowData, quaterlyBalancesheet, 'Net Tangible Assets')
        appendToList(rowData, quaterlyBalancesheet, 'Total Stockholder Equity')

        yearlyBalancesheet = ticker.balancesheet[ticker.balancesheet.columns[0]]
        appendToList(rowData, yearlyBalancesheet, 'Total Liab')
        appendToList(rowData, yearlyBalancesheet, 'Total Assets')
        appendToList(rowData, yearlyBalancesheet, 'Cash')
        appendToList(rowData, yearlyBalancesheet, 'Total Current Liabilities')
        appendToList(rowData, yearlyBalancesheet, 'Short Long Term Debt')
        appendToList(rowData, yearlyBalancesheet, 'Total Current Assets')
        appendToList(rowData, yearlyBalancesheet, 'Long Term Debt')
        appendToList(rowData, yearlyBalancesheet, 'Net Tangible Assets')
        appendToList(rowData, yearlyBalancesheet, 'Total Stockholder Equity')

        yearlyBalancesheet = ticker.balancesheet[ticker.balancesheet.columns[1]]
        appendToList(rowData, yearlyBalancesheet, 'Total Liab')
        appendToList(rowData, yearlyBalancesheet, 'Total Assets')
        appendToList(rowData, yearlyBalancesheet, 'Cash')
        appendToList(rowData, yearlyBalancesheet, 'Total Current Liabilities')
        appendToList(rowData, yearlyBalancesheet, 'Short Long Term Debt')
        appendToList(rowData, yearlyBalancesheet, 'Total Current Assets')
        appendToList(rowData, yearlyBalancesheet, 'Long Term Debt')
        appendToList(rowData, yearlyBalancesheet, 'Net Tangible Assets')
        appendToList(rowData, yearlyBalancesheet, 'Total Stockholder Equity')

        operating_income, net_income = getDataQuarter1YearAgo(stock)
        appendValueToList(rowData, operating_income)
        appendValueToList(rowData, net_income)
        logging.debug("Process End: " + stock)
        return rowData
        
    except Exception as e:
        logging.warning(stock + " " + str(e))
        logging.debug("Process End: " + stock)
        return [stock]
        

def appendToList(row, data_dic, key, index=-99):
    if key in data_dic:
        if(index == -99):
            row.append(data_dic[key])
        else:
            row.append(data_dic[key][index])
    else:
        row.append('')


def appendValueToList(row, val):
    row.append(val)


# Open Excel and Return list
def openExcelFile(filePath):  
    exportList = []
    
    extension = os.path.splitext(filePath)[1]

    if(extension == ".csv"):
        stocklist = pd.read_csv(filePath)    
    else:
        stocklist = pd.read_excel(filePath)

    return stocklist


# Save rows to Excel
def saveListToWorkbook(rows, load_file_path):
    rows.insert(0, getExcelHeader())
    temp_file_name = os.path.splitext(os.path.basename(load_file_path))[0]
    newFile = os.path.dirname(load_file_path) + "/RecordOutput_" + temp_file_name + "_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
    
    df = pd.DataFrame(rows)
    writer = pd.ExcelWriter(newFile, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='sheet', header=False, index=False)
    writer.save()    
    
    logging.debug("A result file has been created: " + newFile)

    
def processExcelFile(file_name):
    logging.debug("Reading file: " + file_name)
    exportList = openExcelFile(file_name)

    rows = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {executor.submit(getStockRowData, sym): sym for sym in exportList['Symbol']}
        for future in concurrent.futures.as_completed(futures):
            row = future.result()
            if row:
                rows.append(row)

    saveListToWorkbook(rows, file_name)


def file_open():
    fTypes = [(".xlsm", "*.xlsx", ".xls"),(".csv ", "*.csv")]
    dir1 = "c:\\"
    files = fd.askopenfilenames(filetypes = fTypes, initialdir=dir1)
    return list(files)    
    
    
# Function end ################################################################


load_files = file_open()
file_path = filePath = os.path.dirname(load_files[0]) + "/"
load_files.sort()

log_file = file_path + "/qreader_log_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ".txt"
print("log file:", log_file)
logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.DEBUG)

   
waiting_time = 5400

for index, load_file in enumerate(load_files):
    processExcelFile(load_file)
    if(index < len(load_files)-1):
        logging.debug("Sleep for {}: ".format(waiting_time) + dt.datetime.now().strftime("%Y%m%d %H:%M:%S"))
        time.sleep(waiting_time)
        logging.debug("Wake up: " + dt.datetime.now().strftime("%Y%m%d %H:%M:%S"))







