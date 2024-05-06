import glob
import os
import time
from datetime import datetime
import pandas as pd
from IPython.display import display
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageTemplate, Frame
from reportlab.lib import colors
from reportlab.lib.units import inch
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

PAYMENT_MODES = [
    "CASH AMT."
    "CREDIT CARDS",
    "HDFC",
    "BANDHAN",
    "GOOGLE PAY",
    "GV AMT",
    "OUTSTANDING",
    "CN AMT ADJS."
    "OTHERS",
    "CHEQUE",
]

GST_TAX_RATES = [
    "GST 03%",
    "GST 05%",
    "GST 12%",
    "GST 18%",
    "GST 28%"
]
RATES = [3,5,12,18,28]
detail_sales_sum_df = pd.read_excel("Detail_Sales_Sum.xlsx", skiprows=2)
detail_sales_sum_df.drop(columns=['S_No.', 'BRANCH NAME', 'BILL DATE', 'AGENT NAME', 'CUSTOMER NAME', 'RETAIL CUSTOMER FIRST NAME'], inplace=True)
detail_sales_sum_df.fillna('',inplace=True)

detail_sales_sum_df.drop(detail_sales_sum_df.index[-1], inplace=True) # drop last row which is just Grand Total

# detail_sales_sum_df['TAX 1(%)']= pd.to_numeric(detail_sales_sum_df['TAX 1(%)'], errors='coerce') # If anyof the row are NaN that means the bill conatins a TAX FREE Product
# detail_sales_sum_df['TAX 3(%)']= pd.to_numeric(detail_sales_sum_df['TAX 3(%)'], errors='coerce') # If anyof the rows are NaN that means the bill conatins a TAX FREE Product
# detail_sales_sum_df['GST Rate'] = detail_sales_sum_df['TAX 1(%)'] + detail_sales_sum_df['TAX 3(%)']
# detail_sales_sum_df.drop(detail_sales_sum_df[detail_sales_sum_df['GST Rate'].isna()].index, inplace=True) # If the rows are NaN then it means it is a TAXFREE product, so drop them
bill_no_dict = {}

for i, row in detail_sales_sum_df.iterrows():
    for key, value in bill_no_dict.items():
        if key in bill_no_dict:
            for i in range(len(PAYMENT_MODES)):
                if row[PAYMENT_MODES[i]] != '':
                    bill_no_dict[key]["Payment Mode"].append(i+1)
            bill_no_dict[key]["GST Rate"].append(row['GST RATE'])
            bill_no_dict[key]["Qty_GST_wise"].append(row['NET QTY'])
############# bill_no_dict[key]["Gross_Amt_GST_wise"].append(row['GROSS AMT']) ############# 
            bill_no_dict[key]["Addl Chrgs"].append(row['ADDITION CHRGS'])
            bill_no_dict[key]["Less Exch"].append(row['LESS EXCHANGE'])
            bill_no_dict[key]["Flat Disc"].append(row['FLAT DISCOUNT'])
            bill_no_dict[key]["Sales Disc"].append(row['SALES DISCOUNT'])
            for i in range(len(GST_TAX_RATES)):
                if row[GST_TAX_RATES[i]] != '':
                    bill_no_dict[key]["Gross Sales"].append(i+1)         
            bill_no_dict[key]["GST"].append(row['GST TAX'])
            bill_no_dict[key]["Net Sales"].append(row['NET SALES'])
        else:
            bill_no_dict[row['BILL NO.']] = {}
            bill_no_dict[row['BILL NO.']]["Payment Mode"] = []
            bill_no_dict[row['BILL NO.']]["GST Rate"] = []
            bill_no_dict[row['BILL NO.']]["Qty_GST_wise"] = []
            bill_no_dict[row['BILL NO.']]["Gross_Amt_GST_wise"] = []
            bill_no_dict[row['BILL NO.']]["Addl Chrgs"] = []
            bill_no_dict[row['BILL NO.']]["Less Exch"] = []
            bill_no_dict[row['BILL NO.']]["Flat Disc"] = []
            bill_no_dict[row['BILL NO.']]["Sales Disc"] = []
            bill_no_dict[row['BILL NO.']]["Gross Sales"] = []
            bill_no_dict[row['BILL NO.']]["GST"] = []
            bill_no_dict[row['BILL NO.']]["Net Sales"] = []

            for i in range(len(PAYMENT_MODES)):
                if row[PAYMENT_MODES[i]] != '':
                    bill_no_dict[row['BILL NO.']]["Payment Mode"].append(i+1)
            bill_no_dict[row['BILL NO.']]["GST Rate"].append(row['GST RATE'])
            bill_no_dict[row['BILL NO.']]["Qty_GST_wise"].append(row['NET QTY'])
############# bill_no_dict[row['BILL NO.']]["Gross_Amt_GST_wise"].append(row['GROSS AMT']) ############# 
            bill_no_dict[row['BILL NO.']]["Addl Chrgs"].append(row['ADDITION CHRGS'])
            bill_no_dict[row['BILL NO.']]["Less Exch"].append(row['LESS EXCHANGE'])
            bill_no_dict[row['BILL NO.']]["Flat Disc"].append(row['FLAT DISCOUNT'])
            bill_no_dict[row['BILL NO.']]["Sales Disc"].append(row['SALES DISCOUNT'])
            for i in range(len(GST_TAX_RATES)):
                if row[GST_TAX_RATES[i]] != '':
                    bill_no_dict[row['BILL NO.']]["Gross Sales"].append(i+1)         
            bill_no_dict[row['BILL NO.']]["GST"].append(row['GST TAX'])
            bill_no_dict[row['BILL NO.']]["Net Sales"].append(row['NET SALES'])
            
# Iterate through the dictionary again to default multi payment modes to cash or credit card dependingon the scenario
"""
If Cash and any other payment mode then give preference to cash
if CC and any other payment mode then give preference to CC
"""
for key, val in bill_no_dict.items():
    if len(val['Payment Mode']) > 1:
        for i in range(val['Payment Mode']):
            if val['Payment Mode'][i] == 1:
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[0]
                break
            if val['Payment Mode'][i] == 2:
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[1]
                break
            if val['Payment Mode'][i] == 3:
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[2]
                break
            if val['Payment Mode'][i] == 4:  
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[3]
                break            
            if val['Payment Mode'][i] == 5:   
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[4]
                break            
            if val['Payment Mode'][i] == 6:  
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[5]
                break            
            if val['Payment Mode'][i] == 7:  
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[6]
                break            
            if val['Payment Mode'][i] == 8:  
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[7]
                break            
            if val['Payment Mode'][i] == 9:  
                val['Payment Mode'].clear()
                val['Payment Mode'] = PAYMENT_MODES[8]
                break     

payment_mode_dict = {}

for key, value in bill_no_dict.items():
    if value['Payment Mode'] in payment_mode_dict:
        ### Combine all values for this payment mode GST Rate wise #####
    else:
        payment_mode_dict[value['Payment Mode']] = {}
        payment_mode_dict[value['Payment Mode']][value['GST Rate']] = {}
        for i in range(len(RATES)):
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]] = {}
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Qty'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Gross_Amt_GST_wise'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Addl Chrgs'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Less Exch'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Flat Disc'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Sales Disc'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['GST'] = 0
            payment_mode_dict[value['Payment Mode']][value['GST Rate']][RATES[i]]['Net Sales'] = 0
        
        ### Initialize the values ######

        

##################### CASH COLUMN MISSING IN TABLE ############################
display(detail_sales_sum_df.columns)
