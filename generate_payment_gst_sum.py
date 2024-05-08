import glob
import os
import time
import math
from datetime import datetime
import pandas as pd
from IPython.display import display
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageTemplate, Frame, PageBreak, CondPageBreak, Spacer,KeepTogether
from reportlab.lib import colors
from reportlab.lib.units import inch
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import json


PAYMENT_MODES = [
    "CASH AMT",           # 1
    "CREDIT CARDS",       # 2
    "HDFC",               # 3
    "BANDHAN",            # 4
    "GOOGLE PAY",         # 5
    "CHEQUE",             # 6
    "GV AMT",             # 7 
    "OTHERS",             # 8
    "OUTSTANDING",        # 9
    "CN AMT ADJS.",       # 10
]

GST_TAX_RATES = [
    "GST 03%",
    "GST 05%",
    "GST 12%",
    "GST 18%",
    "GST 28%"
]
RATES = [3,5,12,18,28]

def myFirstPage(canvas, doc, title, start_date, end_date):
    canvas.saveState()
    canvas.setFont('Helvetica-Bold', 14)  # Set font size to 12 points
    canvas.drawString(20, doc.pagesize[1]+20, title)
    canvas.setFont('Helvetica-Bold', 10)  # Set font size to 12 points
    canvas.drawString(20, doc.pagesize[1], f"Detail Sales Summary from {start_date} to {end_date}")
    canvas.setFont('Helvetica-Bold', 6)  # Set font size to 12 points
    canvas.drawString(4*inch, 0.25 * inch, "Page %d" % (doc.page))
    canvas.restoreState()

def myLaterPages(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica-Bold', 6)  # Set font size to 12 points
    canvas.drawString(4*inch, 0.25 * inch, "Page %d" % (doc.page))
    canvas.restoreState()

detail_sales_sum_df = pd.read_excel("Detail_Sales_Sum.xlsx", skiprows=2)
detail_sales_sum_df.drop(columns=['S_No.', 'BRANCH NAME', 'BILL DATE', 'AGENT NAME', 'CUSTOMER NAME', 'RETAIL CUSTOMER FIRST NAME'], inplace=True)
# detail_sales_sum_df.fillna('',inplace=True)
detail_sales_sum_df.drop(detail_sales_sum_df.index[-1], inplace=True) # drop last row which is just Grand Total

detail_sales_sum_df['CGST%']= pd.to_numeric(detail_sales_sum_df['CGST%'], errors='coerce') # If anyof the row are NaN that means the bill conatins a TAX FREE Product
detail_sales_sum_df['SGST/IGST%']= pd.to_numeric(detail_sales_sum_df['SGST/IGST%'], errors='coerce') # If anyof the rows are NaN that means the bill conatins a TAX FREE Product
detail_sales_sum_df['GST RATE'] = detail_sales_sum_df['CGST%'] + detail_sales_sum_df['SGST/IGST%']
detail_sales_sum_df.drop(detail_sales_sum_df[detail_sales_sum_df['GST RATE'].isna()].index, inplace=True) # If the rows are NaN then it means it is a TAXFREE product, so drop them
detail_sales_sum_df.drop(["CGST%","SGST/IGST%"], axis=1, inplace=True)

detail_sales_sum_df["GROSS AMT"] = pd.to_numeric(detail_sales_sum_df["GROSS AMT"], errors="coerce")
detail_sales_sum_df["BILL AMT"] = pd.to_numeric(detail_sales_sum_df["BILL AMT"], errors="coerce")

detail_sales_sum_df.iloc[:,2:] = detail_sales_sum_df.iloc[:,2:].astype("float64").round(2)
########################### Fill NaN Values for GROSS AMT and BILL AMT COLUMNS ###########################################
for _ in range(5):
    try:
        null_index = detail_sales_sum_df[detail_sales_sum_df["GROSS AMT"].isna()].index
        detail_sales_sum_df.loc[null_index, "ADDITION CHRGS"] = detail_sales_sum_df["ADDITION CHRGS"].shift(1)
        detail_sales_sum_df.loc[null_index, "FLAT DISCOUNT"] = detail_sales_sum_df["FLAT DISCOUNT"].shift(1)
        detail_sales_sum_df.loc[null_index, "SALES DISCOUNT"] = detail_sales_sum_df["SALES DISCOUNT"].shift(1)
        detail_sales_sum_df.loc[null_index, "LESS EXCHANGE"] = detail_sales_sum_df["LESS EXCHANGE"].shift(1)
        detail_sales_sum_df["GROSS AMT"].fillna(detail_sales_sum_df["GROSS AMT"].shift(1), inplace= True)
        detail_sales_sum_df["BILL AMT"].fillna(detail_sales_sum_df["BILL AMT"].shift(1), inplace= True)
    except KeyError as e:
        print(e)

display(detail_sales_sum_df["GROSS AMT"].isna().value_counts())
display(detail_sales_sum_df["BILL AMT"].isna().value_counts())

detail_sales_sum_df["GrossAmt-BillAmt"] = detail_sales_sum_df["GROSS AMT"] - detail_sales_sum_df["BILL AMT"]
detail_sales_sum_df["GrossAmt-BillAmt"].fillna(0, inplace=True)

############################ CALCULATE DISCOUNTS AND ADDL CHRGS GST WISE ###########################
for i in range(len(GST_TAX_RATES)):
    try:
        detail_sales_sum_df[GST_TAX_RATES[i]] = pd.to_numeric(detail_sales_sum_df[GST_TAX_RATES[i]], errors='coerce')
        detail_sales_sum_df[f"ADDITION CHRGS {GST_TAX_RATES[i]}"] = (detail_sales_sum_df["ADDITION CHRGS"]).multiply(detail_sales_sum_df[GST_TAX_RATES[i]].divide(detail_sales_sum_df["BILL AMT"]))
        detail_sales_sum_df[f"SALES DISCOUNT {GST_TAX_RATES[i]}"] = (detail_sales_sum_df["SALES DISCOUNT"]).multiply(detail_sales_sum_df[GST_TAX_RATES[i]].divide(detail_sales_sum_df["BILL AMT"]))
        detail_sales_sum_df[f"FLAT DISCOUNT {GST_TAX_RATES[i]}"] = (detail_sales_sum_df["FLAT DISCOUNT"]).multiply(detail_sales_sum_df[GST_TAX_RATES[i]].divide(detail_sales_sum_df["BILL AMT"]))
        detail_sales_sum_df[f"LESS EXCHANGE {GST_TAX_RATES[i]}"] = (detail_sales_sum_df["LESS EXCHANGE"]).multiply(detail_sales_sum_df[GST_TAX_RATES[i]].divide(detail_sales_sum_df["BILL AMT"]))
    except KeyError as e:
        print(f"No {GST_TAX_RATES[i]} products were bought today :(")

##################### FILL NaN Values for Addl Charges, Discounts and exchanges ###########################

for i in range(len(GST_TAX_RATES)-4):
    try:
        detail_sales_sum_df["ADDITION CHRGS GST WISE"] = detail_sales_sum_df[f"ADDITION CHRGS {GST_TAX_RATES[i]}"].fillna(detail_sales_sum_df[f"ADDITION CHRGS {GST_TAX_RATES[i+1]}"])
        detail_sales_sum_df["SALES DISCOUNT GST WISE"] = detail_sales_sum_df[f"SALES DISCOUNT {GST_TAX_RATES[i]}"].fillna(detail_sales_sum_df[f"SALES DISCOUNT {GST_TAX_RATES[i+1]}"])
        detail_sales_sum_df["FLAT DISCOUNT GST WISE"] = detail_sales_sum_df[f"FLAT DISCOUNT {GST_TAX_RATES[i]}"].fillna(detail_sales_sum_df[f"FLAT DISCOUNT {GST_TAX_RATES[i+1]}"])
        detail_sales_sum_df["LESS EXCHANGE GST WISE"] = detail_sales_sum_df[f"LESS EXCHANGE {GST_TAX_RATES[i]}"].fillna(detail_sales_sum_df[f"LESS EXCHANGE {GST_TAX_RATES[i+1]}"])

        detail_sales_sum_df["ADDITION CHRGS GST WISE"].fillna(detail_sales_sum_df[f"ADDITION CHRGS {GST_TAX_RATES[i+2]}"], inplace=True)
        detail_sales_sum_df["SALES DISCOUNT GST WISE"].fillna(detail_sales_sum_df[f"SALES DISCOUNT {GST_TAX_RATES[i+2]}"], inplace=True)
        detail_sales_sum_df["FLAT DISCOUNT GST WISE"].fillna(detail_sales_sum_df[f"FLAT DISCOUNT {GST_TAX_RATES[i+2]}"], inplace=True)
        detail_sales_sum_df["LESS EXCHANGE GST WISE"].fillna(detail_sales_sum_df[f"LESS EXCHANGE {GST_TAX_RATES[i+2]}"], inplace=True)

        detail_sales_sum_df["ADDITION CHRGS GST WISE"].fillna(detail_sales_sum_df[f"ADDITION CHRGS {GST_TAX_RATES[i+3]}"], inplace=True)
        detail_sales_sum_df["SALES DISCOUNT GST WISE"].fillna(detail_sales_sum_df[f"SALES DISCOUNT {GST_TAX_RATES[i+3]}"], inplace=True)
        detail_sales_sum_df["FLAT DISCOUNT GST WISE"].fillna(detail_sales_sum_df[f"FLAT DISCOUNT {GST_TAX_RATES[i+3]}"], inplace=True)
        detail_sales_sum_df["LESS EXCHANGE GST WISE"].fillna(detail_sales_sum_df[f"LESS EXCHANGE {GST_TAX_RATES[i+3]}"], inplace=True
                                                    )
        detail_sales_sum_df["ADDITION CHRGS GST WISE"].fillna(detail_sales_sum_df[f"ADDITION CHRGS {GST_TAX_RATES[i+4]}"], inplace=True)
        detail_sales_sum_df["SALES DISCOUNT GST WISE"].fillna(detail_sales_sum_df[f"SALES DISCOUNT {GST_TAX_RATES[i+4]}"], inplace=True)
        detail_sales_sum_df["FLAT DISCOUNT GST WISE"].fillna(detail_sales_sum_df[f"FLAT DISCOUNT {GST_TAX_RATES[i+4]}"], inplace=True)
        detail_sales_sum_df["LESS EXCHANGE GST WISE"].fillna(detail_sales_sum_df[f"LESS EXCHANGE {GST_TAX_RATES[i+4]}"], inplace=True)

    except KeyError as e:
        print("")
############################ CALCULATE GST TAXES AND GROSS MRP GST WISE##############################
for i in range(len(GST_TAX_RATES)):
    try:
        detail_sales_sum_df[f"{GST_TAX_RATES[i]} TAX"] = detail_sales_sum_df[GST_TAX_RATES[i]] * (RATES[i]/(100 + RATES[i]))
        detail_sales_sum_df[f"GROSS {GST_TAX_RATES[i]}"] = (detail_sales_sum_df[GST_TAX_RATES[i]]).add((detail_sales_sum_df[GST_TAX_RATES[i]].divide(detail_sales_sum_df["BILL AMT"])).multiply(detail_sales_sum_df["GrossAmt-BillAmt"]))
    
    except KeyError as e:
        print("")
        
detail_sales_sum_df.iloc[:,2:] = detail_sales_sum_df.iloc[:,2:].astype("float64").round(2)
# display(detail_sales_sum_df[GST_TAX_RATES[2]])

################# FILL NaN Values for GST TAX and GROSS SALES and GROSS MRP ###################

for i in range(len(GST_TAX_RATES)-4):
    try:
        detail_sales_sum_df["GST TAX"] = detail_sales_sum_df[f"{GST_TAX_RATES[i]} TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+1]} TAX"])
        detail_sales_sum_df["GROSS SALES"] = detail_sales_sum_df[GST_TAX_RATES[i]].fillna(detail_sales_sum_df[GST_TAX_RATES[i+1]])
        detail_sales_sum_df["GROSS MRP"] = detail_sales_sum_df[f"GROSS {GST_TAX_RATES[i]}"].fillna(detail_sales_sum_df[f"GROSS {GST_TAX_RATES[i+1]}"])

        detail_sales_sum_df["GST TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+2]} TAX"],inplace=True)
        detail_sales_sum_df["GROSS SALES"].fillna(detail_sales_sum_df[GST_TAX_RATES[i+2]],inplace=True)
        detail_sales_sum_df["GROSS MRP"].fillna(detail_sales_sum_df[f"GROSS {GST_TAX_RATES[i+2]}"], inplace=True)

        detail_sales_sum_df["GST TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+3]} TAX"],inplace=True)
        detail_sales_sum_df["GROSS SALES"].fillna(detail_sales_sum_df[GST_TAX_RATES[i+3]],inplace=True)
        detail_sales_sum_df["GROSS MRP"].fillna(detail_sales_sum_df[f"GROSS {GST_TAX_RATES[i+3]}"], inplace=True)

        detail_sales_sum_df["GST TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+4]} TAX"],inplace=True)
        detail_sales_sum_df["GROSS SALES"].fillna(detail_sales_sum_df[GST_TAX_RATES[i+4]],inplace=True)        
        detail_sales_sum_df["GROSS MRP"].fillna(detail_sales_sum_df[f"GROSS {GST_TAX_RATES[i+4]}"], inplace=True)

    except KeyError as e:
        print("")

################################################################################

########################### CALCULATE NET SALES AND GROSS MRP GST WISE#################################

detail_sales_sum_df["NET SALES"] = detail_sales_sum_df["GROSS SALES"] - detail_sales_sum_df["GST TAX"]

#################################################################################

bill_no_dict = {}
detail_sales_sum_df['Payment Modes'] = 0
# display(detail_sales_sum_df['CREDIT CARDS'])
for i, row in detail_sales_sum_df.iterrows():
    bill_no = row['BILL NO.']
    if bill_no in bill_no_dict:
        try:
            for i in range(len(PAYMENT_MODES)):
                if not math.isnan(row[PAYMENT_MODES[i]]):
                    bill_no_dict[bill_no]["Payment Modes"].append(i+1)
                    bill_no_dict[bill_no]["Recieved Amt"].append(row[PAYMENT_MODES[i]])
        except KeyError as e:
            print(f"No Payments were made using {PAYMENT_MODES[i]}")  
    else:
        bill_no_dict[bill_no] = {}
        bill_no_dict[bill_no]["Payment Modes"] = []
        bill_no_dict[bill_no]["Recieved Amt"] = []
        bill_no_dict[bill_no]["Bill Amt"] = row["BILL AMT"]
        bill_no_dict[bill_no]["Flag"] = 0
        try:
            for i in range(len(PAYMENT_MODES)):
                if not math.isnan(row[PAYMENT_MODES[i]]):
                    bill_no_dict[bill_no]["Payment Modes"].append(i+1)
                    bill_no_dict[bill_no]["Recieved Amt"].append(row[PAYMENT_MODES[i]])
        except KeyError as e:
            print(f"No Payments were made using {PAYMENT_MODES[i]}")    

########### Check if Recieved AMt is equal to Bill Amount ############
for key, val in bill_no_dict.items():
    total = 0.0
    for i in range(len(val["Recieved Amt"])):
        total += val["Recieved Amt"][i]
    if total != val["Bill Amt"]:
        val["Flag"] = 1
        print("Recieved Amt and Bill Amt doesn't match! for bill no: ", key)

# Iterate through the dictionary again to default multi payment modes to cash or credit card dependingon the scenario
"""
If Cash and any other payment mode then give preference to cash
if CC and any other payment mode then give preference to CC
"""
for key, val in bill_no_dict.items():
    if len(val['Payment Modes']) > 1:
        for i in range(len(val['Payment Modes'])):
            if val['Payment Modes'][i] == 1:
                val['Payment Modes'].clear()
                val['Payment Modes'].append(1)
                break
            if val['Payment Modes'][i] == 2:
                val['Payment Modes'].clear()
                val['Payment Modes'].append(2)
                break
            if val['Payment Modes'][i] == 3:
                val['Payment Modes'].clear()
                val['Payment Modes'].append(3)
                break
            if val['Payment Modes'][i] == 4:  
                val['Payment Modes'].clear()
                val['Payment Modes'].append(4)
                break            
            if val['Payment Modes'][i] == 5:   
                val['Payment Modes'].clear()
                val['Payment Modes'].append(5)
                break            
            if val['Payment Modes'][i] == 6:  
                val['Payment Modes'].clear()
                val['Payment Modes'].append(6)
                break            
            if val['Payment Modes'][i] == 7:  
                val['Payment Modes'].clear()
                val['Payment Modes'].append(7)
                break            
            if val['Payment Modes'][i] == 8:  
                val['Payment Modes'].clear()
                val['Payment Modes'].append(8)
                break            
            if val['Payment Modes'][i] == 9:  
                val['Payment Modes'].clear()
                val['Payment Modes'].append(9)
                break     
            if val['Payment Modes'][i] == 10:  
                val['Payment Modes'].clear()
                val['Payment Modes'].append(10)
                break
with open('bill_no_dict.json','w') as f:
    json.dump(bill_no_dict,f, indent=4)  


for i, row in detail_sales_sum_df.iterrows():
    try:
        detail_sales_sum_df.loc[detail_sales_sum_df.index[i],"Payment Modes"] = bill_no_dict[row['BILL NO.']]["Payment Modes"][0]
    except IndexError as e:
        print("No payment info for Bill No: ", row['BILL NO.'])


for i in range(len(GST_TAX_RATES)):
    try:
        detail_sales_sum_df.drop([GST_TAX_RATES[i]], axis=1, inplace=True)
        detail_sales_sum_df.drop([f"{GST_TAX_RATES[i]} TAX"], axis=1, inplace=True)
        detail_sales_sum_df.drop([f"GROSS {GST_TAX_RATES[i]}"], axis=1, inplace=True)
    except KeyError as e:
        print("")

for i in range(len(PAYMENT_MODES)):
    try:
        detail_sales_sum_df.drop([PAYMENT_MODES[i]], axis=1, inplace=True)
    except KeyError as e:
        print("")

detail_sales_sum_df.drop(["BILL NO.","CGST","SGST/IGST","ROUND AMOUNT"], axis=1, inplace=True)
output_cols = [col for col in detail_sales_sum_df.columns if 'OUTPUT' in col]
detail_sales_sum_df.drop(output_cols, axis=1, inplace=True)
detail_sales_sum_df.drop("GST SALE TAXFREE", axis=1, inplace=True)
detail_sales_sum_df.to_excel('detail_sales_sum_df.xlsx')

gst_sum_dict = {}
payment_modes = sorted(detail_sales_sum_df['Payment Modes'].unique())

######################## Pre Initialize the dictionary ###########################
for x in range(len(payment_modes)):

    # print(payment_modes[x].dtype)
    key = int(payment_modes[x])
    gst_sum_dict[payment_modes[x]] = {}
    for i in range(len(RATES)):
        key_gst = int(RATES[i])
        gst_sum_dict[key][key_gst] = {}
        gst_sum_dict[key][key_gst]['Quantity'] = 0.0
        gst_sum_dict[key][key_gst]['Gross Mrp'] = 0.0
        gst_sum_dict[key][key_gst]['Addl Chrgs'] = 0.0
        gst_sum_dict[key][key_gst]['Sales Disc'] = 0.0
        gst_sum_dict[key][key_gst]['Flat Disc'] = 0.0
        gst_sum_dict[key][key_gst]['Less Exch'] = 0.0
        gst_sum_dict[key][key_gst]['Gross Sales'] = 0.0
        gst_sum_dict[key][key_gst]['GST'] = 0.0
        gst_sum_dict[key][key_gst]['Net Sales'] = 0.0

    gst_sum_dict[key]['Total'] = {}
    gst_sum_dict[key]['Total']['Quantity'] = 0.0
    gst_sum_dict[key]['Total']['Gross Mrp'] = 0.0
    gst_sum_dict[key]['Total']['Addl Chrgs'] = 0.0
    gst_sum_dict[key]['Total']['Sales Disc'] = 0.0
    gst_sum_dict[key]['Total']['Flat Disc'] = 0.0
    gst_sum_dict[key]['Total']['Less Exch'] = 0.0
    gst_sum_dict[key]['Total']['Gross Sales'] = 0.0
    gst_sum_dict[key]['Total']['GST'] = 0.0
    gst_sum_dict[key]['Total']['Net Sales'] = 0.0

######################## Initialize Grand Total #############################
gst_sum_dict['Grand Total'] = {}
for i in range(len(RATES)):
        key_gst = int(RATES[i])
        gst_sum_dict['Grand Total'][key_gst] = {}
        gst_sum_dict['Grand Total'][key_gst]['Quantity'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Gross Mrp'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Addl Chrgs'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Sales Disc'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Flat Disc'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Less Exch'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Gross Sales'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['GST'] = 0.0
        gst_sum_dict['Grand Total'][key_gst]['Net Sales'] = 0.0

gst_sum_dict['Grand Total']['Total'] = {}
gst_sum_dict['Grand Total']['Total']['Quantity'] = 0.0
gst_sum_dict['Grand Total']['Total']['Gross Mrp'] = 0.0
gst_sum_dict['Grand Total']['Total']['Addl Chrgs'] = 0.0
gst_sum_dict['Grand Total']['Total']['Sales Disc'] = 0.0
gst_sum_dict['Grand Total']['Total']['Flat Disc'] = 0.0
gst_sum_dict['Grand Total']['Total']['Less Exch'] = 0.0
gst_sum_dict['Grand Total']['Total']['Gross Sales'] = 0.0
gst_sum_dict['Grand Total']['Total']['GST'] = 0.0
gst_sum_dict['Grand Total']['Total']['Net Sales'] = 0.0

detail_sales_sum_df.fillna(round(0,2),inplace=True)

for i, row in detail_sales_sum_df.iterrows():
    key = int(row['Payment Modes'])
    gst_rate = row['GST RATE']
    qty = row['NET QTY']
    gross_mrp = row['GROSS MRP']
    addl_chrgs = row['ADDITION CHRGS GST WISE']
    sales_disc = row['SALES DISCOUNT GST WISE']
    flat_disc = row['FLAT DISCOUNT GST WISE']
    less_exch= row['LESS EXCHANGE GST WISE']
    gross_sales = row['GROSS SALES']
    gst_tax = row['GST TAX']
    net_sales = row['NET SALES']

    gst_sum_dict[key][gst_rate]['Quantity'] += qty
    gst_sum_dict[key][gst_rate]['Gross Mrp'] += gross_mrp
    gst_sum_dict[key][gst_rate]['Addl Chrgs'] += addl_chrgs
    gst_sum_dict[key][gst_rate]['Sales Disc'] += sales_disc
    gst_sum_dict[key][gst_rate]['Flat Disc'] += flat_disc
    gst_sum_dict[key][gst_rate]['Less Exch'] += less_exch
    gst_sum_dict[key][gst_rate]['Gross Sales'] += gross_sales
    gst_sum_dict[key][gst_rate]['GST'] += gst_tax
    gst_sum_dict[key][gst_rate]['Net Sales'] += net_sales

########################## CALCULATE TOTAL #######################
for key, value in gst_sum_dict.items():

    qty = 0.0
    gross_mrp = 0.0
    addl_chrgs = 0.0
    sales_disc = 0.0
    flat_disc = 0.0
    less_exch = 0.0
    gross_sales = 0.0
    gst_tax = 0.0
    net_sales = 0.0

    for k, v in value.items():
        qty += v['Quantity']
        gross_mrp += v['Gross Mrp']
        addl_chrgs += v['Addl Chrgs']
        sales_disc += v['Sales Disc']
        flat_disc += v['Flat Disc']
        less_exch += v['Less Exch']
        gross_sales += v['Gross Sales']
        gst_tax += v['GST']
        net_sales += v['Net Sales']

    value['Total']['Quantity'] = int(qty)
    value['Total']['Gross Mrp'] = gross_mrp
    value['Total']['Addl Chrgs'] = addl_chrgs
    value['Total']['Sales Disc'] = sales_disc
    value['Total']['Flat Disc'] = flat_disc
    value['Total']['Less Exch'] = less_exch
    value['Total']['Gross Sales'] = gross_sales
    value['Total']['GST'] = gst_tax
    value['Total']['Net Sales'] = net_sales
    
################### CALCULATE GRAND TOTAL #########################
for i in range(len(RATES)):
    gst_rate = RATES[i]
    qty = 0.0
    gross_mrp = 0.0
    addl_chrgs = 0.0
    sales_disc = 0.0
    flat_disc = 0.0
    less_exch = 0.0
    gross_sales = 0.0
    gst_tax = 0.0
    net_sales = 0.0

    for key, value in gst_sum_dict.items():
        qty += value[gst_rate]['Quantity']
        gross_mrp += value[gst_rate]['Gross Mrp']
        addl_chrgs += value[gst_rate]['Addl Chrgs']
        sales_disc += value[gst_rate]['Sales Disc']
        flat_disc += value[gst_rate]['Flat Disc']
        less_exch += value[gst_rate]['Less Exch']
        gross_sales += value[gst_rate]['Gross Sales']
        gst_tax += value[gst_rate]['GST']
        net_sales += value[gst_rate]['Net Sales']

    gst_sum_dict['Grand Total'][gst_rate] = {}
    gst_sum_dict['Grand Total'][gst_rate]['Quantity'] = qty
    gst_sum_dict['Grand Total'][gst_rate]['Gross Mrp'] = gross_mrp
    gst_sum_dict['Grand Total'][gst_rate]['Addl Chrgs'] = addl_chrgs
    gst_sum_dict['Grand Total'][gst_rate]['Sales Disc'] = sales_disc
    gst_sum_dict['Grand Total'][gst_rate]['Flat Disc'] = flat_disc
    gst_sum_dict['Grand Total'][gst_rate]['Less Exch'] = less_exch
    gst_sum_dict['Grand Total'][gst_rate]['Gross Sales'] = gross_sales
    gst_sum_dict['Grand Total'][gst_rate]['GST'] = gst_tax
    gst_sum_dict['Grand Total'][gst_rate]['Net Sales'] = net_sales

################# CALCULATE TOTAL OF GRAND TOTAL #################################    
qty = 0.0
gross_mrp = 0.0
addl_chrgs = 0.0
sales_disc = 0.0
flat_disc = 0.0
less_exch = 0.0
gross_sales = 0.0
gst_tax = 0.0
net_sales = 0.0

for key, v in gst_sum_dict['Grand Total'].items():
    qty += v['Quantity']
    gross_mrp += v['Gross Mrp']
    addl_chrgs += v['Addl Chrgs']
    sales_disc += v['Sales Disc']
    flat_disc += v['Flat Disc']
    less_exch += v['Less Exch']
    gross_sales += v['Gross Sales']
    gst_tax += v['GST']
    net_sales += v['Net Sales']

gst_sum_dict['Grand Total']['Total']['Quantity'] = int(qty)
gst_sum_dict['Grand Total']['Total']['Gross Mrp'] = gross_mrp
gst_sum_dict['Grand Total']['Total']['Addl Chrgs'] = addl_chrgs
gst_sum_dict['Grand Total']['Total']['Sales Disc'] = sales_disc
gst_sum_dict['Grand Total']['Total']['Flat Disc'] = flat_disc
gst_sum_dict['Grand Total']['Total']['Less Exch'] = less_exch
gst_sum_dict['Grand Total']['Total']['Gross Sales'] = gross_sales
gst_sum_dict['Grand Total']['Total']['GST'] = gst_tax
gst_sum_dict['Grand Total']['Total']['Net Sales'] = net_sales

print_gst_sum_dict = {str(key): value for key, value in gst_sum_dict.items()}

with open('gst_sum_dict.json','w') as f:
    json.dump(print_gst_sum_dict,f, indent=4)  

table_data = []         
pdf_table = []
table_style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.white),  # Merge all columns in the first row
    ('FONTSIZE', (0, 0), (-1,0),12),
    ('FONTSIZE', (0, 0), (-1,-1),11),
    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),  # First Column
    ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),  # Second Row font
    ('COLWIDTH', (0, 0), (-1, -1), [100,100,100,100,100,100,100]),  # Set column width for all columns  
    ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Top border
    ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),  # Bottom border
    ('LINEBEFORE', (0, 0), (0, -1), 1, colors.black),  # Left border
    ('LINEAFTER', (-1, 0), (-1, -1), 1, colors.black),  # Right border
    ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
    ('SPAN', (0,0), (6,0)),
    # ('INNERGRID', (0, 0), (-1, 0), 0, colors.white),  # Remove grid lines
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Text alignment
])

for key, value in gst_sum_dict.items():

    table_data = []
    if key != 0:
        try:
            payment_modes = [PAYMENT_MODES[key-1], ""*7]
        except TypeError as e:
            print("Reached Grand Total")
            payment_modes = ["GRAND TOTAL", ""*7]

        table_data.append(payment_modes)

        gst_header = ["GST RATE", "3%", "5%", "12%", "18%", "28%", "Total"]

        quantity_row = ["Quantity"]
        gross_amt_row = ["MRP"]
        addl_charges_row = ["Addl Charges"]
        sales_disc_row = ["Sales Disc."]
        flat_disc_row = ["Flat Disc."]
        less_exch_row = ["(-) Exch"]
        gross_sales_row = ["Gross Sales"]
        gst_row = ["GST"]
        net_sales_row = ["Net Sales"]

        for k, v in value.items():
            if v['Quantity'] != 0.0:
                quantity_row.append(int(v['Quantity']))
            else:
                quantity_row.append("")

            if v['Gross Mrp'] != 0.0:
                gross_amt_row.append(round(v['Gross Mrp'],2))
            else:
                gross_amt_row.append("")

            if v['Addl Chrgs'] != 0.0:
                addl_charges_row.append(round(v['Addl Chrgs'],2))
            else:
                addl_charges_row.append("")

            if v['Sales Disc'] != 0.0:
                sales_disc_row.append(round(v['Sales Disc'],2))
            else:
                sales_disc_row.append("")

            if v['Flat Disc'] != 0.0:    
                flat_disc_row.append(round(v['Flat Disc'],2))
            else:
                flat_disc_row.append("")

            if v['Less Exch']!= 0.0:    
                less_exch_row.append(round(v['Less Exch'],2))
            else:
                less_exch_row.append("")

            if v['Gross Sales'] != 0.0:    
                gross_sales_row.append(round(v['Gross Sales'],2))
            else:
                gross_sales_row.append("")

            if v['GST'] != 0.0:    
                gst_row.append(round(v['GST'],2))
            else:
                gst_row.append("")

            if v['Net Sales'] != 0.0:    
                net_sales_row.append(round(v['Net Sales'],2))
            else:
                net_sales_row.append("")


        table_data.append(gst_header)
        table_data.append(quantity_row)
        table_data.append(gross_amt_row)
        table_data.append(addl_charges_row)
        table_data.append(sales_disc_row)
        table_data.append(flat_disc_row)
        table_data.append(less_exch_row)
        table_data.append(gross_sales_row)
        table_data.append(gst_row)
        table_data.append(net_sales_row)
        display(addl_charges_row)
        table = Table(table_data,colWidths=80)
        # table.hAlign = 'LEFT'
        table.setStyle(table_style)
        spacer = 1 * inch  # Adjust this value according to your preference        
        pdf_table.append(table)
        # if key == list(gst_sum_dict.keys())[-2]:
        #     print("Grand Total")
        #     pdf_table.append(PageBreak())
        # else:
        pdf_table.append(Spacer(1,25))    
        pdf_table.append(CondPageBreak(150))    

pdf = SimpleDocTemplate("GST_payment_summary.pdf", pagesize=letter)

frameT = Frame(x1=0,y1=0,width=8.3*inch, height=11.7*inch,topPadding=60, bottomPadding=0.25*inch,leftPadding=0,rightPadding=0)
pdf.addPageTemplates([PageTemplate(id='First',frames=frameT, onPage=lambda canvas, doc: myFirstPage(canvas, doc, "Devaa Annex", "25/07/2024", "27/07/2024"), pagesize=A4)])
frameT = Frame(x1=0,y1=0,width=8.3*inch, height=11.7*inch,topPadding=0.5*inch, bottomPadding=0.25*inch,leftPadding=0,rightPadding=0)
pdf.addPageTemplates([PageTemplate(id='Later',frames=frameT, onPage=myLaterPages,pagesize=A4)])

pdf.build(pdf_table)
