import glob
import os
import time
import math
from datetime import datetime
import pandas as pd
from IPython.display import display
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageTemplate, Frame, PageBreak, CondPageBreak, Spacer
from reportlab.lib import colors
from reportlab.lib.units import inch
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import json
from flatten_dict import flatten

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

detail_sales_sum_df = pd.read_excel("Detail_Sales_Sum.xlsx", skiprows=2)
detail_sales_sum_df.drop(columns=['S_No.', 'BRANCH NAME', 'BILL DATE', 'AGENT NAME', 'CUSTOMER NAME', 'RETAIL CUSTOMER FIRST NAME'], inplace=True)
# detail_sales_sum_df.fillna('',inplace=True)
detail_sales_sum_df.drop(detail_sales_sum_df.index[-1], inplace=True) # drop last row which is just Grand Total

detail_sales_sum_df['CGST%']= pd.to_numeric(detail_sales_sum_df['CGST%'], errors='coerce') # If anyof the row are NaN that means the bill conatins a TAX FREE Product
detail_sales_sum_df['SGST/IGST%']= pd.to_numeric(detail_sales_sum_df['SGST/IGST%'], errors='coerce') # If anyof the rows are NaN that means the bill conatins a TAX FREE Product
detail_sales_sum_df['GST RATE'] = detail_sales_sum_df['CGST%'] + detail_sales_sum_df['SGST/IGST%']
detail_sales_sum_df.drop(detail_sales_sum_df[detail_sales_sum_df['GST RATE'].isna()].index, inplace=True) # If the rows are NaN then it means it is a TAXFREE product, so drop them
detail_sales_sum_df.drop(["CGST%","SGST/IGST%"], axis=1, inplace=True)

############################ CALCULATE GST TAXES ##############################
for i in range(len(GST_TAX_RATES)):
    try:
        detail_sales_sum_df[GST_TAX_RATES[i]] = pd.to_numeric(detail_sales_sum_df[GST_TAX_RATES[i]], errors='coerce')
        detail_sales_sum_df[f"{GST_TAX_RATES[i]} TAX"] = detail_sales_sum_df[GST_TAX_RATES[i]] * (RATES[i]/(100 + RATES[i]))
    except KeyError as e:
        print(f"No {GST_TAX_RATES[i]} products were bought today :(")
        
detail_sales_sum_df.iloc[:,2:] = detail_sales_sum_df.iloc[:,2:].astype("float64").round(2)
display(detail_sales_sum_df[GST_TAX_RATES[2]])

for i in range(len(GST_TAX_RATES)-4):
    try:
        detail_sales_sum_df["GST TAX"] = detail_sales_sum_df[f"{GST_TAX_RATES[i]} TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+1]} TAX"])
        detail_sales_sum_df["GROSS SALES"] = detail_sales_sum_df[GST_TAX_RATES[i]].fillna(detail_sales_sum_df[GST_TAX_RATES[i+1]])

        detail_sales_sum_df["GST TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+2]} TAX"],inplace=True)
        detail_sales_sum_df["GROSS SALES"].fillna(detail_sales_sum_df[GST_TAX_RATES[i+2]],inplace=True)

        detail_sales_sum_df["GST TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+3]} TAX"],inplace=True)
        detail_sales_sum_df["GROSS SALES"].fillna(detail_sales_sum_df[GST_TAX_RATES[i+3]],inplace=True)

        detail_sales_sum_df["GST TAX"].fillna(detail_sales_sum_df[f"{GST_TAX_RATES[i+4]} TAX"],inplace=True)
        detail_sales_sum_df["GROSS SALES"].fillna(detail_sales_sum_df[GST_TAX_RATES[i+4]],inplace=True)        

    except KeyError as e:
        print("")

################################################################################

########################### CALCULATE NET SALES #################################

detail_sales_sum_df["NET SALES"] = detail_sales_sum_df["GROSS SALES"] - detail_sales_sum_df["GST TAX"]

#################################################################################

bill_no_dict = {}
detail_sales_sum_df['Payment Modes'] = 0
display(detail_sales_sum_df['CREDIT CARDS'])
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
display(detail_sales_sum_df)
detail_sales_sum_df.to_excel('detail_sales_sum_df.xlsx')

gst_sum_dict = {}
payment_modes = sorted(detail_sales_sum_df['Payment Modes'].unique())
gst_rates = sorted(detail_sales_sum_df['GST RATE'].unique())
display(sorted(gst_rates))
######################## Pre Initialize the dictionary ###########################
for x in range(len(payment_modes)):

    # print(payment_modes[x].dtype)
    key = int(payment_modes[x])
    gst_sum_dict[payment_modes[x]] = {}
    for i in range(len(gst_rates)):
        key_gst = int(gst_rates[i])
        gst_sum_dict[key][key_gst] = {}
        gst_sum_dict[key][key_gst]['Quantity'] = round(0,2)
        gst_sum_dict[key][key_gst]['Gross Amt'] = round(0,2)
        gst_sum_dict[key][key_gst]['Addl Chrgs'] = round(0,2)
        gst_sum_dict[key][key_gst]['Sales Disc'] = round(0,2)
        gst_sum_dict[key][key_gst]['Flat Disc'] = round(0,2)
        gst_sum_dict[key][key_gst]['Less Exch'] = round(0,2)
        gst_sum_dict[key][key_gst]['Gross Sales'] = round(0,2)
        gst_sum_dict[key][key_gst]['GST'] = round(0,2)
        gst_sum_dict[key][key_gst]['Net Sales'] = round(0,2)

    gst_sum_dict[key]['Total'] = {}
    gst_sum_dict[key]['Total']['Quantity'] = round(0,2)
    gst_sum_dict[key]['Total']['Gross Amt'] = round(0,2)
    gst_sum_dict[key]['Total']['Addl Chrgs'] = round(0,2)
    gst_sum_dict[key]['Total']['Sales Disc'] = round(0,2)
    gst_sum_dict[key]['Total']['Flat Disc'] = round(0,2)
    gst_sum_dict[key]['Total']['Less Exch'] = round(0,2)
    gst_sum_dict[key]['Total']['Gross Sales'] = round(0,2)
    gst_sum_dict[key]['Total']['GST'] = round(0,2)
    gst_sum_dict[key]['Total']['Net Sales'] = round(0,2)

detail_sales_sum_df.fillna(round(0,2),inplace=True)

for i, row in detail_sales_sum_df.iterrows():
    key = int(row['Payment Modes'])
    gst_rate = round(row['GST RATE'],2)
    qty = round(row['NET QTY'],2)
    gross_amt = round(row['GROSS AMT'],2)
    addl_chrgs = round(row['ADDITION CHRGS'],2)
    sales_disc = round(row['SALES DISCOUNT'],2)
    flat_disc = round(row['SALES DISCOUNT'],2)
    less_exch= round(row['LESS EXCHANGE'],2)
    gross_sales = round(row['GROSS SALES'],2)
    gst_tax = round(row['GST TAX'],2)
    net_sales = round(row['NET SALES'],2)

    gst_sum_dict[key][gst_rate]['Quantity'] += qty
    gst_sum_dict[key][gst_rate]['Gross Amt'] += gross_amt
    gst_sum_dict[key][gst_rate]['Addl Chrgs'] += addl_chrgs
    gst_sum_dict[key][gst_rate]['Sales Disc'] += sales_disc
    gst_sum_dict[key][gst_rate]['Flat Disc'] += flat_disc
    gst_sum_dict[key][gst_rate]['Less Exch'] += less_exch
    gst_sum_dict[key][gst_rate]['Gross Sales'] += gross_sales
    gst_sum_dict[key][gst_rate]['GST'] += gst_tax
    gst_sum_dict[key][gst_rate]['Net Sales'] += net_sales


print_gst_sum_dict = {str(key): value for key, value in gst_sum_dict.items()}

with open('gst_sum_dict.json','w') as f:
    json.dump(print_gst_sum_dict,f, indent=4)  

table_data = []         
pdf_table = []
table_style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.white),  # Merge all columns in the first row
    ('FONTSIZE', (0, 0), (-1,0),12),
    ('FONTSIZE', (0, 0), (-1,-1),8),
    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),  # First Column
    ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),  # Second Row font
    ('COLWIDTH', (0, 0), (-1, -1), 100),  # Set column width for all columns  
    ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Top border
    ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),  # Bottom border
    ('LINEBEFORE', (0, 0), (0, -1), 1, colors.black),  # Left border
    ('LINEAFTER', (-1, 0), (-1, -1), 1, colors.black),  # Right border
    ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
    ('SPAN', (0,0), (6,0)),
    # ('INNERGRID', (0, 0), (-1, 0), 0, colors.white),  # Remove grid lines
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Text alignment
])
######################### GROSS AMT YET tO BE CALCULATED ##################################################
######################### TOTAL YET TO BE CALCULATED ##############################
for key, value in gst_sum_dict.items():

    table_data = []
    if key != 0:
        payment_modes = [PAYMENT_MODES[key-1], ""*7]
        table_data.append(payment_modes)

        gst_header = ["GST RATE", "3%", "5%", "12%", "18%", "28%", "Total"]

        quantity_row = ["Quantity"]
        gross_amt_row = ["MRP"]
        addl_charges_row = ["Addl Charges"]
        sales_disc_row = ["Sales Discount"]
        flat_disc_row = ["Flat Discount"]
        less_exch_row = ["Less Exchange"]
        gross_sales_row = ["Gross Sales"]
        gst_row = ["GST"]
        net_sales_row = ["Net Sales"]

        for k, v in value.items():
            if v['Quantity'] != 0.0:
                quantity_row.append(int(v['Quantity']))
            else:
                quantity_row.append("")

            if v['Gross Amt'] != 0.0:
                gross_amt_row.append(round(v['Gross Amt'],2))
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

        table = Table(table_data)
        table.setStyle(table_style)
        pdf_table.append(table)
        pdf_table.append(Spacer(1,20))    
        pdf_table.append(CondPageBreak(100))  
pdf = SimpleDocTemplate("GST_payment_summary.pdf", pagesize=letter)
pdf.build(pdf_table)

# display(detail_sales_sum_df)
# for i, row in detail_sales_sum_df.iterrows():
#     key = row['BILL NO.']
#     addl_chrgs = row["ADDITION CHRGS"] 
#     exch = row['LESS EXCHANGE']
#     flat_disc = row['FLAT DISCOUNT']
#     sales_disc = row['SALES DISCOUNT']

#     if key in bill_no_dict:
#         try:
#             for i in range(len(PAYMENT_MODES)):
#                 if row[PAYMENT_MODES[i]] != '':
#                     bill_no_dict[key]["Payment Mode"].append(i+1)
#         except KeyError as e:
#             print(f"No Payments were made using {PAYMENT_MODES[i]}")
#             # Nobreak here because the order of columns and the order of list is not the same.
#             # So it is wise to check all Payment modes in the list
#         bill_no_dict[key]["GST Rate"].append(row['GST RATE'])
#         bill_no_dict[key]["Qty_GST_wise"].append(row['NET QTY'])
#         # bill_no_dict[key]["Gross_Amt_GST_wise"].append(row['GROSS AMT'])
#         if addl_chrgs != '':  
#             bill_no_dict[key]["Addl Chrgs"].append(addl_chrgs)
#         if  exch != '':
#             bill_no_dict[key]["Less Exch"].append(exch)
#         if flat_disc != '':
#             bill_no_dict[key]["Flat Disc"].append(flat_disc)
#         if sales_disc != '':
#             bill_no_dict[key]["Sales Disc"].append(sales_disc)

#         for i in range(len(GST_TAX_RATES)):
#             try:
#                 if row[GST_TAX_RATES[i]] != '':
#                     bill_no_dict[key]["Gross Sales"].append(i+1)
#             except KeyError as e:
#                 print(f"No {GST_TAX_RATES[i]} products were bought today :(")
#                 break         
#         # bill_no_dict[key]["GST"].append(row['GST TAX'])
#         # bill_no_dict[key]["Net Sales"].append(row['NET SALES'])
#     else:
#         bill_no_dict[key] = {}
#         bill_no_dict[key]["Payment Mode"] = []
#         bill_no_dict[key]["GST Rate"] = []
#         bill_no_dict[key]["Qty_GST_wise"] = []
#         # bill_no_dict[row['BILL NO.']]["Gross_Amt_GST_wise"] = []
#         bill_no_dict[key]["Addl Chrgs"] = []
#         bill_no_dict[key]["Less Exch"] = []
#         bill_no_dict[key]["Flat Disc"] = []
#         bill_no_dict[key]["Sales Disc"] = []
#         bill_no_dict[key]["Gross Sales"] = []
#         bill_no_dict[key]["GST"] = []
#         bill_no_dict[key]["Net Sales"] = []

#         for i in range(len(PAYMENT_MODES)):
#             try: 
#                 if row[PAYMENT_MODES[i]] != '':
#                     bill_no_dict[row['BILL NO.']]["Payment Mode"].append(i+1)
#             except KeyError as e:
#                 print(f"No Payments were made using {PAYMENT_MODES[i]}")
#                 # Nobreak here because the order of columns and the order of list is not the same.
#                 # So it is wise to check all Payment modes in the list
#         bill_no_dict[key]["GST Rate"].append(row['GST RATE'])
#         bill_no_dict[key]["Qty_GST_wise"].append(row['NET QTY'])
#         # bill_no_dict[row['BILL NO.']]["Gross_Amt_GST_wise"].append(row['GROSS AMT'])
#         if addl_chrgs != '':  
#             bill_no_dict[key]["Addl Chrgs"].append(addl_chrgs)
#         if  exch != '':
#             bill_no_dict[key]["Less Exch"].append(exch)
#         if flat_disc != '':
#             bill_no_dict[key]["Flat Disc"].append(flat_disc)
#         if sales_disc != '':
#             bill_no_dict[key]["Sales Disc"].append(sales_disc)
#         for i in range(len(GST_TAX_RATES)):
#             try:
#                 if row[GST_TAX_RATES[i]] != '':
#                     bill_no_dict[key]["Gross Sales"].append(i+1)
#             except KeyError as e:
#                 print(f"No {GST_TAX_RATES[i]} products were bought today :(")
#                 break         
#         # bill_no_dict[key]["GST"].append(row['GST TAX'])
#         # bill_no_dict[key]["Net Sales"].append(row['NET SALES'])