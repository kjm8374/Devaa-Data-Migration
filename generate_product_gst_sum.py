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


product_gst_sum_df = pd.read_excel("Product_Gst_Sum.xlsx", skiprows=2)
product_gst_sum_df.drop(columns=['SNO.', 'CGST%', 'TOTAL CGST AMOUNT', 'SGST/IGST%', 'TOTAL SGST AMOUNT', 'TOTAL IGST AMOUNT'], inplace=True)
product_gst_sum_df.drop(product_gst_sum_df.index[-1], inplace=True) # drop last row which is just Grand Total

product_gst_sum_df.iloc[:,1:] = product_gst_sum_df.iloc[:,1:].astype("float64").round(2)

product_gst_sum_df['PRODUCT'] = product_gst_sum_df['PRODUCT'].astype(str)
product_gst_sum_df['GST RATE'] = product_gst_sum_df['GST RATE'].astype(int)
product_gst_sum_df['SALE QTY'] = product_gst_sum_df['SALE QTY'].astype(int)

#################### CHECK TO SEE IF GST RATE IS 0 ######################

if 0 in product_gst_sum_df['GST RATE']:
    print('Dropping row with GST RATE = 0')
    product_gst_sum_df.drop(product_gst_sum_df[product_gst_sum_df['GST RATE'] == 0].index, inplace=True)


product_gst_sum_dict = {}
products = sorted(product_gst_sum_df['PRODUCT'].unique())
# display(products)
######################## Pre Initialize the dictionary ###########################
for x in range(len(products)):

    # print(payment_modes[x].dtype)
    key = products[x]

    product_gst_sum_dict[products[x]] = {}
    for i in range(len(RATES)):
        key_gst = int(RATES[i])
        product_gst_sum_dict[key][key_gst] = {}
        product_gst_sum_dict[key][key_gst]['Quantity'] = 0.0
        product_gst_sum_dict[key][key_gst]['Gross Mrp'] = 0.0
        product_gst_sum_dict[key][key_gst]['Addl Chrgs'] = 0.0
        product_gst_sum_dict[key][key_gst]['Sales Disc'] = 0.0
        product_gst_sum_dict[key][key_gst]['Flat Disc'] = 0.0
        product_gst_sum_dict[key][key_gst]['Less Exch'] = 0.0
        product_gst_sum_dict[key][key_gst]['Gross Sales'] = 0.0
        product_gst_sum_dict[key][key_gst]['GST'] = 0.0
        product_gst_sum_dict[key][key_gst]['Net Sales'] = 0.0

    product_gst_sum_dict[key]['Total'] = {}
    product_gst_sum_dict[key]['Total']['Quantity'] = 0.0
    product_gst_sum_dict[key]['Total']['Gross Mrp'] = 0.0
    product_gst_sum_dict[key]['Total']['Addl Chrgs'] = 0.0
    product_gst_sum_dict[key]['Total']['Sales Disc'] = 0.0
    product_gst_sum_dict[key]['Total']['Flat Disc'] = 0.0
    product_gst_sum_dict[key]['Total']['Less Exch'] = 0.0
    product_gst_sum_dict[key]['Total']['Gross Sales'] = 0.0
    product_gst_sum_dict[key]['Total']['GST'] = 0.0
    product_gst_sum_dict[key]['Total']['Net Sales'] = 0.0
######################## Initialize Grand Total #############################
product_gst_sum_dict['Grand Total'] = {}
for i in range(len(RATES)):
        key_gst = int(RATES[i])
        product_gst_sum_dict['Grand Total'][key_gst] = {}
        product_gst_sum_dict['Grand Total'][key_gst]['Quantity'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Gross Mrp'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Addl Chrgs'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Sales Disc'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Flat Disc'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Less Exch'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Gross Sales'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['GST'] = 0.0
        product_gst_sum_dict['Grand Total'][key_gst]['Net Sales'] = 0.0

product_gst_sum_dict['Grand Total']['Total'] = {}
product_gst_sum_dict['Grand Total']['Total']['Quantity'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Gross Mrp'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Addl Chrgs'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Sales Disc'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Flat Disc'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Less Exch'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Gross Sales'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['GST'] = 0.0
product_gst_sum_dict['Grand Total']['Total']['Net Sales'] = 0.0

product_gst_sum_df.fillna(round(0,2),inplace=True)

with open('product_gst_sum_dict.json','w') as f:
    json.dump(product_gst_sum_dict,f, indent=4)  

for i, row in product_gst_sum_df.iterrows():
    key = row['PRODUCT']
    gst_rate = row['GST RATE']
    qty = row['SALE QTY']
    gross_mrp = row['GROSS MRP AMT']
    addl_chrgs = row['OTHER CHARGES']
    sales_disc = row['DISCOUNT VALUE']
    flat_disc = row['FLAT DISCOUNT']
    less_exch= row['LESS/EXCHANGE']
    gross_sales = row['GROSS SALES AMT']
    gst_tax = row['TOTAL GST AMOUNT']
    net_sales = row['NET SALES AMT']

    product_gst_sum_dict[key][gst_rate]['Quantity'] += qty
    product_gst_sum_dict[key][gst_rate]['Gross Mrp'] += gross_mrp
    product_gst_sum_dict[key][gst_rate]['Addl Chrgs'] += addl_chrgs
    product_gst_sum_dict[key][gst_rate]['Sales Disc'] += sales_disc
    product_gst_sum_dict[key][gst_rate]['Flat Disc'] += flat_disc
    product_gst_sum_dict[key][gst_rate]['Less Exch'] += less_exch
    product_gst_sum_dict[key][gst_rate]['Gross Sales'] += gross_sales
    product_gst_sum_dict[key][gst_rate]['GST'] += gst_tax
    product_gst_sum_dict[key][gst_rate]['Net Sales'] += net_sales    

# ########################## CALCULATE TOTAL #######################
for key, value in product_gst_sum_dict.items():

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

    for key, value in product_gst_sum_dict.items():
        qty += value[gst_rate]['Quantity']
        gross_mrp += value[gst_rate]['Gross Mrp']
        addl_chrgs += value[gst_rate]['Addl Chrgs']
        sales_disc += value[gst_rate]['Sales Disc']
        flat_disc += value[gst_rate]['Flat Disc']
        less_exch += value[gst_rate]['Less Exch']
        gross_sales += value[gst_rate]['Gross Sales']
        gst_tax += value[gst_rate]['GST']
        net_sales += value[gst_rate]['Net Sales']

    product_gst_sum_dict['Grand Total'][gst_rate] = {}
    product_gst_sum_dict['Grand Total'][gst_rate]['Quantity'] = qty
    product_gst_sum_dict['Grand Total'][gst_rate]['Gross Mrp'] = gross_mrp
    product_gst_sum_dict['Grand Total'][gst_rate]['Addl Chrgs'] = addl_chrgs
    product_gst_sum_dict['Grand Total'][gst_rate]['Sales Disc'] = sales_disc
    product_gst_sum_dict['Grand Total'][gst_rate]['Flat Disc'] = flat_disc
    product_gst_sum_dict['Grand Total'][gst_rate]['Less Exch'] = less_exch
    product_gst_sum_dict['Grand Total'][gst_rate]['Gross Sales'] = gross_sales
    product_gst_sum_dict['Grand Total'][gst_rate]['GST'] = gst_tax
    product_gst_sum_dict['Grand Total'][gst_rate]['Net Sales'] = net_sales

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

for key, v in product_gst_sum_dict['Grand Total'].items():
    qty += v['Quantity']
    gross_mrp += v['Gross Mrp']
    addl_chrgs += v['Addl Chrgs']
    sales_disc += v['Sales Disc']
    flat_disc += v['Flat Disc']
    less_exch += v['Less Exch']
    gross_sales += v['Gross Sales']
    gst_tax += v['GST']
    net_sales += v['Net Sales']

product_gst_sum_dict['Grand Total']['Total']['Quantity'] = int(qty)
product_gst_sum_dict['Grand Total']['Total']['Gross Mrp'] = gross_mrp
product_gst_sum_dict['Grand Total']['Total']['Addl Chrgs'] = addl_chrgs
product_gst_sum_dict['Grand Total']['Total']['Sales Disc'] = sales_disc
product_gst_sum_dict['Grand Total']['Total']['Flat Disc'] = flat_disc
product_gst_sum_dict['Grand Total']['Total']['Less Exch'] = less_exch
product_gst_sum_dict['Grand Total']['Total']['Gross Sales'] = gross_sales
product_gst_sum_dict['Grand Total']['Total']['GST'] = gst_tax
product_gst_sum_dict['Grand Total']['Total']['Net Sales'] = net_sales

        
with open('product_gst_sum_dict.json','w') as f:
    json.dump(product_gst_sum_dict,f, indent=4)  

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

for key, value in product_gst_sum_dict.items():

    table_data = []
    products = [key, ""*7]
    table_data.append(products)

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
    # if key == list(product_gst_sum_dict.keys())[-2]:
    #     print("Grand Total")
    #     pdf_table.append(PageBreak())
    # else:
    pdf_table.append(Spacer(1,25))    
    pdf_table.append(CondPageBreak(150))  

pdf = SimpleDocTemplate("Product_GST_summary_women.pdf", pagesize=letter)

frameT = Frame(x1=0,y1=0,width=8.3*inch, height=11.7*inch,topPadding=60, bottomPadding=0.25*inch,leftPadding=0,rightPadding=0)
pdf.addPageTemplates([PageTemplate(id='First',frames=frameT, onPage=lambda canvas, doc: myFirstPage(canvas, doc, "Devaa for Women", "25/07/2024", "27/07/2024"), pagesize=A4)])
frameT = Frame(x1=0,y1=0,width=8.3*inch, height=11.7*inch,topPadding=0.5*inch, bottomPadding=0.25*inch,leftPadding=0,rightPadding=0)
pdf.addPageTemplates([PageTemplate(id='Later',frames=frameT, onPage=myLaterPages,pagesize=A4)])

pdf.build(pdf_table)
