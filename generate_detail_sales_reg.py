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
    "CASH AMT.",           # 1
    "CREDIT CARDS",       # 2
    "HDFC",               # 3
    "BANDHAN",            # 4
    "GOOGLE PAY",         # 5
    "CHEQUE",             # 6
    "GV AMT.",             # 7 
    "OTHERS",             # 8
    "OUTSTANDING",        # 9
    "CN AMT. ADJ.",       # 10
]

sender_email = 'kushal@devaanx.com'
receiver_email = 'kushal@devaanx.com'
password = 'hazx wecz xooc okjw' 
filename = ["detail_sales_reg_mens_25_07_2024.pdf","detail_sales_reg_women_27_07_2024.pdf"] 
title = ['DEVAA ANNEX', 'DEVAA FOR WOMEN']

message = MIMEMultipart()
message['From'] = sender_email
message['To'] = receiver_email
message['Subject'] = '25/07/2024 Reports'

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

def custom_key(A, element):
    try:
        return A.index(element)
    except ValueError:
        return float('inf')  # If element not found in A, place it at the end
    
detail_sales_reg_df = pd.read_excel("detail_sales_reg.xlsx")
detail_sales_reg_df.drop(detail_sales_reg_df.tail(1).index, inplace=True)
display(detail_sales_reg_df)
payment_modes = detail_sales_reg_df.columns.to_list()[3:]
ordered_payment_modes = sorted(payment_modes, key=lambda x: custom_key(PAYMENT_MODES, x))
# detail_sales_reg.sort_values(cols, ascending=[True]*len(cols),inplace=True)
display(ordered_payment_modes)
detail_sales_reg_dict = {}

############## PRE INITIALIZE DICTIONARY #######################
for x in range(len(ordered_payment_modes)):
    detail_sales_reg_dict[ordered_payment_modes[x]] = {}
    detail_sales_reg_dict[ordered_payment_modes[x]]["Bill No."] = []
    detail_sales_reg_dict[ordered_payment_modes[x]]["Qty"] = []
    detail_sales_reg_dict[ordered_payment_modes[x]]["Amount"] = []



for i, row in detail_sales_reg_df.iterrows():
    bill_no = row["BILL NO."]
    qty = row["NET QTY"]
    for x in range(len(ordered_payment_modes)):
        




# table.setStyle(table_style)
# pdf_table = []
# pdf_table.append(table)
# pdf = SimpleDocTemplate(filename[i], pagesize=letter)
# frameT = Frame(x1=0,y1=0,width=8.3*inch, height=11.7*inch,topPadding=60, bottomPadding=0.25*inch,leftPadding=0,rightPadding=0)
# pdf.addPageTemplates([PageTemplate(id='First',frames=frameT, onPage=lambda canvas, doc: myFirstPage(canvas, doc, title[i], "25/07/2024", "27/07/2024"), pagesize=A4)])
# frameT = Frame(x1=0,y1=0,width=8.3*inch, height=11.7*inch,topPadding=0.5*inch, bottomPadding=0.25*inch,leftPadding=0,rightPadding=0)
# pdf.addPageTemplates([PageTemplate(id='Later',frames=frameT, onPage=myLaterPages,pagesize=A4)])
# pdf.build(pdf_table)

#     with open(filename[i], 'rb') as attachment:
#         part = MIMEBase('application', 'octet-stream')
#         part.set_payload(attachment.read())
#     encoders.encode_base64(part)
#     part.add_header('Content-Disposition', f'attachment; filename= {filename[i]}')

#     message.attach(part)

# # Add body to email
# body = "Detail Sales Summary Report and Payment mode wise GST Report 25/04/2024"
# message.attach(MIMEText(body, 'plain'))

# # Connect to SMTP server and send email
# with smtplib.SMTP('smtp.gmail.com', 587) as server:
#     server.starttls()
#     server.login(sender_email, password)
#     text = message.as_string()
#     server.sendmail(sender_email, receiver_email, text)
#     print("Email sent successfully!")
