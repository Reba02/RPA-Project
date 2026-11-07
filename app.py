# Insert your python code here
# Insert your python code here
#Import packages
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *
#Q1
pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

while True:
    directory=single_input('Please select the directory where your filse are located')
    if directory[-1]!='\\':
        directory=f'{directory}\\'
    dir_contents=os.listdir(directory)
    if ('ProductData.xlsx' in dir_contents and 'MasterInvoice.pdf' in dir_contents and 'Order1.xlsx' in dir_contents and 'Order2.xlsx' in dir_contents and 'Order3.xlsx' in dir_contents and 'Order4.xlsx' in dir_contents):
        break
    else:
         show_message('error','Your directory does not contain the needed files, please try again')    
print(directory)
#Q2
product_descriptions={}
product_prices={}

df_proddescr=pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')
for index, row in df_proddescr.iterrows():
    product_descriptions[row['Unit number']]= row['Description']
df_prices=pd.read_excel(directory+ 'ProductData.xlsx', sheet_name='Prices')
for index, row in df_prices.iterrows():
    product_prices[row['Unit number']]= row['Price per unit']

#3    
df=get_pdf(directory+'MasterInvoice.pdf')
def get_order_data (invoice_num, filename, directory):
    template= openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet=template['OrderForm']
    name=str(selected_sheet['C14'].value)
    email=str(selected_sheet['C13'].value)
    date=str(selected_sheet['C9'].value).split()[0].split('-')
    month=date[1]
    day=date[2]
    year=date[0]
    date=(f'{month}/{day}/{year}')
    company=str(selected_sheet['C15'].value)
    address=str(selected_sheet['C16'].value)
    city=str(selected_sheet['C17'].value)
    state=str(selected_sheet['C18'].value)
    zipcode=str(selected_sheet['C19'].value)
    update_dict={}
    update_dict['InvoiceNum']=str(invoice_num)
    update_dict['Date']=date
    update_dict['CustomerName']=name
    update_dict['CustomerCompany']=company
    update_dict['CustomerAddress1']=address
    update_dict['CustomerAddress2']=f'{city},{state} {zipcode}'
    order_total=0.0
    row=24
    item=0
    while True: 
        unit_num=selected_sheet[f'B{row}'].value
        quantity=selected_sheet[f'C{row}'].value
        if unit_num is None:
            break
        description=product_descriptions[unit_num]
        price=product_prices[unit_num]
        total=price*quantity
        order_total=total+order_total
        update_dict[f'Description{item}']=str(description)
        update_dict[f'Price{item}']=f'${price:,.2f}'
        update_dict[f'Quantity{item}']=str(quantity)
        update_dict[f'Amount{item}']=f'${total:,.2f}'
        row=row+1
        item=item+1
    update_dict['Total']=f'${order_total:,.2f}'
    return update_dict, name, email, date

#Q4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+ f'Invoice {invoice_num}.pdf', order_data)
order_data, name, email, date= get_order_data(8515, 'Order1.xlsx', directory)
create_invoice(8515,order_data,directory)

#Q5
gmail_user=single_input('Please provide your Email Address')
gmail_password=open(select_file('Please provide your gmail password file')).read()
def send_email(invoice_num, name, email, date, directory):
    subject=f"Your Order from PythonTech Co. - Invoice{invoice_num}"

    body=f"""Dear {name},

Attatched is your invoice dated {date}.

Thank you for your business.

PythonTechCo.
10 Python Dr
Pleasant Grove, UT 846062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message=EmailMessage()
    message['From']=gmail_user

    message['To']= email

    message['Subject']=subject
    message.set_content(body)

    path_to_file=directory+ f'Invoice {invoice_num}.pdf'
    message=add_attachment(path_to_file,message)

    mail_server=smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user,gmail_password)

    mail_server.send_message(message)
    mail_server.quit()


#Q6
filenames=os.listdir(directory)
invoice_num=8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data,name,email,date=get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num=invoice_num+1