import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *
from pynput.keyboard import Key, Controller as KeyboardController
import time
import subprocess
import datetime as dt
keyboard = KeyboardController()

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows# Insert your python code here

# INSERT ANSWER FOR QUESTION 1 HERE

while True:
    filepath = select_file('Please select a file in the folder where the ProductData.xlsx, MasterInvoices.pdf, and Order files are located')
    directory = filepath.split('/')
    del directory[-1]
    directory = '/'.join(directory)
    directory = directory + '/'
    directory
    filenames = os.listdir(directory)
    
    if 'MasterInvoice.pdf' in filenames and 'ProductData.xlsx' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message('Error','You have not selected a file in the correct folder. Please try again, or make sure all of the needed files are in your selected folder.')
print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
df = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
for index, row in df.iterrows():
    unitnum = row['Unit number']
    description = row['Description']
    product_descriptions[unitnum] = description
display(product_descriptions)

product_prices = {}
df2 = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
for index, row in df2.iterrows():
     unitnum = row['Unit number']
     price = row['Price per unit']
     product_prices[unitnum] = price
product_prices

# INSERT ANSWER FOR QUESTION 3 HERE
# df3 = get_pdf('MasterInvoice.pdf')
# df3
# there is an error and it is that i have a $NaN amount in my invoice 815
def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory+filename)
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = (selected_sheet['C9'].value).strftime('%m/%d/%Y')

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = dt.date.today().strftime('%m/%d/%Y')
    total = product_prices[selected_sheet['B24'].value]*selected_sheet['C24'].value
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'
    
    price0 = product_prices[selected_sheet['B24'].value]
    order_data['Price0'] = f'${price0:,.2f}'
    order_data['Quantity0'] = str(selected_sheet['C24'].value)
    amount0= product_prices[selected_sheet['B24'].value]*selected_sheet['C24'].value
    order_data['Amount0'] = f'${amount0:,.2f}'

    if selected_sheet['C25'].value is not None:
        price1= product_prices[selected_sheet['B25'].value]
        order_data['Price1'] = f'${price1:,.2f}'
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        amount1= product_prices[selected_sheet['B25'].value]*selected_sheet['C25'].value
        order_data['Amount1'] = f'${amount1:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
        total = total + product_prices[selected_sheet['B25'].value]*selected_sheet['C25'].value

    if selected_sheet['C26'].value is not None:
        price2= product_prices[selected_sheet['B26'].value]
        order_data['Price2'] = f'${price2:,.2f}'
        order_data['Quantity2'] = str(selected_sheet['C26'].value)
        amount2= product_prices[selected_sheet['B26'].value]*selected_sheet['C26'].value
        order_data['Amount2'] = f'${amount2:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]
        total = total + product_prices[selected_sheet['B26'].value]*selected_sheet['C26'].value

    if selected_sheet['C27'].value is not None:
        price3= product_prices[selected_sheet['B27'].value]
        order_data['Price3'] = f'${price3:,.2f}'
        order_data['Quantity3'] = str(selected_sheet['C27'].value)
        amount3= product_prices[selected_sheet['B27'].value]*selected_sheet['C27'].value
        order_data['Amount3'] = f'${amount3:,.2f}'
        order_data['Description3'] = product_descriptions[selected_sheet['B27'].value]
        total = total + product_prices[selected_sheet['B27'].value]*selected_sheet['C27'].value

    if selected_sheet['C28'].value is not None:
        price4= product_prices[selected_sheet['B28'].value]
        order_data['Price4'] = f'${price4:,.2f}'
        order_data['Quantity4'] = str(selected_sheet['C28'].value)
        amount4= product_prices[selected_sheet['B28'].value]*selected_sheet['C28'].value
        order_data['Amount4'] = f'${amount4:,.2f}'
        order_data['Description4'] = product_descriptions[selected_sheet['B28'].value]
        total = total + product_prices[selected_sheet['B28'].value]*selected_sheet['C28'].value

    order_data['Total'] = f'${total:,.2f}'
    order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]
   
    return order_data, name, email, date
order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
# Enter your gmail user name and password
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    gmail_user = "aubmillar1416@gmail.com"

    # List of recipients

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.
     
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone:(801) 555-2372
    Email: sales@pythontechco.com
    """

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()
send_email(8515,'Thomas Holt','aubmillar1416@gmail.com','11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames
filenames = sorted(filenames)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num,get_order_data(invoice_num, file, directory)[0],directory)
        send_email(invoice_num,get_order_data(invoice_num, file, directory)[1], get_order_data(invoice_num, file, directory)[2], get_order_data(invoice_num, file, directory)[3], directory)
        invoice_num = invoice_num + 1  
