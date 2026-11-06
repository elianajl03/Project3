# Insert your python code here

#QUESTION 1
while True:
    file_path = select_file("Please select a file in the folder where your documents are located")
    file_path

    directory = file_path.split('/')
    del directory[-1]
    directory = '/'.join(directory)
    directory = directory + '/'

    filenames = os.listdir(directory)    

    if "ProductData.xlsx" in filenames and "MasterInvoice.pdf" in filenames and "Order1.xlsx" in filenames and "Order2.xlsx" in filenames and "Order3.xlsx" in filenames and "Order4.xlsx" in filenames:
        print(directory)
        break
    else:
        show_message('Error', 'you much select a folder containing the files ProductData.xlsx, MasterInvoice.pdf, and all your orders')

print(filenames)

#QUESTION 2
product_descriptions = {}
product_prices = {}


df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = "Descriptions")
for index,row in df1.iterrows():
    unit = row['Unit number']
    desc = row['Description']
    product_descriptions[unit] = desc
display(product_descriptions)


df2 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = "Prices")
for index,row in df2.iterrows():
    unit = row['Unit number']
    price = row['Price per unit']
    product_prices[unit] = price
display(product_prices)


#QUESTION 3
def get_order_data (invoice_num, filename, directory):

# 1. A dictionary containing the order data. Use the dictionary format required to fill the `MasterInvoice.pdf` file.
# df = get_pdf('MasterInvoice.pdf')
# df
# 2. The name of the customer.
    template = openpyxl.load_workbook(directory+filename)
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
# 3. The email of the customer.
    email = selected_sheet['C13'].value
# 4. The date of the order. make it a string 
    date = selected_sheet['C9'].value
    date = date.strftime('%m/%d/%Y')


    order_data = {}

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = dt.datetime.today().strftime('%m/%d/%Y')
    total = product_prices[selected_sheet['B24'].value]*selected_sheet['C24'].value
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value 
    order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'

    price0 = product_prices[selected_sheet['B24'].value]
    order_data['Price0'] = f'${price0:,.2f}'
    order_data['Quantity0'] = str(selected_sheet['C24'].value)
    amount0 = product_prices[selected_sheet['B24'].value]*selected_sheet['C24'].value
    order_data['Amount0'] = f'${amount0:,.2f}'

    if selected_sheet['C25'].value is not None: 
        price1 = product_prices[selected_sheet['B25'].value]
        order_data['Price1'] = f'${price1:,.2f}'
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        amount1 = product_prices[selected_sheet['B25'].value]*selected_sheet['C25'].value
        order_data['Amount1'] = f'${amount1:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
        total = total + product_prices[selected_sheet['B25'].value]*selected_sheet['C25'].value

    if selected_sheet['C26'].value is not None: 
        price2 = product_prices[selected_sheet['B26'].value]
        order_data['Price2'] = f'${price2:,.2f}'
        order_data['Quantity2'] = str(selected_sheet['C26'].value)
        amount2 = product_prices[selected_sheet['B26'].value]*selected_sheet['C26'].value
        order_data['Amount2'] = f'${amount2:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value] 
        total = total + product_prices[selected_sheet['B26'].value]*selected_sheet['C26'].value

    if selected_sheet['C27'].value is not None: 
        price3 = product_prices[selected_sheet['B27'].value]
        order_data['Price3'] = f'${price3:,.2f}'
        order_data['Quantity3'] = str(selected_sheet['C27'].value)
        amount3 = product_prices[selected_sheet['B27'].value]*selected_sheet['C27'].value
        order_data['Amount3'] = f'${amount3:,.2f}'
        order_data['Description3'] = product_descriptions[selected_sheet['B27'].value] 
        total = total + product_prices[selected_sheet['B27'].value]*selected_sheet['C27'].value

    if selected_sheet['C28'].value is not None: 
        price4 = product_prices[selected_sheet['B28'].value]
        order_data['Price4'] = f'${price4:,.2f}'
        order_data['Quantity4'] = str(selected_sheet['C28'].value)
        amount4 = product_prices[selected_sheet['B28'].value]*selected_sheet['C28'].value
        order_data['Amount4'] = f'${amount4:,.2f}'
        order_data['Description4'] = product_descriptions[selected_sheet['B28'].value] 
        total = total + product_prices[selected_sheet['B28'].value]*selected_sheet['C28'].value
   
    order_data['Total'] = f'${total:,.2f}' 


    order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]


    return order_data, name, email, date


#QUESTION 4
def create_invoice (invoice_num, order_data, directory):
    create_pdf("MasterInvoice.pdf", f"Invoice {invoice_num}.pdf", order_data)


#QUESTION 5
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email (invoice_num, name, email, date, directory):
   
    # Enter your gmail user name and password

    gmail_user = "eliana.laclare@gmail.com"

    # List of recipients
    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email
    #tripple quote allow you to span multiple lines of code so you type how you would on email
    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment
    path_to_file = directory + f"Invoice {invoice_num}.pdf" # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

    #WHAT DO I RETURN FOR THE FUNCTION 
    

#QUESTION 6
filenames = os.listdir(directory)
filenames = sorted(filenames)

invoice = 8515
for file in filenames:
    if "Order" and "xlsx" in file:
        get_order_data(invoice, file, directory)
        create_invoice(invoice, get_order_data(invoice,file, directory)[0], directory)
        send_email(invoice, get_order_data(invoice,file, directory)[1], get_order_data(invoice,file, directory)[2], get_order_data(invoice,file, directory)[3], directory)
        invoice = invoice + 1 


#QUESTION 7