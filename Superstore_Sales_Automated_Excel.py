import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Load the dataset
file_path = 'superstore_sales.xlsx'
df = pd.read_excel(file_path)

# Display basic info about the dataset
print(df.info())

# Clean the data
# Handling missing values (drop or fill)
df = df.dropna(subset=['Sales', 'Profit'])  # Drop rows with missing sales or profit

# Ensure 'Order Date' is in datetime format
df['Order Date'] = pd.to_datetime(df['Order Date'])

# Create a new column for 'Year' and 'Month' for easy grouping
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month_name()

# Group by Year and Month for sales and profit aggregation
monthly_sales = df.groupby(['Year', 'Month'])[['Sales', 'Profit']].sum().reset_index()

# Create a new Excel workbook and worksheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Sales Dashboard'

# Add the header row
headers = ['Year', 'Month', 'Total Sales', 'Total Profit']
ws.append(headers)

# Add the aggregated data into the Excel sheet
for index, row in monthly_sales.iterrows():
    ws.append(row.tolist())

# Create a Line Chart for Sales over time
chart_sales = LineChart()
data = Reference(ws, min_col=3, min_row=1, max_col=3, max_row=len(monthly_sales)+1)
chart_sales.add_data(data, titles_from_data=True)
chart_sales.title = "Sales Over Time"
ws.add_chart(chart_sales, "F5")

# Create a Line Chart for Profit over time
chart_profit = LineChart()
data = Reference(ws, min_col=4, min_row=1, max_col=4, max_row=len(monthly_sales)+1)
chart_profit.add_data(data, titles_from_data=True)
chart_profit.title = "Profit Over Time"
ws.add_chart(chart_profit, "F20")

# Save the workbook
wb.save('sales_dashboard.xlsx')

# Step 4: Automate Sending the Dashboard via Email
def send_email_with_attachment():
    from_addr = 'mogadpally.a@northeastern.edu'  # Replace with your email
    to_addr = 'mogadpallya@gmail.com'  # Replace with the recipient's email
    subject = "Automated Sales Dashboard"
    body = "Please find the latest sales dashboard attached."

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach the Excel file
    attachment = open('sales_dashboard.xlsx', 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={attachment.name}')
    msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_addr, 'Aditya@123')  # Replace with your email password
        text = msg.as_string()
        server.sendmail(from_addr, to_addr, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error: {str(e)}")

# Call the function to send the email
send_email_with_attachment()
