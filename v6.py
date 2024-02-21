import xlsxwriter
from bs4 import BeautifulSoup
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64

# Read the HTML file
with open("btx.htm", "r") as file:
    html_content = file.read()

# Parse the HTML
soup = BeautifulSoup(html_content, "html.parser")

# Find the first table
# table1 = soup.find("table")
table1 = soup.find("table")

# Find the second table with the id "gvEodEnqSumm"
table2 = soup.find("table", id="gvEodEnqSumm")

process_date_label = soup.find("span", id="Label1").text
process_date_value = soup.find("span", id="lblProcDate").text

# Print the extracted data
print(process_date_label + ":", process_date_value)

# Extract data from the first table
data1 = []
for row in table1.find_all("td", id="tdBG"):
    row_data = []
    for cell in row.find_all(["span"]):
        cell_text = cell.get_text(strip=True)
        if cell_text:  # Check if cell text is not empty
            row_data.append(cell_text)
    if row_data:  # Only append if row_data is not empty
        data1.append(row_data)

# Extract table headers from the second table
table2_headers = []
for th in table2.find_all("th"):
    header_text = th.get_text(strip=True)
    if header_text:  # Check if header text is not empty
        table2_headers.append(header_text)

# Extract data from the second table
data2 = []
for row in table2.find_all("tr"):
    row_data = []
    for cell in row.find_all(["td"]):
        cell_text = cell.get_text(strip=True)
        if cell_text:  # Check if cell text is not empty
            row_data.append(cell_text)
    if row_data:  # Only append if row_data is not empty
        data2.append(row_data)

# Combine data from table1 and table2
combined_data = data1 + [table2_headers] + data2 

# Create a new workbook and add a worksheet
workbook = xlsxwriter.Workbook('tableContents.xlsx')
worksheet = workbook.add_worksheet('Sheet1')

# Define colors in hexadecimal format
white_color_hex = '#FFFFFF'
grey_color_hex = '#f5f5f5'
deep_grey_color_hex = '#d7dae1'
blue_font_hex = '#0b14fe'

# Define bold format
bold_format = workbook.add_format({'bold': True})

# Write combined data to XLSX
for row_index, row_data in enumerate(combined_data):
    # Choose color based on row index
    if row_index == len(data1):  # Header row
        cell_format = workbook.add_format({'bg_color': deep_grey_color_hex, 'font_color': blue_font_hex})
    elif row_index % 2 == 0:
        cell_format = workbook.add_format({'bg_color': white_color_hex})  # White color
    else:
        cell_format = workbook.add_format({'bg_color': grey_color_hex})  # Grey color

    for col_index, cell_data in enumerate(row_data):
        if row_index == len(data1):  # Header row
            if cell_data == "Task ID":  # Adjust width for "Task Name" header
                worksheet.set_column(col_index, col_index, 25)  # Set width to 300px (approx.)
            elif cell_data == "Task Name":
                worksheet.set_column(col_index, col_index, 40)  # Set width for "Start Time" header
            elif cell_data == "Start Time":
                worksheet.set_column(col_index, col_index, 20)  # Set width for "Start Time" header
            elif cell_data == "Actual Start Time":
                worksheet.set_column(col_index, col_index, 20)  # Set width for "Actual Start Time" header
            elif cell_data == "Actual End Time":
                worksheet.set_column(col_index, col_index, 20)  # Set width for "Actual End Time" header
            elif cell_data == "Duration":
                worksheet.set_column(col_index, col_index, 20)  # Set width for "Duration" header
            elif cell_data == "Status":
                worksheet.set_column(col_index, col_index, 25)  # Set width for "Status" header
            worksheet.write(row_index, col_index, cell_data, bold_format)
        else:
             worksheet.write(row_index, col_index, cell_data, cell_format)

    
# Adjust column widths
for i, column in enumerate(zip(*combined_data)):
    max_length = max(len(str(cell)) for cell in column)
    worksheet.set_column(i, i, max_length +5)  # Add a little extra space

# Close the workbook
workbook.close()

# Read the Excel file into a pandas DataFrame with specifying columns
# df = pd.read_excel('tableContents.xlsx', header=None)  # Assuming the data doesn't have headers

# Load the Excel file
file_path = 'tableContents.xlsx'
try:
    df = pd.read_excel(file_path,header=None)
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found.")
    exit()
  
df = df.fillna('')

# Convert the DataFrame to an HTML table with no index
html_table = df.to_html(index=False, header=False)

# Modify the HTML table to make the header grey
soup = BeautifulSoup(html_table, "html.parser")
tr_elements = soup.find_all("tr")[1:]  # Select rows starting from the third row

for idx, element in enumerate(soup.find_all(["td"])):
    if element.name == "td":
        if element.text.strip() in ["Task ID", "Task Name", "Start Time", "Actual Start Time", "Actual End Time", "Duration", "Next Day", "Status"]:
            element['style'] = 'background-color:#d7dae1; color:#0b14fe; font-weight:bold;'
    # elif element.name == "tr":
    #     if idx % 2 == 0:
    #         element['style'] = 'background-color: #FFFFFF;'
    #     else:
    #         element['style'] = 'background-color: #f5f5f5;'

for tr_element in tr_elements:
    td_elements = tr_element.find_all("td")
    for td_element in td_elements:
        td_element['style'] = 'padding: 5px;'
        if td_element.text.strip() == "Process Succeeded!":
            td_element['style'] += 'background-color: green; color: white;'
        elif td_element.text.strip() == "Process Failed!":
            td_element['style'] += 'background-color: red; color: white;'

html_table = str(soup)

def get_base64_encoded_image(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')

body = '''
        <p style="color: #a6a698; font-size:13px; font-family: Arial, sans-serif;">Regards,</p>
        <br>
        <p style="color: black; font-size: 14px; font-family: Calibri, sans-serif; font-weight: bold; margin: 0; padding:0;margin-bottom: 3px">IT OPERATIONS</p>
        <p style="color: #a6a698; font-size:13px; font-family: Arial, sans-serif; margin: 0; padding:0;margin-bottom: 3px">Group Digital, Technology & Transformation</p>
        <p style="color: #a6a698; font-size: 13px; font-family: Arial, sans-serif; font-weight: bold; margin: 0; padding:0;margin-bottom: 3px">Kenanga Investment Bank Berhad</p>
        <p style="color: #a6a698; font-size: 12px; font-family: Arial, sans-serif; margin: 0; padding:0;margin-bottom: 3px">Level 6, Kenanga Tower</p>
        <p style="color: #a6a698; font-size: 12px; font-family: Arial, sans-serif;margin: 0; padding:0;margin-bottom: 4px">237, Jalan Tun Razak, 50400 Kuala Lumpur</p>
        <p style="color: #4472c4; font-size: 11px; font-family: Arial, sans-serif;margin: 0; padding:0;margin-bottom: 3px">Tel: GL +60 3 21722888 (Ext:8364 / 8365 / 8366 / 8357) </p>
        <br>
        <img src="data:image/png;base64, {}" alt="image1"> <!-- Embed image1 -->
        <br>
        <img src="data:image/png;base64, {}" alt="image2" > <!-- Embed image2 -->
            '''.format(get_base64_encoded_image("C:/Users/danny/OneDrive/Desktop/BTX Monitoring/image1.png"), get_base64_encoded_image("C:/Users/danny/OneDrive/Desktop/BTX Monitoring/image2.png"))


# html_content = f"<p>Process Date: {process_date_value}</p>\n" + html_table + body
html_content = f"<p>Process Date: {process_date_value}</p>\n\n{html_table}\n{body}"
# Set up the email details
sender_email = "danny-wong-02@hotmail.com"
password = 'tzepcyedvkultdtc'
receiver_email = ["whysodamn2012@gmail.com"]
# cc_emails = ["dannywong@kenanga.com.my"]


# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = sender_email
message["To"] =  ','.join(receiver_email)
# message['Cc'] = ','.join(cc_emails)
message["Subject"] = f"BTX Start Of Day process monitoring {process_date_value} - checking @ 4.30am"

# Add HTML table to the email body
message.attach(MIMEText(html_content, "html"))

# Connect to the SMTP server and send the email
with smtplib.SMTP("smtp-mail.outlook.com", 25) as server:
    server.starttls()
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message.as_string())
    print("Email successfully sent!")
