from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import datetime
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import base64
from selenium.common.exceptions import NoSuchElementException

class taskTwo():

    driver = webdriver.Chrome()
    #Initial URL
    driver.get('https://btx.kenanga.com.my/btxadmin/default.aspx')

    # Maximizing window
    driver.maximize_window()

    #Wait for the website to fully load
    time.sleep(2)

    # Make selenium to automate the login process (username,password,login button)
    username_input = driver.find_element(By.ID, 'ctl00_cntPlcHldrContent_txtUsrID')
    password_input = driver.find_element(By.ID, 'ctl00_cntPlcHldrContent_txtUsrPwd')
    submit_button = driver.find_element(By.ID,'ctl00_cntPlcHldrContent_ibSignIn')

    time.sleep(5)

    username_input.send_keys('ITHQOPR')
    password_input.send_keys('Kibb8888')
    submit_button.click()

    time.sleep(2)
    # Click on the image button to navigate to another page
    # The button image name is 'Day End Maintenance'
    dayEndM_image = driver.find_element(By.XPATH,"//img[@src='/btxadmin/images/demo/icons/i_dayEndM_off.jpg']")
    dayEndM_image.click()

    # In here, after navigating to the new page, get the new url again
    # In this page, it should be let user to choose the 'Day End Enquiry'

    time.sleep(4)
    second_url = driver.current_url
    print("Second URL:", second_url)

    # Define Day End Enquiry XPaths
    day_end_enquiry_xpaths = [
        "//img[@src='/btxadmin/images/demo/icons/i_dayEndE_on.jpg']",
        "//img[@src='/btxadmin/images/demo/icons/i_dayEndE_off.jpg']"
    ]

    # Iterate through each XPath
    for xpath in day_end_enquiry_xpaths:
        try:
            # Try to find the element
            dayEndEnquiry_image = driver.find_element(By.XPATH,xpath)
            # If found, click on it
            dayEndEnquiry_image.click()
            # Exit loop if element is found and clicked
            break 
        except NoSuchElementException:
            # If element is not found, continue to the next XPath
            continue

    # In this page, it should be let user to choose the 'Day End Qnquiry' and 'Process Date'
    third_url = driver.current_url
    print("Third URL:", third_url)

    time.sleep(3)
    # Selecting the multi-select element by locating its id
    select = Select(driver.find_element(By.ID,"ctl00_cntPlcHldrContent_selEODEnquiry"))

    # Select the value from the dropdown 
    select.select_by_value("1,S")

    # Locate the datepicker input element
    datepicker_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_cntPlcHldrContent_txtDate")))

    # Get yesterday's date
    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    yesterday_str = yesterday.strftime('%d/%m/%Y')  

    # Enter yesterday's date into the input field
    datepicker_input.clear()  # Clear any existing value
    datepicker_input.send_keys(yesterday_str)

    time.sleep(3)

    searchButton = driver.find_element(By.ID, "ctl00_cntPlcHldrContent_btnTpltUpdate_btnSearch")
    searchButton.click()

    forth_url = driver.current_url
    print("Forth URL:", forth_url)

    # Switch to the new window
    new_window = driver.window_handles[1]
    driver.switch_to.window(new_window)

    # Maximize the window
    driver.maximize_window()

    fifth_url = driver.current_url
    print("Fifth URL:", fifth_url)

    process_dates = driver.find_elements(By.XPATH, "//table[@class='clsTable']/tbody/tr[2]/td[@id='tdBG']/span")

    for process in process_dates:
        print(process.text)

    driver.get(fifth_url)

    # Get the HTML content of the fifth URL
    html_content_fifth_url = driver.page_source

    # Print the HTML content of the fifth URL
    # print("HTML Content of Fifth URL:", html_content_fifth_url)

    # Parse the HTML
    soup = BeautifulSoup(html_content_fifth_url, "html.parser")

    # Find the first table
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
    # Select rows starting from the third row
    tr_elements = soup.find_all("tr")[1:]  

    for idx, element in enumerate(soup.find_all(["td"])):
        if element.name == "td":
            if element.text.strip() in ["Task ID", "Task Name", "Start Time", "Actual Start Time", "Actual End Time", "Duration", "Next Day", "Status"]:
                element['style'] = 'background-color:#d7dae1; color:#0b14fe; font-weight:bold;'
    

    for tr_element in tr_elements:
        td_elements = tr_element.find_all("td")
        for td_element in td_elements:
            td_element['style'] = 'padding: 5px;'
            # Check if the text content matches "Process Succeeded!"
            if "Process Succeeded!" in td_element.text.strip():
                td_element['style'] += 'background-color: green; color: white;'
            # Check if the text content matches "Process Failed!"
            elif "Process Failed!" in td_element.text.strip():
                td_element['style'] += 'background-color: red; color: white;'


    html_table = str(soup)
    # print(html_table)

    # tasks = driver.find_elements(By.XPATH,"//table[@id='gvEodEnqSumm']/tbody/tr/td[1]")
    # # tasks = driver.find_elements(By.XPATH,"//table[@id='gvEodEnqSumm']/tbody/tr")

    # for task in tasks:
    #     print(task.text)
    # Concatenate process date with HTML table

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
                '''.format(get_base64_encoded_image("C:/Users/dannywong/Desktop/BTX Monitoring/image1.png"), get_base64_encoded_image("C:/Users/dannywong/Desktop/BTX Monitoring/image2.png"))


    # html_content = f"<p>Process Date: {process_date_value}</p>\n" + html_table + body
    html_content = f"<p>Process Date: {process_date_value}</p>\n\n{html_table}\n{body}"
    # Set up the email details
    sender_email = "dannywong@kenanga.com.my"
    receiver_email = ["dannywong@kenanga.com.my"]
    cc_emails = ["dannywong@kenanga.com.my"]


    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] =  ','.join(receiver_email)
    message['Cc'] = ','.join(cc_emails)
    message["Subject"] = f"BTX Start Of Day process monitoring {process_date_value} - checking @ 5.45am"

    # Add HTML table to the email body
    message.attach(MIMEText(html_content, "html"))

    # Connect to the SMTP server and send the email
    with smtplib.SMTP("172.21.5.60", 25) as server:
        # server.starttls()
        # server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print("Email successfully sent!")

    time.sleep(10)
    driver.quit()
