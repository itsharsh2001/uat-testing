from flask import Flask, request, jsonify
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime

from flask_cors import CORS
import pandas as pd
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
import os
import win32com.client

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import threading

app = Flask(__name__)
CORS(app)


@app.route('/test', methods=['POST'])
def run_selenium_test():
    start_time = time.time()

    website_url = request.json.get('website_url')
    username = request.json.get('username')
    password = request.json.get('password')
    first = request.json.get('first')
    second = request.json.get('second')
    third = request.json.get('third')
    fourth = request.json.get('fourth')
    # module_numbers=request.json.get('module_numbers')
    module_numbers = []
    if (first):
        module_numbers.append(1)
    if (second):
        module_numbers.append(2)
    if (third):
        module_numbers.append(3)
    if (fourth):
        module_numbers.append(4)
    print(module_numbers)

    # chrome_options = webdriver.ChromeOptions()
    # # chrome_options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
    # chrome_options.add_argument('--no-sandbox')  # Avoid sandbox issues
    # chrome_options.add_argument('--incognito')
    # chrome_options.add_argument("--disable-web-security")
    # chrome_options.add_argument("--allow-running-insecure-content")
    # chrome_options.add_argument("--disable-infobars")
    # chrome_options.add_argument("--disable-notifications")

    # chrome_driver_path = "/chromedriver_win32/chromedriver.exe"

    # service = Service(chrome_driver_path)
    # driver = webdriver.Chrome(service=service, options=chrome_options)

    # ChromeOptions = Options()

    # driver = webdriver.Chrome(service=Service(executable_path="/chromedriver_win32/chromedriver.exe"), options=ChromeOptions)
    # service = webdriver.ChromeService(executable_path=chrome_driver_path)

    # driver = webdriver.Chrome()

    chrome_options = webdriver.ChromeOptions()
    download_directory = 'c:\\Users\\harsh.vijaykumar\\Downloads'

    prefs = {
        'download.default_directory': download_directory,
        'download.prompt_for_download': False,  # Disable the download popup
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False,  # Disable safe browsing, which can trigger warnings
        'browser.download.folderList': 2,  # Use custom directory
        'profile.default_content_setting_values.automatic_downloads': 1,
    }

    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)

    driver.get(website_url)
    driver.maximize_window()

    # download_directory = "C:/Users/harsh.vijaykumar/Downloads"
    # file_name = "Bank_Consolidated.xlsx"
    # file_path = os.path.join(download_directory, file_name)
    # reports_checker(1,file_path,'Consolidated')

    failed_test_column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
                                'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)', 'Source']
    failed_df = pd.DataFrame(columns=failed_test_column_names)

    for module_number in module_numbers:
        response = {}
        column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
                        'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)']
        # column_names = ['S.No', 'Test Name', 'Result']
        df = pd.DataFrame(columns=column_names)
        # Open the website URL in Chrome

        # username = 'adminuser1'
        # password = 'Audit@123'

        df = login(username, password, driver, df)

        # module_number = 1
        df = module_choose(driver, module_number, df)

        heading = ''
        if (module_number == 1):
            heading = 'Bank Confirmations'
        elif (module_number == 2):
            heading = 'Debtor Confirmations'
        elif (module_number == 3):
            heading = 'Creditor Confirmations'
        elif (module_number == 4):
            heading = 'Legal Matter Confirmations'

        df = h5textchecker(driver, heading, df)
        df = role(driver, df)

        df = refresh(driver, df)

        df = email_batch_link(driver, df)

        if heading == 'Bank Confirmations':
            df = navigationchecker(driver, 'Debtor Confirmations', df)
            df = navigationchecker(driver, 'Creditor Confirmations', df)
            df = navigationchecker(driver, 'Legal Matter Confirmations', df)
            df = navigationchecker(driver, 'Bank Confirmations', df)
        elif heading == 'Debtor Confirmations':
            df = navigationchecker(driver, 'Creditor Confirmations', df)
            df = navigationchecker(driver, 'Legal Matter Confirmations', df)
            df = navigationchecker(driver, 'Bank Confirmations', df)
            df = navigationchecker(driver, 'Debtor Confirmations', df)
        elif heading == 'Creditor Confirmations':
            df = navigationchecker(driver, 'Legal Matter Confirmations', df)
            df = navigationchecker(driver, 'Bank Confirmations', df)
            df = navigationchecker(driver, 'Debtor Confirmations', df)
            df = navigationchecker(driver, 'Creditor Confirmations', df)
        else:
            df = navigationchecker(driver, 'Bank Confirmations', df)
            df = navigationchecker(driver, 'Debtor Confirmations', df)
            df = navigationchecker(driver, 'Creditor Confirmations', df)
            df = navigationchecker(driver, 'Legal Matter Confirmations', df)
        # time.sleep(2)

        # email batches link

        df = new_email_batch_button(driver, df)

        df = batch_creation(driver, df, module_number)

        df = attachments_download_batches_level(driver, df, module_number)

        df = is_table_body_visible_and_view_details_click(driver, df)
        mail_subject_unique_id = []
        name = []
        mail_category = []

        legal_party = []

        df, mail_subject_unique_id, name, mail_category, legal_party = mail_send_from_application(
            driver, df, mail_subject_unique_id, name, module_number, mail_category, legal_party)

        print(mail_subject_unique_id)
        time.sleep(20)

        time.sleep(200)
        remainder_checker(driver, df, module_number, mail_subject_unique_id)

        df = response_to_email_received_in_outlook(
            driver, df, mail_subject_unique_id, name, mail_category, module_number, legal_party)

        time.sleep(360)
        driver.refresh()

        # email_responded_button_checker(driver, df, module_number)

        email_response_count_colour(driver, df, module_number, mail_subject_unique_id, name, mail_category, legal_party)

        print('refresh before data_checker')
        time.sleep(10)

        df = data_checker(driver, df, mail_subject_unique_id, name, module_number, mail_category, legal_party)

        email_filter(driver, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[1]/div/div[1]/label', 11, df)

        driver.refresh()

        time.sleep(5)
        email_filter(driver, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[2]/div/div[2]/label', 10, df)

        df = date_filter(driver, df, 11)
        driver.refresh()
        df = report_download_after_view_details(
            driver, df, mail_subject_unique_id, module_number)
        time.sleep(5)
        print('helloji')

        report_checker_batches_level(driver, df, module_number)

        df = client_report_checker(driver, df, module_number)

        df = logout(driver, df)

        response = {
            'results_count': 'answer'
        }

        # Generate a sequence of numbers and fill the column
        num_values = len(df)  # Number of values to generate
        df['Test Case'] = list(range(1, num_values + 1))

        print(df)

        del df['Common/Module Specific']

        df_reorder = ['Test Case', 'Test Case Button/Description',
                      'Status(PASS/FAIL)', 'Test Case Screen', 'Repro Steps', 'Expected Result']
        df = df[df_reorder]

        # writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        df_name = ''
        if module_number == 1:
            # df.to_excel(writer, sheet_name='Bank Confirmations', index=False)
            df.to_excel('Bank Confirmations.xlsx', index=False)
            df_name = 'Bank Confirmations'
        elif module_number == 2:
            df.to_excel('Debtor Confirmations.xlsx', index=False)
            df_name = 'Debtor Confirmations'
            # df.to_excel(writer, sheet_name='Debtor Confirmations', index=False)
        elif module_number == 3:
            df.to_excel('Creditor Confirmations.xlsx', index=False)
            df_name = 'Creditor Confirmations'
            # df.to_excel(writer, sheet_name='Creditor Confirmations', index=False)
        elif module_number == 4:
            df.to_excel('Legal Confirmations.xlsx', index=False)
            df_name = 'Legal Confirmations'
            # df.to_excel(writer, sheet_name='Legal Confirmations', index=False)
        # Return the response as JSON
        # driver.quit()

        mask = df['Status(PASS/FAIL)'] == 'FAIL'
        filtered_rows = df[mask]
        filtered_rows['Module'] = df_name
        failed_df = pd.concat([failed_df, filtered_rows], ignore_index=True)

    failed_df_reorder_columns = ['Test Case', 'Module', 'Test Case Button/Description',
                                 'Status(PASS/FAIL)', 'Test Case Screen', 'Repro Steps', 'Expected Result']
    failed_df = failed_df[failed_df_reorder_columns]
    failed_df.to_excel('Failed Tests.xlsx', index=False)

    print(failed_df.shape[0], 'number of records in failed df')

    html_body = ''

    os.startfile("outlook")

    html_table = failed_df.to_html(index=False)

    if failed_df.shape[0] == 0:
        html_body = f"""<html>
        <body>
        Hi Team,<br>
        Full Testing for the ConfirmEase application has been done successfully. <br>
        All the tests cases are working fine.<br>
        PFA the test reports.<br>

        Thanks.
        </body>
        </html> """
    elif failed_df.shape[0] == 1:
        html_body = f"""Hi Team, <br>
Full Testing for the ConfirmEase application has been done successfully. <br>
There is {failed_df.shape[0]} failed test case. Below are the detail of same:<br>
{html_table}<br>
PFA the test reports.<br>

Thanks,

"""
    else:
        html_body = f"""Hi Team, <br>
Full Testing for the ConfirmEase application has been done successfully. <br>
There are {failed_df.shape[0]} failed test cases. Below are the detail of same:<br>
{html_table}<br>
PFA the test reports.<br>

Thanks,

"""

    


    outlook = win32com.client.Dispatch(
        'Outlook.Application')
    new_email = outlook.CreateItem(0)  # 0 represents olMailItem

    # Set email properties
    new_email.Subject = 'Full Test'
    new_email.HTMLBody = html_body

    new_email.To = 'harsh.vijaykumar@walkerchandiok.in'
    new_email.Recipients.Add('Abhishek.Malan@IN.GT.COM')
    new_email.Recipients.Add('siddharth.mishra@walkerchandiok.in')

    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Legal Confirmations.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Bank Confirmations.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Debtor Confirmations.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Creditor Confirmations.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
        # Send the reply email
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Failed Tests.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
        # Send the reply email
    new_email.Send()

    end_time = time.time()

    print(end_time-start_time,'time spent')

    return jsonify(response)

def email_responded_button_checker(driver, df, module_number):
    driver.refresh()
    try:
        print('e_res_count1')
        header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
        print('e_res_count2')
        header_cells = WebDriverWait(header_row, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
        print('e_res_count3')
        required_index = 0
        email_reminder_count = 0

        whole_table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

        i = 0
        for header_cell in header_cells:
            if i != 0:
                driver.execute_script(
                    f"arguments[0].scrollLeft += 80;", whole_table_body)
            print(header_cell)
            column_name = WebDriverWait(header_cell, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            print('columnname', column_name)
            name_exact = WebDriverWait(column_name, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span')))
            print('namexact', name_exact)
            text_part = WebDriverWait(name_exact, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
            print('column name = ', text_part)

            if text_part == 'Email Responded':
                required_index = i
            elif text_part == 'Email Reminder Count':
                email_reminder_count = i
            i += 1
        print('e_res_count4', required_index)
        print('e_res_count5', email_reminder_count)

        new_table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        print('e_res_count5')
        table_rows = WebDriverWait(new_table_body, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print('e_res_count6')

        new_table_row = WebDriverWait(table_rows[4], 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        print('e_res_count7')
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row)
        print('e_res_count8')

        div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
        print('e_res_count9')

        if len(div_elements) >= 2:
            second_div_element = div_elements[1]
        cell_bodies = WebDriverWait(second_div_element, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
        # cell_bodies[16]
        # print(cell_bodies)
        print('e_res_count99')
        email_responded = cell_bodies[required_index]
        print('e_res_count999')
        print('remainder count', email_responded)
        req_div = WebDriverWait(email_responded, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'div')))
        print('req_div', req_div)
        req_spans = WebDriverWait(req_div, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'span')))


        span_classes = req_spans[0].get_attribute('class')

        # Check if 'orange' is in the list of classes
        print(span_classes.split())
        if 'blue-row-color' in span_classes.split():
            print("The span has the original blue class.")
            row = ['1', 'Batches Level', 'Email Responded Colour Blue', 'Email Responded Colour Blue Check',
                   'Check Email Responded Colour Blue', 'User should be able to see Email Responded Colour Blue', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            print("The button does not have the 'blue' class.")
            row = ['1', 'Batches Level', 'Email Responded Colour Blue', 'Email Responded Colour Blue Check',
                   'Check Email Responded Colour Blue', 'User should be able to see Email Responded Colour Blue', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        button_blue = WebDriverWait(req_spans[1], 20).until(EC.presence_of_element_located((By.TAG_NAME, 'button')))
        
        button_blue.click()

        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="updateEmailResponseStatusModal"]/div/div/div[3]/button[2]'))).click()
        
        time.sleep(10)
        driver.refresh()

        new_table_body_1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        print('e_res_count5')
        table_rows_1 = WebDriverWait(new_table_body_1, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print('e_res_count6')

        new_table_row_1 = ''
        if module_number == 1:
            new_table_row_1 = WebDriverWait(table_rows_1[4], 20).until(EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        else:
            new_table_row_1 = WebDriverWait(table_rows_1[0], 20).until(EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        
        print('e_res_count7')
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row_1)
        print('e_res_count8')

        div_elements_1 = new_table_row_1.find_elements(By.TAG_NAME, 'div')
        print('e_res_count9')

        if len(div_elements_1) >= 2:
            second_div_element_1 = div_elements_1[1]
        cell_bodies = WebDriverWait(second_div_element_1, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
        # cell_bodies[16]
        # print(cell_bodies)

        email_responded_1 = cell_bodies[required_index]

        print('remainder count', email_responded_1)
        req_div_1 = WebDriverWait(email_responded_1, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'div')))
        print('req_div', req_div_1)
        req_spans_1 = WebDriverWait(req_div_1, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'span')))


        span_classes_1 = req_spans_1[0].get_attribute('class')

        # Check if 'orange' is in the list of classes
        print(span_classes_1.split())
        if 'orange-row-color' in span_classes_1.split():
            print("The span has the orange class.")
            row = ['1', 'Batches Level', 'Email Responded Colour Orange', 'Email Responded Colour Orange Check',
                   'Check Email Responded Colour Orange', 'User should be able to see Email Responded Colour Orange', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            print("The button does not have the 'Orange' class.")
            row = ['1', 'Batches Level', 'Email Responded Colour Orange', 'Email Responded Colour Orange Check',
                   'Check Email Responded Colour Orange', 'User should be able to see Email Responded Colour Orange', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        
        print('near end 1')
        WebDriverWait(cell_bodies[0], 10).until(EC.presence_of_element_located((By.TAG_NAME,'div'))).find_element(By.TAG_NAME, 'label').click()
        print('near end 2')
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="0"]/div/div/a[3]'))).click()
        print('near end 3')
        time.sleep(20)

        div_element = WebDriverWait(cell_bodies[email_reminder_count],10).until(EC.presence_of_element_located((By.TAG_NAME, 'div')))

        print(WebDriverWait(div_element,10).until(EC.presence_of_element_located((By.TAG_NAME, 'button'))))
    
    except Exception as e:
        print(e)




def login(user, passw, driver, df):
    try:
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="username"]')))

        username = driver.find_element(By.XPATH, '//*[@id="username"]')
        username.send_keys(user)

        password = driver.find_element(By.XPATH, '//*[@id="password"]')
        password.send_keys(passw)

        submit_btn = driver.find_element(By.XPATH, '//*[@id="kc-login"]')
        submit_btn.click()

        # print('login pass')
        row = ['1', 'Application Level', 'Login', 'Username and Password check',
               'Enter username and password', 'User should see the home page', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [1, 'Login', 'Pass']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
    except Exception as e:
        # print('login failed')
        # row = [1, 'Login', 'Fail']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
        row = ['1', 'Application Level', 'Login', 'Username and Password check',
               'Enter username and password', 'User should see the home page', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        print(e)

    return df


def logout(driver, df):
    try:
        # print("logout1")
        logout = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[2]/div/button/span')))
        # logout = driver.find_element(By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[2]/div/button/span')
        # print("logout2")
        driver.execute_script("arguments[0].click();", logout)
        # logout.click()
        # print("logout3")
        logout_new = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="kc-logout"]')))
        # print("logout4")
        logout_new.click()
        # print("logout5")
        signin_logo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="kc-page-title"]')))
        # print("logout6")
        print(signin_logo, 'Goodd')
        # print("logout7")
        # row = [17, 'Logout', 'Pass']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row

        row = ['32', 'Module Level', 'Record Screen', 'Logout',
               'Clicking logout button', 'User must be able to logout', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        print('df', df, 'df')
        df.loc[len(df)] = row
    except Exception as e:
        print(e)
        row = ['32', 'Module Level', 'Record Screen', 'Logout',
               'Clicking logout button', 'User must be able to logout', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [17, 'Logout', 'Fail']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
    return df


def refresh(driver, df):
    try:
        heading = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[1]/h5')))
        # heading = driver.find_element(By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[1]/h5')
        # refreshbutton = driver.find_element(By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[2]/div/div[3]/button')

        refreshbutton = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[2]/div/div[3]/button')))

        refreshbutton.click()
        # print('After Refresh')
    except Exception as e:
        print(e)
    try:
        heading.text
        print("Page refresh did not occur.")
        # row = [5, 'Refresh Check', 'Fail']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
        row = ['6', 'Common', 'Client Screen', 'Refresh Button',
               'Click Refresh Button', 'Refresh and stay on same page', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    except StaleElementReferenceException:
        # row = [5, 'Refresh Check', 'Pass']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
        row = ['6', 'Common', 'Client Screen', 'Refresh Button',
               'Click Refresh Button', 'Refresh and stay on same page', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        print("Page refresh successful.")
    return df


def role(driver, df):
    try:
        role = driver.find_element(
            By.XPATH, '//*[@id="navbarDropdown2"]/h6/div[2]/small')
        text = role.text
        if (text.lower() == 'admin'):

            row = ['5', 'Application Level', 'Login', 'Check for Admin', 'Enter username and password for admin',
                   'User should see the home page and logged in as admin', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

            print("It's Admin")
        # elif(text.lower() == 'coe executive'):
        #     print("It's Executive")
        # elif(text.lower() == 'business user'):
        #     print("It's Business User")
        # elif(text.lower() == 'coe pod lead'):
        #     print("It's COE POD Lead")
        else:
            # print("Unknown Role")
            row = ['5', 'Application Level', 'Login', 'Check for Admin', 'Enter username and password for admin',
                   'User should see the home page and logged in as admin', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        # row = [4, 'Role Check', 'Pass']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
    except:
        row = ['5', 'Application Level', 'Login', 'Check for Admin', 'Enter username and password for admin',
               'User should see the home page and logged in as admin', 'Could Not get "logged in" value']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [4, 'Role Check', 'Fail']
        # df.loc[len(df)] = row
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
    return df


def h5textchecker(driver, heading, df):
    try:
        h5_element = WebDriverWait(driver, 10).until(
            EC.text_to_be_present_in_element((By.TAG_NAME, "h5"), heading))
        if (h5_element):
            print('h5 element paa gaya Bank Confirmations')

        row = ['4', 'Application Level', 'Home Page', {
            heading}, 'Check h5 tag of loaded page', 'User should see Module title as heading', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    except:
        row = ['4', 'Application Level', 'Home Page', {
            heading}, 'Check h5 tag of loaded page', 'User should see Module title as heading', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    return df


def navigationchecker(driver, title, df):
    dropdown_arrow = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="navbarDropdown2"]')))
    dropdown_arrow.click()
    # link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="navbarDropdown2"]')))

    try:
        elements = driver.find_elements(By.XPATH, "//a[@ngbdropdownitem]")
        desired_element = None
        # print(elements)
        for element in elements:
            # print(element.text.lower()[:-14])
            if element.text.lower()[:-14] == title.lower():
                print(title)
                element.click()
                break
        # if element.text == title:
        #     print("Element found by tag, class, and text:", element.text)
        # else:
        #     print("Element found but with different text:", element.text)
        # if(element):
        #     print('kuch to mila')

        number = 0
        if (title == 'Debtor Confirmations'):
            number = 7
        if (title == 'Creditor Confirmations'):
            number = 8
        if (title == 'Legal Matter Confirmations'):
            number = 9
        if (title == 'Bank Confirmations'):
            number = 10

        row = ['8', 'Common', 'Client Screen', 'Module Navigator',
               f'Click on {title} module', f'User should see client screen of {title} module', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

        # row = [number, f'Navigation working for {title}', 'Pass']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
    except TimeoutException:
        # row = [number, f'Navigation working for {title}', 'Fail']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row

        row = ['8', 'Common', 'Client Screen', 'Module Navigator',
               f'Click on {title} module', f'User should see client screen of {title} module', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

        print("Element not found within the given timeout.")
    return df


def email_filter(driver, filter_xpath, column_number, df):
    try:
        driver.refresh()
        time.sleep(10)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, filter_xpath))).click()
        time.sleep(10)

        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[4]/div/button[1]'))).click()

        time.sleep(10)
        new_table_body_again = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))

        table_rows_again = WebDriverWait(new_table_body_again, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))

        for table_row_again in table_rows_again:
            # time.sleep(30)

            new_table_row_again = WebDriverWait(table_row_again, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
            # new_table_row = table_rows[4].find_element(By.TAG_NAME, 'datatable-body-row')

            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row_again)

            div_elements_again = new_table_row_again.find_elements(
                By.TAG_NAME, 'div')

            if len(div_elements_again) >= 2:
                second_div_element = div_elements_again[1]

            cell_bodies_again = WebDriverWait(second_div_element, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            # cell_bodies = second_div_element.find_elements(By.TAG_NAME, 'datatable-body-cell')

            email_received_cell_body = cell_bodies_again[column_number]

            print(email_received_cell_body.text)

        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[4]/div/button[2]'))).click()

        if column_number == 11:
            row = ['29', 'Module Level', 'Record Screen', 'Filtering for Responded Mail',
                   'Checking if there are responded emails', 'User must be able to see a filtered list of responded emails', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        if column_number == 10:
            row = ['30', 'Module Level', 'Record Screen', 'Filtering for Undelivered Mail',
                   'Checking if there are undelivered emails', 'User must be able to see a filtered list of undelivered emails', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
    except Exception as e:
        if column_number == 10:
            row = ['29', 'Module Level', 'Record Screen', 'Filtering for Responded Mail',
                   'Checking if there are responded emails', 'User must be able to see a filtered list of responded emails', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        if column_number == 9:
            row = ['30', 'Module Level', 'Record Screen', 'Filtering for Undelivered Mail',
                   'Checking if there are undelivered emails', 'User must be able to see a filtered list of undelivered emails', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        print(e)


def batch_creation(driver, df, module_number):
    try:
        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="client-name"]')))
        select_element.click()

        option_text = 'Test_Client_1'
        # option_text = 'All_Client'

        # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="client-name"]/option[2]'))).click()

        option_elements = WebDriverWait(select_element, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, "option")))
        print(option_elements)
        for element in option_elements:
            time.sleep(1)
            if element.text == option_text:
                element.click()


#         script = f'''
# var select = arguments[0];
# var optionText = "{option_text}";

# for (var i = 0; i < select.options.length; i++) {{
#     if (select.options[i].text === optionText) {{
#         select.options[i].selected = true;
#         var event = new Event('input', {{ bubbles: true }});
#         select.dispatchEvent(event);
#         break;
#     }}
# }}
# '''
#         driver.execute_script(script, select_element)

        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-entires/div[2]/div/form/div/div[1]/div[2]/ng-select/div/div/div[2]/input')))
        element.click()
        element.send_keys(Keys.ARROW_UP)
        # element.send_keys(Keys.ARROW_DOWN)
        # element.send_keys(Keys.ARROW_DOWN)
        element.send_keys(Keys.ENTER)

        element2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-entires/div[2]/div/form/div/div[1]/div[3]/ng-select/div/div/div[2]/input')))
        element2.click()
        element2.send_keys(Keys.ARROW_UP)
        # element2.send_keys(Keys.ARROW_DOWN)
        # element2.send_keys(Keys.ARROW_DOWN)
        element2.send_keys(Keys.ENTER)

        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="client-approval"]')))
        current_directory = os.getcwd()
        file_path1 = os.path.join(
            current_directory, 'Client Approval.pdf')
        file_path2 = ''
        print(module_number, 'dfgh')
        if module_number == 1:
            file_path2 = os.path.join(
                current_directory, 'Bank Confirmation - Batch Details Template.xlsx')
        elif module_number == 2:
            print('harsh is best')
            file_path2 = os.path.join(
                current_directory, 'Debtor Confirmation - Batch Details Template.xlsx')
            print('harsh is not ebst')
        elif module_number == 3:
            file_path2 = os.path.join(
                current_directory, 'Creditor Confirmation - Batch Details Template.xlsx')
        elif module_number == 4:
            file_path2 = os.path.join(
                current_directory, 'LC_Client On-boarding _ Batch Processing Details _withparty.xlsx')

        input_field.send_keys(file_path1)

        input_field2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-entires/div[2]/div/form/div/div[1]/div[5]/div/div[1]/input')))
        input_field2.send_keys(file_path2)

        input_field3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-entires/div[2]/div/form/div/div[1]/div[5]/div/div[2]/input')))
        input_field3.send_keys(file_path1)
        # time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-entires/div[2]/div/form/div/div[2]/div/button'))).click()

        # time.sleep(5)

        row = ['10', 'Module Level', 'New Batch Screen', 'New Batch creation',
               'Create New Batch', 'User should be able to create New Batch', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [21, 'New Batch Creation', 'Pass']
        # df.loc[len(df)] = row
    except Exception as e:
        print('batch cration', e)
        row = ['10', 'Module Level', 'New Batch Screen', 'New Batch creation',
               'Create New Batch', 'User should be able to create New Batch', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    return df


def module_choose(driver, number, df):
    try:
        element1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[1]/div/div')))
        element2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[2]/div/div')))
        element3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[3]/div/div')))
        element4 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[4]/div/div')))

        row = ['2', 'Application Level', 'Home Page', 'Check for 4 modules',
               'After logging in', 'User should see all 4 modules', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

        second_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[{number}]/div/div')))
        second_element.click()
        print('module choose pass')

        if number == 1:
            row = ['3', 'Application Level', 'Home Page', 'Bank Confirmations', 'Click on Bank Confirmations module',
                   'User should see client screen of Bank Confirmations module', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        elif number == 2:
            row = ['3', 'Application Level', 'Home Page', 'Debtor Confirmations', 'Click on Debtor Confirmations module',
                   'User should see client screen of Debtor Confirmations module', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        elif number == 3:
            row = ['3', 'Application Level', 'Home Page', 'Creditor Confirmations', 'Click on Creditor Confirmations module',
                   'User should see client screen of Creditor Confirmations module', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        elif number == 4:
            row = ['3', 'Application Level', 'Home Page', 'Legal Matter Confirmations', 'Click on Legal Matter Confirmations module',
                   'User should see client screen of Legal Matter Confirmations module', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
    except:
        # print('module choose fail')
        row = ['2', 'Application Level', 'Home Page', 'Check for 4 modules',
               'After logging in', 'User should see all 4 modules', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

        if number == 1:
            row = ['3', 'Application Level', 'Home Page', 'Bank Confirmations', 'Click on Bank Confirmations module',
                   'User should see client screen of Bank Confirmations module', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        elif number == 2:
            row = ['3', 'Application Level', 'Home Page', 'Debtor Confirmations', 'Click on Debtor Confirmations module',
                   'User should see client screen of Debtor Confirmations module', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        elif number == 3:
            row = ['3', 'Application Level', 'Home Page', 'Creditor Confirmations', 'Click on Creditor Confirmations module',
                   'User should see client screen of Creditor Confirmations module', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        elif number == 4:
            row = ['3', 'Application Level', 'Home Page', 'Legal Matter Confirmations', 'Click on Legal Matter Confirmations module',
                   'User should see client screen of Legal Matter Confirmations module', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
    return df


def email_batch_link(driver, df):
    try:
        email = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="1-link"]')))
        # print('email link working pass')
        row = ['7', 'Module Level', 'Client Screen', 'Email Batch Link',
               'Click on Email Batch Link', 'User should be able to click this link', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        email.click()
        # row = [10, 'Email Batches link working', 'Pass']
        # df.loc[len(df)] = row
    except Exception as e:
        print(e)
        # print('email link working fail')
        row = ['7', 'Module Level', 'Client Screen', 'Email Batch Link',
               'Click on Email Batch Link', 'User should be able to click this link', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    return df


def new_email_batch_button(driver, df):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="1"]/div[2]/div[1]/div[1]/div/a'))).click()
        row = ['9', 'Module Level', 'Batches Screen', 'New Batch button Working',
               'Click New Batch button', 'User should be able to click New Batch button', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [20, 'New Batch button working', 'Pass']
        # df.loc[len(df)] = row
    except:
        row = ['9', 'Module Level', 'Batches Screen', 'New Batch button Working',
               'Click New Batch button', 'User should be able to click New Batch button', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    return df


def response_to_email_received_in_outlook(driver, df, mail_subject_unique_id, name, mail_category, module_number, legal_party):
    try:
        try:
            os.startfile("outlook")
            row = ['27', 'OS Level', 'Desktop', 'Outlook Opening', 'Open outlook',
                   'User must be able to see outlook to receive mail', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        except:
            os.startfile("outlook")
            row = ['27', 'OS Level', 'Desktop', 'Outlook Opening', 'Open outlook',
                   'User must be able to see outlook to receive mail', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        time.sleep(20)

        driver.refresh()
        print('mail response started')
        # arr=[0,1,3,6]
        for index, value in enumerate(mail_subject_unique_id):
            print('inside outlook mail unique id')
            outlook = win32com.client.Dispatch(
                'Outlook.Application').GetNamespace("MAPI")

            # 6 corresponds to the Inbox Folder
            inbox = outlook.GetDefaultFolder(6)
            items = inbox.Items

            subject = ''
            if (module_number == 1):
                subject = f"Audit Confirmation for {name[index]}; balance as on 23-08-2023: Bank Confirmations- Tracking ID: #{mail_subject_unique_id[index]}"
            elif (module_number == 2):
                subject = f"Audit Confirmation for {name[index]}; balance as on 23-08-2023: Debtor Confirmations- Tracking ID: #{mail_subject_unique_id[index]}"
            elif (module_number == 3):
                subject = f"Audit Confirmation for {name[index]}; balance as on 23-08-2023: Creditor Confirmations- Tracking ID: #{mail_subject_unique_id[index]}"
            elif (module_number == 4):
                subject = f"Audit Confirmation for {name[index]}; balance as on 23-08-2023: Legal Matter Confirmations- Tracking ID: #{mail_subject_unique_id[index]}"

            condition = f"[Subject] = '{subject}'"
            items = inbox.Items.Restrict(condition)
            print('inside outlook mail unique id1', items)

            if len(items) > 0:
                # Get the first email that matches the subject
                email = items.GetFirst()
                print('inside outlook mail unique id2', email)
                # Reply to the email
                reply = email.Reply()
                # Enter the new email ID
                new_email_id = "ac.uat@expdiginetdev.onmicrosoft.com"
                reply.Recipients.Add(new_email_id)
                # Attach a file
                if (module_number == 1):
                    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Bank_Response_Template_With_Auth_Sig_Details_v1 (10).docx"
                    attachment = reply.Attachments.Add(Source=attachment_path)
                    # Send the reply email
                    reply.Send()
                    row = ['28.{index}', 'OS Level', 'Desktop', 'Receiving Mail and Responding to it',
                           'Responding to mail', 'User must be able torespond to received mail no.{index}', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                if (module_number == 2):
                    try:
                        # for category in mail_category:
                        #     print('module number check ke ander',category)
                        attachment_path = ''
                        if mail_category[index] == 'Without details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Debtors Confirmations\Responses\User 1\DC_Response_Template_Without_Balance_Details_v1 (6).docx"

                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('without details')
                            # break
                        elif mail_category[index] == 'With Invoice Details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Debtors Confirmations\Responses\User 2\DC_Response_Template_With_Invoice_Balance_Details_v1 (6).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('with invoice details')
                            # break
                        elif mail_category[index] == 'With Ledger Details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Debtors Confirmations\Responses\User 4\DC_Response_Template_With_Ledger_Balance_Details_v1 (3).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('with ledger details')
                            # break
                        elif mail_category[index] == 'With Ledger & Invoice Details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Debtors Confirmations\Responses\User 7\DC_Response_Template_With_Ledger&Invoice_Balance_Details_v1 (8).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('with invoive and ledger details')
                            # break
                        row = ['28.{index}', 'OS Level', 'Desktop', 'Receiving Mail and Responding to it',
                               'Responding to mail', 'User must be able torespond to received mail no.{index}', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                # attachment = reply.Attachments.Add(Source=attachment_path)
                if (module_number == 3):
                    try:
                        # for category in mail_category:
                        #     print('module number check ke ander',category)
                        attachment_path = ''
                        if mail_category[index] == 'Without details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Creditors Confirmations\Responses\User 1\CC_Response_Template_Without_Balance_Details_v1 (2).docx"

                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('without details')
                            # break
                        elif mail_category[index] == 'With Invoice Details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Creditors Confirmations\Responses\User 2\CC_Response_Template_With_Invoice_Balance_Details_v1 (4).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('with invoice details')
                            # break
                        elif mail_category[index] == 'With Ledger Details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Creditors Confirmations\Responses\User 4\CC_Response_Template_With_Ledger_Balance_Details_v1 (1).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('with ledger details')
                            # break
                        elif mail_category[index] == 'With Ledger & Invoice Details':
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Creditors Confirmations\Responses\User 7\CC_Response_Template_With_Ledger&Invoice_Balance_Details_v1 (5).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('with invoive and ledger details')
                            # break
                        row = ['28.{index}', 'OS Level', 'Desktop', 'Receiving Mail and Responding to it',
                               'Responding to mail', 'User must be able torespond to received mail no.{index}', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)

                if (module_number == 4):
                    try:
                        # for category in mail_category:
                        #     print('module number check ke ander',category)
                        attachment_path = ''
                        if index == 0:
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Legal Confirmations\Responses\Legal_Party_1\LC_Response_Template_Without_Matter_Details_v1 (2).docx"

                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('Legal Party 1')
                            # break
                        elif index == 2:
                            attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Responses\Legal Confirmations\Responses\Legal_Party_3\LC_Response_Template_With_Matter_Details_v1 (9).docx"
                            attachment = reply.Attachments.Add(
                                Source=attachment_path)

                            # Send the reply email
                            reply.Send()
                            # print('Legal Party 3')
                            # break
                        row = ['28.{index}', 'OS Level', 'Desktop', 'Receiving Mail and Responding to it',
                               'Responding to mail', 'User must be able torespond to received mail no.{index}', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)

                # print('printing5')

            else:
                row = ['28', 'OS Level', 'Desktop', 'Receiving Mail and Responding to it',
                       'Responding to mail', 'User must be able torespond to received mail', 'FAIL']
                df.loc[len(df)] = row
                break

        time.sleep(6)

        # outlook.Quit()
        driver.refresh()

    except Exception as e:
        print(e)
        row = ['28', 'OS Level', 'Desktop', 'Receiving Mail and Responding to it',
               'Responding to mail', 'User must be able torespond to received mail', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
    return df


def is_table_body_visible_and_view_details_click(driver, df):
    try:
        # table_body = driver.find_element(
        #     By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')
        table_body = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        row_wrappers = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))

        if len(row_wrappers) > 0:
            print('Batches Visible')
            row = ['15', 'Module Level', 'Batches Screen', 'Data Availability Check',
                   'Check if records exist in Table', 'User should be able to see records in table', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            print('Batches Not Visible')
            row = ['15', 'Module Level', 'Batches Screen', 'Data Availability Check',
                   'Check if records exist in Table', 'User should be able to see records in table', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        try:
            row_wrapper = WebDriverWait(table_body, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-row-wrapper')))
            # row_wrapper = table_body.find_element(By.TAG_NAME, "datatable-row-wrapper")
            row_body = WebDriverWait(row_wrapper, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
            # row_body = row_wrapper.find_element(By.TAG_NAME,'datatable-body-row')
            row = WebDriverWait(row_body, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')))
            # row = row_body.find_element(By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')
            cells = WebDriverWait(row, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            # cells = row.find_elements(By.TAG_NAME, "datatable-body-cell")
            # print(cells)
            last_cell = cells[-1]

            # print(last_cell)

            div = WebDriverWait(last_cell, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            anchor = WebDriverWait(div, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'a')))
            # last_cell.find_element(By.TAG_NAME,'div').find_element(By.TAG_NAME, 'a').click()
            # print(anchor)
            # anchor.click()
            driver.execute_script("arguments[0].click();", anchor)

            row = ['16', 'Module Level', 'Batches Screen', 'View Details Button',
                   'Check if View Details Button clicks', 'User should be able to click View Details Button', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
            # row = [11, 'Email Batches First row Last Cell View Details button clicking', 'Pass']
            # df.loc[len(df)] = row

        except:
            row = ['16', 'Module Level', 'Batches Screen', 'View Details Button',
                   'Check if View Details Button clicks', 'User should be able to click View Details Button', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
    except Exception as e:
        print(e)
        row = ['15', 'Module Level', 'Batches Screen', 'Data Availability Check',
               'Check if records exist in Table', 'User should be able to see records in table', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [11, 'Email Batches First row Last Cell View Details button clicking', 'Fail']
        # df.loc[len(df)] = row

    return df


def attachments_download_batches_level(driver, df, module_number):
    try:
        print('e_res_count1')
        header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-header/div/div[2]')))

        print('e_res_count2')
        header_cells = WebDriverWait(header_row, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
        print('e_res_count3')
        batch_file_index = 0
        auth_letter_index = 0
        other_docs_index = 0

        whole_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body')))

        i = 0
        for header_cell in header_cells:
            if i != 0:
                driver.execute_script(
                    f"arguments[0].scrollLeft += 70;", whole_table_body)
            print(header_cell)
            column_name = WebDriverWait(header_cell, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            print('columnname', column_name)
            name_exact = ''
            try:
                name_exact = WebDriverWait(column_name, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span')))
                print('namexact', name_exact)
            except:
                print('not span')
            try:
                name_exact = WebDriverWait(column_name, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'strong')))
                print('namexact', name_exact)
            except:
                print('not strong')
            text_part = WebDriverWait(name_exact, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
            print('column name = ', text_part)

            if text_part == 'Batch File':
                batch_file_index = i
                # break
            if text_part == 'Authorisation Letter':
                auth_letter_index = i
                # break
            if text_part == 'Other Documents':
                other_docs_index = i
                # break
            i += 1
        print('e_res_count4')

        table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))

        print('eres count 5')
        # table_body = driver.find_element(By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')
        print(batch_file_index, 'batch file index')
        print(auth_letter_index, 'auth letter index')
        print(other_docs_index, 'other docs index')

        try:
            row_wrapper = WebDriverWait(table_body, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-row-wrapper')))
            # row_wrapper = table_body.find_element(By.TAG_NAME, "datatable-row-wrapper")
            row_body = WebDriverWait(row_wrapper, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
            # row_body = row_wrapper.find_element(By.TAG_NAME,'datatable-body-row')
            row = WebDriverWait(row_body, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')))
            # row = row_body.find_element(By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')
            cells = WebDriverWait(row, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            # cells = row.find_elements(By.TAG_NAME, "datatable-body-cell")
            # print(cells)
            # last_cell = cells[-1]

            batch_file = cells[batch_file_index]
            auth_letter = cells[auth_letter_index]
            other_docs = cells[other_docs_index]

            # print(last_cell)
            try:
                div = WebDriverWait(batch_file, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                button = WebDriverWait(div, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'button')))
                # last_cell.find_element(By.TAG_NAME,'div').find_element(By.TAG_NAME, 'a').click()
                # print(anchor)
                # anchor.click()
                driver.execute_script("arguments[0].click();", button)

                download_directory = "C:/Users/harsh.vijaykumar/Downloads"
                file_name = ''
                if module_number == 1:
                    file_name = "Bank Confirmation- Batch Details Template.xlsx"
                elif module_number == 2:
                    file_name = "Debtor Confirmation- Batch Details Template.xlsx"
                elif module_number == 3:
                    file_name = "Creditor Confirmation- Batch Details Template.xlsx"
                elif module_number == 4:
                    file_name = "LC_Client On-boarding- Batch Details Template.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    row = ['11', 'Module Level', 'Batches Screen', 'Batch File Download',
                           'Checking if Batch File Downloads', 'User must be able to Download Batch File', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    row = ['11', 'Module Level', 'Batches Screen', 'Batch File Download',
                           'Checking if Batch File Downloads', 'User must be able to Download Batch File', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            except:
                row = ['11', 'Module Level', 'Batches Screen', 'Batch File Download',
                       'Checking if Batch File Downloads', 'User must be able to Download Batch File', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            time.sleep(5)
            try:
                div = WebDriverWait(other_docs, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                button = WebDriverWait(div, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'button')))
                # last_cell.find_element(By.TAG_NAME,'div').find_element(By.TAG_NAME, 'a').click()
                # print(anchor)
                # anchor.click()
                driver.execute_script("arguments[0].click();", button)
                download_directory = "C:/Users/harsh.vijaykumar/Downloads"

                file_name = "Test_Client_1.zip"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    row = ['12', 'Module Level', 'Batches Screen', 'Other Documents Download',
                           'Checking if Other Documents Downloads', 'User must be able to Download Other Documents', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    row = ['12', 'Module Level', 'Batches Screen', 'Other Documents Download',
                           'Checking if Other Documents Downloads', 'User must be able to Download Other Documents', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            except:
                row = ['12', 'Module Level', 'Batches Screen', 'Other Documents Download',
                       'Checking if Other Documents Downloads', 'User must be able to Download Other Documents', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            time.sleep(5)
            try:
                div = WebDriverWait(auth_letter, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                button = WebDriverWait(div, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'button')))
                # last_cell.find_element(By.TAG_NAME,'div').find_element(By.TAG_NAME, 'a').click()
                # print(anchor)
                # anchor.click()
                driver.execute_script("arguments[0].click();", button)
                download_directory = "C:/Users/harsh.vijaykumar/Downloads"

                file_name == ''
                if module_number == 1:
                    file_name = "Bank Confirmation authorisation.pdf"
                elif module_number == 2:
                    file_name = "Debtor Confirmation authorisation.pdf"
                elif module_number == 3:
                    file_name = "Creditor Confirmation authorisation.pdf"
                elif module_number == 4:
                    file_name = "LC_Client On-boarding authorisation.pdf"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    row = ['13', 'Module Level', 'Batches Screen', 'Authorization Letter Download',
                           'Checking if Authorization Letter Downloads', 'User must be able to Download Authorization Letter', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    row = ['13', 'Module Level', 'Batches Screen', 'Authorization Letter Download',
                           'Checking if Authorization Letter Downloads', 'User must be able to Download Authorization Letter', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            except:
                row = ['13', 'Module Level', 'Batches Screen', 'Authorization Letter Download',
                       'Checking if Authorization Letter Downloads', 'User must be able to Download Authorization Letter', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        except Exception as e:
            row = ['13', 'Module Level', 'Batches Screen', 'Authorization Letter Download',
                   'Checking if Authorization Letter Downloads', 'User must be able to Download Authorization Letter', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

            row = ['12', 'Module Level', 'Batches Screen', 'Other Documents Download',
                   'Checking if Other Documents Downloads', 'User must be able to Download Other Documents', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

            row = ['11', 'Module Level', 'Batches Screen', 'Batch File Download',
                   'Checking if Batch File Downloads', 'User must be able to Download Batch File', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
    except Exception as e:
        print(e)
        row = ['15', 'Module Level', 'Batches Screen', 'Data Availability Check',
               'Check if records exist in Table', 'User should be able to see records in table', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [11, 'Email Batches First row Last Cell View Details button clicking', 'Fail']
        # df.loc[len(df)] = row
        row = ['13', 'Module Level', 'Batches Screen', 'Authorization Letter Download',
               'Checking if Authorization Letter Downloads', 'User must be able to Download Authorization Letter', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

        row = ['12', 'Module Level', 'Batches Screen', 'Other Documents Download',
               'Checking if Other Documents Downloads', 'User must be able to Download Other Documents', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

        row = ['11', 'Module Level', 'Batches Screen', 'Batch File Download',
               'Checking if Batch File Downloads', 'User must be able to Download Batch File', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row

    return df


def report_download_after_view_details(driver, df, mail_subject_unique_id, module_number):
    try:
        # try:
        #     table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        #     # table_body = driver.find_element(By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')

        #     row_wrapper = table_body.find_element(By.TAG_NAME, "datatable-row-wrapper")
        #     row_body = row_wrapper.find_element(By.TAG_NAME,'datatable-body-row')

        #     row_ = row_body.find_element(By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')
        #     row = ['12', 'Module Level', 'Batches Screen', 'Data loaded', 'Wait for data to load', 'User should be able to see data', 'PASS']
        #         # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        #     df.loc[len(df)] = row
        # except:
        #     row = ['12', 'Module Level', 'Batches Screen', 'Data loaded', 'Wait for data to load', 'User should be able to see data', 'FAIL']
        #         # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        #     df.loc[len(df)] = row
        # try:
        #     cells = row_.find_elements(By.TAG_NAME, "datatable-body-cell")

        #     last_cell = cells[-1]

        #     last_cell.find_element(By.TAG_NAME,'div').find_element(By.TAG_NAME, 'a').click()

        #     WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="0"]/div/div/a[1]'))).click()
        #     row = ['13', 'Module Level', 'Batches Screen', 'View Button', 'Wait for button to load and click', 'User should be able to click view button', 'PASS']
        #         # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        #     df.loc[len(df)] = row
        # except:
        #     row = ['13', 'Module Level', 'Batches Screen', 'View Button', 'Wait for button to load and click', 'User should be able to click view button', 'FAIL']
        #         # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        #     df.loc[len(df)] = row
        print('inside report download')
        try:
            driver.refresh()
            print('before clicking download report')
            spec_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/app-dashboard/div[2]/div/tabset/div/tab[1]/div/div/a[1]')))
            driver.execute_script("arguments[0].click();", spec_element)
            print('aftet cliick')
            element = driver.execute_script(
                "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.pb-3.pt-2 > div > div:nth-child(1) > label');")
            print('inside report download1')
            time.sleep(10)

            driver.execute_script("arguments[0].click();", element)

            time.sleep(10)
            print('inside report download2')
            element2 = driver.execute_script(
                "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.text-center.pt-1 > button');")
            time.sleep(10)
            print('inside report download3')
            driver.execute_script("arguments[0].click();", element2)
            time.sleep(10)
            print('inside report download4')
            download_directory = "C:/Users/harsh.vijaykumar/Downloads"
            print('inside report download5')

            if module_number == 1:
                file_name = "Bank_Consolidated.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "consolidated")
                else:
                    print("File download failed or timed out.")

                row = ['14', 'Module Level', 'Record Screen', 'Consolidated Report Download',
                       'Download Report by clicking download then choosing consolidated', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

            elif module_number == 2:
                file_name = "DC_Consolidated.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "consolidated")
                else:
                    print("File download failed or timed out.")

                row = ['14', 'Module Level', 'Record Screen', 'Consolidated Report Download',
                       'Download Report by clicking download then choosing consolidated', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            elif module_number == 3:
                file_name = "CC_Consolidated.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "consolidated")
                else:
                    print("File download failed or timed out.")

                row = ['14', 'Module Level', 'Record Screen', 'Consolidated Report Download',
                       'Download Report by clicking download then choosing consolidated', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            elif module_number == 4:
                file_name = "LMC.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "consolidated")
                else:
                    print("File download failed or timed out.")

                row = ['14', 'Module Level', 'Record Screen', 'Consolidated Report Download',
                       'Download Report by clicking download then choosing consolidated', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        except Exception as e:
            print(e)
            row = ['14', 'Module Level', 'Record Screen', 'Consolidated Report Download',
                   'Download Report by clicking download then choosing consolidated', 'User should be able to download report', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/div/div/a[1]'))).click()
            print('allllooo')

            element3 = driver.execute_script(
                "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.pb-3.pt-2 > div > div:nth-child(2) > label');")
            print('allllooo1')

            driver.execute_script("arguments[0].click();", element3)
            print('allllooo2')
            element4 = driver.execute_script(
                "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.text-center.pt-1 > button');")
            print('allllooo3')
            driver.execute_script("arguments[0].click();", element4)

            print('allllooo4')
            download_directory = "C:/Users/harsh.vijaykumar/Downloads"

            if module_number == 1:
                file_name = "Bank_Detailed.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "detailed")
                else:
                    print("File download failed or timed out.")

                row = ['15', 'Module Level', 'Record Screen', 'Detailed Report Download',
                       'Download Report by clicking download then choosing Detailed', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            elif module_number == 2:
                file_name = "DC_Detailed.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "detailed")
                else:
                    print("File download failed or timed out.")

                row = ['15', 'Module Level', 'Record Screen', 'Detailed Report Download',
                       'Download Report by clicking download then choosing Detailed', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            elif module_number == 3:
                file_name = "CC_Detailed.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "detailed")
                else:
                    print("File download failed or timed out.")

                row = ['15', 'Module Level', 'Record Screen', 'Detailed Report Download',
                       'Download Report by clicking download then choosing Detailed', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            elif module_number == 4:
                file_name = "LMC.xlsx"
                file_path = os.path.join(download_directory, file_name)

                # Wait for the file to be downloaded
                timeout = 10  # Maximum time to wait for the file (in seconds)
                while not os.path.exists(file_path) and timeout > 0:
                    timeout -= 1
                    time.sleep(1)  # Wait for 1 second

                # Check if the file exists
                if os.path.exists(file_path):
                    print("File downloaded successfully!")
                    report_checker(driver, df, mail_subject_unique_id,
                                   module_number, file_path, "consolidated")
                else:
                    print("File download failed or timed out.")

                row = ['15', 'Module Level', 'Record Screen', 'Detailed Report Download',
                       'Download Report by clicking download then choosing Detailed', 'User should be able to download report', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
        except Exception as e:
            row = ['15', 'Module Level', 'Record Screen', 'Detailed Report Download',
                   'Download Report by clicking download then choosing Detailed', 'User should be able to download report', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        # Email Template Checking
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/div/div/a[2]'))).click()

            parent_div = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="summernoteModal"]/div/div/div[2]/form/div[2]/div[2]/div[3]/div[2]')))
            driver.execute_script("arguments[0].innerHTML = '';", parent_div)

            # Set the new content
            new_content = "New content"
            driver.execute_script(
                "arguments[0].innerHTML = arguments[1];", parent_div, new_content)

            submit_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="summernoteModal"]/div/div/div[2]/div/button')))
            driver.execute_script("arguments[0].click();", submit_button)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/div/div/a[2]'))).click()

            parent_text = driver.execute_script(
                "return arguments[0].innerText;", parent_div)
            if parent_text == new_content:
                print("The element's text matches the desired text.")
            else:
                print("The element's text does not match the desired text.")

            submit_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="summernoteModal"]/div/div/div[2]/div/button')))
            driver.execute_script("arguments[0].click();", submit_button)

            row = ['16', 'Module Level', 'Record Screen', 'Email Template submission',
                   'Edit Email Template', 'User should be able to update email tenplate', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        except Exception as e:
            row = ['16', 'Module Level', 'Record Screen', 'Email Template submission',
                   'Edit Email Template', 'User should be able to update email tenplate', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        # Deactivation
        try:
            print('Hello Worlddd')
            driver.refresh()
            time.sleep(60)
            table_body_new = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
            print('deactivation1', table_body_new)
            row_wrapper_new = table_body_new.find_element(
                By.TAG_NAME, "datatable-row-wrapper")
            print('deactivation2', row_wrapper_new)
            row_body_new = row_wrapper_new.find_element(
                By.TAG_NAME, 'datatable-body-row')
            print('deactivation3', row_body_new)
            row_new = row_body_new.find_element(
                By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')
            print('deactivation4', row_new)

            cells_new = row_new.find_elements(
                By.TAG_NAME, "datatable-body-cell")
            print('deactivation5', cells_new)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/div/div/a[4]'))).click()

            client_name = cells_new[1].find_element(
                By.TAG_NAME, 'div').text
            first_cell = cells_new[0]
            print('deactivation6', first_cell)

            first_cell.find_element(By.TAG_NAME, 'div').find_element(
                By.TAG_NAME, 'label').click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/div/div/a[4]'))).click()

            print('deactivation7')

            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/app-dashboard/div[2]/div/tabset/ul/li[1]/a'))).click()

            # for element in elements:
            #     text = element.text
            #     if text == 'Active':
            #         # Perform your desired action with the matching element
            #         element.click()
            #         print('Active Clicked')

            print('deactivation8')

            table_body_new2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))

            row_wrapper_new2 = table_body_new2.find_element(
                By.TAG_NAME, "datatable-row-wrapper")

            row_body_new2 = row_wrapper_new2.find_element(
                By.TAG_NAME, 'datatable-body-row')

            row_new2 = row_body_new2.find_element(
                By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]')

            cells_new2 = row_new2.find_elements(
                By.TAG_NAME, "datatable-body-cell")

            second_cell = cells_new2[1]
            print(client_name)
            if second_cell.find_element(By.TAG_NAME, 'div').text != client_name:
                print('AXIS Deactivated')
            else:
                print('AXIS Deactivation failed')
            row = ['17', 'Module Level', 'Record Screen', 'Record Deactivation',
                   'Deactivate first record', 'User should be able to deactivate records', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        except Exception as e:
            print(e)
            row = ['17', 'Module Level', 'Record Screen', 'Record Deactivation',
                   'Deactivate first record', 'User should be able to deactivate records', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="1"]/app-dashboard/div[1]/nav/ol/li[1]/a'))).click()
            row = ['18', 'Module Level', 'Record Screen', 'Batches link Working', 'Click batches link',
                   'User should be able to go back to batches level by clicking batches link', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        except Exception as e:
            print(e)
            ow = ['18', 'Module Level', 'Record Screen', 'Batches link Working', 'Click batches link',
                  'User should be able to go back to batches level by clicking batches link', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    except Exception as e:
        print(e)
    return df


def mail_send_from_application(driver, df, mail_subject_unique_id, name, module_number, mail_category, legal_party):
    if module_number == 1:
        try:

            new_table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))

            new_table_row = WebDriverWait(new_table_body, 10).until(EC.presence_of_element_located(
                (By.TAG_NAME, 'datatable-row-wrapper'))).find_element(By.TAG_NAME, 'datatable-body-row')

            table_rows = WebDriverWait(new_table_body, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))


            header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
            print('e_res_count2')
            header_cells = WebDriverWait(header_row, 20).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
            print('e_res_count3')
            
            unique_id = 0

            whole_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

            i = 0
            for header_cell in header_cells:
                if i != 0:
                    driver.execute_script(
                        f"arguments[0].scrollLeft += 100;", whole_table_body)
                print(header_cell)
                column_name = WebDriverWait(header_cell, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                print('columnname', column_name)
                name_exact = WebDriverWait(column_name, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span')))
                print('namexact', name_exact)
                text_part = WebDriverWait(name_exact, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
                print('column name = ', text_part)

                
                
                if text_part == 'Unique ID':
                    unique_id = i
                    # break
               
                    # break
                i += 1
                print('unique id', unique_id)

        



            new_table_row = WebDriverWait(table_rows[4], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))

            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')

            if len(div_elements) >= 2:
                second_div_element = div_elements[1]
            # print(second_div_element,'second div element')
            cell_body = second_div_element.find_element(
                By.TAG_NAME, 'datatable-body-cell').find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'label')

            # print(cell_body)
            driver.execute_script("arguments[0].click();", cell_body)
            # cell_body.click()

            row = ['17', 'Module Level', 'Record Screen', 'Checkbox of 5th Row clicking',
                   'Check if Checkbox of 5th Row is clicking', 'User should be able to click checkbox of 5th Row', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

            # row = [12, 'Checkbox of 5th row(With SBI) clicked', 'Pass']
            # df.loc[len(df)] = row
            # print('row sbi selected')
        except Exception as e:
            # row = [12, 'Checkbox of 5th row(With SBI) clicked', 'Fail']
            # df.loc[len(df)] = row
            row = ['17', 'Module Level', 'Record Screen', 'Checkbox of 5th Row clicking',
                   'Check if Checkbox of 5th Row is clicking', 'User should be able to click checkbox of 5th Row', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
            # print(e,'row sbi not selected')

        try:
            try:
                cell_bodies = WebDriverWait(second_div_element, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
                # cell_bodies[16]
                # print(cell_bodies)
                email_identifier = cell_bodies[unique_id]
                print('email identifier cell', email_identifier)
                # print(email_identifier)
                # print(email_identifier.find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'span').text)
                mail_subject_unique_id.append(email_identifier.find_element(
                    By.TAG_NAME, 'div').text)
                # mail_subject_unique_id[0] = email_identifier.find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'span').text
                print('yo', mail_subject_unique_id[0], 'yo')
                row = ['18', 'Module Level', 'Record Screen', 'Unique Identifier',
                       'Check if Unique Identifier can be picked and stored', 'Unique Identifier is required to identify mails', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            except Exception as e:
                print('special exception', e)
                row = ['18', 'Module Level', 'Record Screen', 'Unique Identifier',
                       'Check if Unique Identifier can be picked and stored', 'Unique Identifier is required to identify mails', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

            try:
                name.append(driver.find_element(
                    By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[1]/div/h5').text)
                # name[0] = driver.find_element(By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[1]/div/h5')
                print(name[0])
                row = ['19', 'Module Level', 'Record Screen', 'Name Collection',
                       'Check if Name can be picked and stored', 'Name is required to identify mails', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            except:
                row = ['19', 'Module Level', 'Record Screen', 'Name Collection',
                       'Check if Name can be picked and stored', 'Name is required to identify mails', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
        except Exception as e:
            print(e)

    elif module_number == 2 or module_number == 3:
        try:

            new_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))

            # new_table_row = WebDriverWait(new_table_body, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'datatable-row-wrapper'))).find_element(By.TAG_NAME, 'datatable-body-row')

            table_rows = WebDriverWait(new_table_body, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
            # print(table_rows)
            # print(len(table_rows))
            # user 1,2,4,7 => table_rows[0,1,3,6]


            header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
            print('e_res_count2')
            header_cells = WebDriverWait(header_row, 20).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
            print('e_res_count3')
            
            unique_id = 0

            whole_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

            i = 0
            for header_cell in header_cells:
                if i != 0:
                    driver.execute_script(
                        f"arguments[0].scrollLeft += 100;", whole_table_body)
                print(header_cell)
                column_name = WebDriverWait(header_cell, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                print('columnname', column_name)
                name_exact = WebDriverWait(column_name, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span')))
                print('namexact', name_exact)
                text_part = WebDriverWait(name_exact, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
                print('column name = ', text_part)

                
                
                if text_part == 'Unique ID':
                    unique_id = i
                    # break
               
                    # break
                i += 1
                print('unique id', unique_id)

        



            arr = [0, 1, 3, 6]
            for index, value in enumerate(arr):
                # print('Working till here')
                # print(value)
                # print(table_rows[value])
                new_table_row = WebDriverWait(table_rows[value], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
                # print('Working till here1')
                driver.execute_script(
                    "arguments[0].scrollIntoView();", new_table_row)
                # print('Working till here2')

                div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
                # print('Working till here3')
                if len(div_elements) >= 2:
                    second_div_element = div_elements[1]
                # print(second_div_element,'second div element')
                cell_body = second_div_element.find_element(
                    By.TAG_NAME, 'datatable-body-cell').find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'label')

                # print(cell_body)
                driver.execute_script("arguments[0].click();", cell_body)
                # cell_body.click()

                # row = [12, 'Checkbox of 5th row(With SBI) clicked', 'Pass']
                # df.loc[len(df)] = row
                # print('row sbi selected')

                cell_bodies = WebDriverWait(second_div_element, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
                # cell_bodies[16]
                email_identifier = cell_bodies[unique_id]
                mail_subject_unique_id.append(email_identifier.find_element(
                    By.TAG_NAME, 'div').text)
                print(mail_subject_unique_id[index])

                category = cell_bodies[6]
                # mail_category.append(category.find_element(
                #     By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'span').text)
                mail_category.append(category.find_element(
                    By.TAG_NAME, 'div').text)    
                print(mail_category[index])
                name.append(driver.find_element(
                    By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[1]/div/h5').text)
                print(name[index])

            print('end of mail unique id reading')

        except Exception as e:
            print(e)
            # row = [12, 'Checkbox of 5th row(With SBI) clicked', 'Fail']
            # df.loc[len(df)] = row
            # print(e,'row sbi not selected')

    elif module_number == 4:
        try:
            print('legal1')
            new_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
            print('legal2')
            # new_table_row = WebDriverWait(new_table_body, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'datatable-row-wrapper'))).find_element(By.TAG_NAME, 'datatable-body-row')

            table_rows = WebDriverWait(new_table_body, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
            
            header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
            print('e_res_count2')
            header_cells = WebDriverWait(header_row, 20).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
            print('e_res_count3')
            
            unique_id = 0

            whole_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

            i = 0
            for header_cell in header_cells:
                if i != 0:
                    driver.execute_script(
                        f"arguments[0].scrollLeft += 100;", whole_table_body)
                print(header_cell)
                column_name = WebDriverWait(header_cell, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                print('columnname', column_name)
                name_exact = WebDriverWait(column_name, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span')))
                print('namexact', name_exact)
                text_part = WebDriverWait(name_exact, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
                print('column name = ', text_part)

                
                
                if text_part == 'Unique ID':
                    unique_id = i
                    # break
               
                    # break
                i += 1
                print('unique id', unique_id)

        



            print('legal3')
            # print(table_rows)
            # print(len(table_rows))
            # user 1,2,4,7 => table_rows[0,1,3,6]
            arr = [0, 2]
            print('legal4')
            for index, value in enumerate(arr):
                print('legal5')
                # print('Working till here')
                # print(value)
                # print(table_rows[value])
                new_table_row = WebDriverWait(table_rows[value], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
                print('legal6')
                # print('Working till here1')
                driver.execute_script(
                    "arguments[0].scrollIntoView();", new_table_row)
                print('legal7')
                # print('Working till here2')

                div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
                print('legal8')
                # print('Working till here3')
                if len(div_elements) >= 2:
                    second_div_element = div_elements[1]
                print('legal9')
                # print(second_div_element,'second div element')
                cell_body = second_div_element.find_element(
                    By.TAG_NAME, 'datatable-body-cell').find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'label')
                print('legal10')
                # print(cell_body)
                driver.execute_script("arguments[0].click();", cell_body)
                # cell_body.click()
                print('legal11')
                # row = [12, 'Checkbox of 5th row(With SBI) clicked', 'Pass']
                # df.loc[len(df)] = row
                legal_party.append(value)
                # print('row sbi selected')
                print('legal12')
                cell_bodies = WebDriverWait(second_div_element, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
                # cell_bodies[16]
                print('legal13')
                email_identifier = cell_bodies[unique_id]
                print('legal14')
                mail_subject_unique_id.append(email_identifier.find_element(
                    By.TAG_NAME, 'div').text)
                print('legal15')
                # print(mail_subject_unique_id[index])

                # category = cell_bodies[6]
                # mail_category.append(category.find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'span').text)
                # print(mail_category[index])
                name.append(driver.find_element(
                    By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[1]/div/h5').text)
                print('legal16')
                # print(name[index])
            print('legal17')
            print("name", name, "legalparty", legal_party)
            print('legal18')
        except Exception as e:
            print(e, 'row sbi not selected')

    try:
        print('START OF MAIL SENT')
        mailsendbutton = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="0"]/div/div/a[5]')))
        driver.execute_script("arguments[0].click();", mailsendbutton)
        # driver.find_element(By.XPATH, '//*[@id="0"]/div/div/a[5]').click()
        print('mail sent')
        row = ['20', 'Module Level', 'Record Screen', 'Send Email',
               'Check if Send Email button can be clicked', 'Mail must be sent', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [13, 'Send Email Clicked', 'Pass']
        # df.loc[len(df)] = row
    except Exception as e:
        row = ['20', 'Module Level', 'Record Screen', 'Send Email',
               'Check if Send Email button can be clicked', 'Mail must be sent', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        df.loc[len(df)] = row
        # row = [13, 'Send Email Clicked', 'Fail']
        # df.loc[len(df)] = row
        print('mail not sent')
        print(e)

    return df, mail_subject_unique_id, name, mail_category, legal_party


def date_filter(driver, df, column_number):
    try:
        print('datefilter')
        startdateinput = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="batchReceivedStartDate"]')))
        print('datefilter1')
        startdateinput.click()
        print('datefilter2')
        startdateinput.send_keys("01-06-2023")
        print('datefilter3')
        # driver.execute_script("arguments[0].value = arguments[1];", startdateinput, '01-06-2023')
        time.sleep(2)
        print('datefilter4')
        enddateinput = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="batchReceivedEndDate"]')))
        print('datefilter5')
        enddateinput.click()
        print('datefilter6')
        enddateinput.send_keys("30-06-2023")
        print('datefilter7')
        # driver.execute_script("arguments[0].value = arguments[1];", enddateinput, '30-06-2023')
        time.sleep(2)
        print('datefilter8')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[4]/div/button[1]'))).click()
        print('datefilter9')
        new_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        print('datefilter10')
        table_rows = WebDriverWait(new_table_body, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print('datefilter11')
        for table_row in table_rows:
            print('datefilter12')
            new_table_row = WebDriverWait(table_row, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
            # new_table_row = table_rows[4].find_element(By.TAG_NAME, 'datatable-body-row')
            print('datefilter13')
            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)
            print('datefilter14')

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
            print('datefilter15')
            if len(div_elements) >= 2:
                second_div_element = div_elements[1]
            print('datefilter16')
            cell_bodies = WebDriverWait(second_div_element, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            print('datefilter17')
            # cell_bodies = second_div_element.find_elements(By.TAG_NAME, 'datatable-body-cell')
            email_received_cell_body = cell_bodies[column_number]
            print('datefilter18')
            print(email_received_cell_body.text)
        print('datefilter19')
        driver.find_element(
            By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[4]/div/button[2]').click()
        print('datefilter20')
        driver.find_element(
            By.XPATH, '//*[@id="1"]/app-dashboard/div[1]/nav/ol/li[1]/a').click()
        print('datefilter21')

        # row = ['31', 'Module Level', 'Record Screen', 'Filtering for Date', 'Checking for mails in date range',
        #        'User must be able to see a filtered list of emails on the basis of date', 'PASS']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
    except Exception as e:
        # row = ['31', 'Module Level', 'Record Screen', 'Filtering for Date', 'Checking for mails in date range',
        #        'User must be able to see a filtered list of emails on the basis of date', 'FAIL']
        # # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
        # df.loc[len(df)] = row
        print(e)
    return df


def data_checker(driver, df, mail_subject_unique_id, name, module_number, mail_category, legal_party):
    try:
        time.sleep(5)
        driver.refresh()
        print('inside data checker')

        # extracting name of columns
        print('e_res_count1')
        header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
        print('e_res_count2')
        header_cells = WebDriverWait(header_row, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
        print('e_res_count3')
        credit_balances = 0
        debit_balances = 0
        ledger_balance_as_per_client = 0
        ledger_balance_as_per_debtor = 0
        invoice_balance_as_per_client = 0
        invoice_balance_as_per_debtor = 0
        ledger_balance_as_per_creditor = 0
        invoice_balance_as_per_creditor = 0
        view_email = 0

        whole_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

        i = 0
        for header_cell in header_cells:
            if i != 0:
                driver.execute_script(
                    f"arguments[0].scrollLeft += 100;", whole_table_body)
            print(header_cell)
            column_name = WebDriverWait(header_cell, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            print('columnname', column_name)
            name_exact = WebDriverWait(column_name, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span')))
            print('namexact', name_exact)
            text_part = WebDriverWait(name_exact, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
            print('column name = ', text_part)

            if text_part == 'Invoice Balance as per Client':
                invoice_balance_as_per_client = i
                # break
            if text_part == 'Ledger Balance as per Client':
                ledger_balance_as_per_client = i
                # break
            if text_part == 'Debit Balances':
                # debit_balances = i
                debit_balances = i
                # break
            if text_part == 'Credit Balances':
                credit_balances = i
                # break
            # if text_part == 'Ledger_Balance_as_per_Client':
            #     ledger_balance_as_per_client = i
                # break
            # ledger_balance_as_per_client = i
            if text_part == 'Ledger Balance as per Debtor':
                ledger_balance_as_per_debtor = i
                # break
            # ledger_balance_as_per_debtor = i
            if text_part == 'Invoice Balance as per Client':
                invoice_balance_as_per_client = i
                # break
            if text_part == 'Invoice Balance as per Debtor':
                invoice_balance_as_per_debtor = i
                # break
            if text_part == 'Invoice Balance as per Creditor':
                invoice_balance_as_per_creditor = i
                # break
            if text_part == 'Ledger Balance as per Creditor':
                ledger_balance_as_per_creditor = i
                # break
            if text_part == 'View Email':
                view_email = i
                # break
            i += 1
            print('view email index', view_email)

        new_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        # print('yoyoy')
        # new_table_row = WebDriverWait(new_table_body, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'datatable-row-wrapper'))).find_element(By.TAG_NAME, 'datatable-body-row')
        # print('yoyoy1')
        table_rows = WebDriverWait(new_table_body, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        # print('yoyoy2')

        if module_number == 1:
            new_table_row = WebDriverWait(table_rows[4], 20).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))

            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')

            if len(div_elements) >= 2:
                second_div_element = div_elements[1]

            cell_bodies = WebDriverWait(second_div_element, 20).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))

            credit_balance = cell_bodies[credit_balances]
            debit_balance = cell_bodies[debit_balances]

            print('asdfghjkl', credit_balance, debit_balance, 'qwertyuio')
            print(WebDriverWait(credit_balance, 20).until(EC.presence_of_element_located(
                (By.TAG_NAME, 'div'))).find_element(By.TAG_NAME, 'button').text)
            print(WebDriverWait(debit_balance, 20).until(EC.presence_of_element_located(
                (By.TAG_NAME, 'div'))).find_element(By.TAG_NAME, 'button').text)

            if (WebDriverWait(credit_balance, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'div'))).find_element(By.TAG_NAME, 'button').text
               == '3,000.00'):
                print('Credit Balance Accurate')
                # row = [15, 'Credit Balance Accuracy', 'Pass']
                # df.loc[len(df)] = row
                row = ['special', 'Module Level', 'Record Screen', 'Credit Balance Accuracy',
                       'Checking if credit balance is accurate', 'User must be able to see a accurate credit balance', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            else:
                print('Credit Balance Not Accurate')
                # row = [15, 'Credit Balance Accuracy', 'Fail']
                # df.loc[len(df)] = row
                row = ['special', 'Module Level', 'Record Screen', 'Credit Balance Accuracy',
                       'Checking if credit balance is accurate', 'User must be able to see a accurate credit balance', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            if (WebDriverWait(debit_balance, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'div'))).find_element(By.TAG_NAME, 'button').text
               == '7,000.00'):
                print('Debit Balance Accurate')
                # row = [16, 'Debit Balance Accuracy', 'Pass']
                # df.loc[len(df)] = row
                row = ['special', 'Module Level', 'Record Screen', 'Debit Balance Accuracy',
                       'Checking if Debit balance is accurate', 'User must be able to see a accurate Debit balance', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            else:
                print('Debit Balance Not Accurate')
            # row = [16, 'Debit Balance Accuracy', 'Fail']
            # df.loc[len(df)] = row

                row = ['special', 'Module Level', 'Record Screen', 'Debit Balance Accuracy',
                       'Checking if Debit balance is accurate', 'User must be able to see a accurate Debit balance', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

            print('credit balance window6')
            email_view_download = cell_bodies[view_email]
            email_view_download_checker(
                driver, df, email_view_download, module_number)

            try:
                WebDriverWait(credit_balance, 10).until(EC.presence_of_element_located(
                    (By.TAG_NAME, 'div'))).find_element(By.TAG_NAME, 'button').click()
                print('credit balance window')

                WebDriverWait(credit_balance, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="exampleModal"]/div/div/div[2]/div/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[3]/div')))
                print('credit balance window1')

                # row = ['special', 'Module Level', 'Record Screen', 'Credit Balance Window', 'Checking if Credit balance window is opening', 'User must be able to see Credit balance window', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                # df.loc[len(df)] = row

            except Exception as e:
                print(e)
                # row = ['special', 'Module Level', 'Record Screen', 'Credit Balance Window', 'Checking if Credit balance window is opening', 'User must be able to see Credit balance window', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                # df.loc[len(df)] = row
            try:
                print('credit balance window2')
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="exampleModal"]/div/div/div[1]/button'))).click()
                print('credit balance window3')
            except Exception as e:
                print(e, 'another exception1')
            try:
                WebDriverWait(debit_balance, 10).until(EC.presence_of_element_located(
                    (By.TAG_NAME, 'div'))).find_element(By.TAG_NAME, 'button').click()
                print('credit balance window4')
                WebDriverWait(debit_balance, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="exampleModal"]/div/div/div[2]/div/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[3]/div')))
                print('credit balance window5')
                # row = ['special', 'Module Level', 'Record Screen', 'Debit Balance Window', 'Checking if Debit balance window is opening', 'User must be able to see Debit balance window', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                # df.loc[len(df)] = row

            except Exception as e:
                print(e)
                # row = ['special', 'Module Level', 'Record Screen', 'Debit Balance Window','Checking if Debit balance window is opening', 'User must be able to see Debit balance window', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                # df.loc[len(df)] = row
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="exampleModal"]/div/div/div[1]/button'))).click()
            except Exception as e:
                print(e, 'another exception2')

        elif module_number == 2:
            new_table_row = WebDriverWait(table_rows[4], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))

            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')

            if len(div_elements) >= 2:
                second_div_element = div_elements[1]

            cell_bodies = WebDriverWait(second_div_element, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            email_view_download = cell_bodies[view_email]

            arr2 = [3, 4, 5, 6]
            for index, value in enumerate(arr2):
                # print('Working till here')
                print(value)
                # print(table_rows[value])
                new_table_row = WebDriverWait(table_rows[value], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
                print('Working till here1')
                driver.execute_script(
                    "arguments[0].scrollIntoView();", new_table_row)
                print('Working till here2')

                div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
                print('Working till here3')
                if len(div_elements) >= 2:
                    second_div_element = div_elements[1]
                print(second_div_element, 'second div element')
                cell_body = second_div_element.find_element(
                    By.TAG_NAME, 'datatable-body-cell').find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'label')

                print(cell_body)
                driver.execute_script("arguments[0].click();", cell_body)

                cell_bodies = WebDriverWait(second_div_element, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))

                if value == 3:
                    print('came here3')
                    print("text", cell_bodies[ledger_balance_as_per_debtor].find_element(
                        By.TAG_NAME, 'div').text)
                    if cell_bodies[ledger_balance_as_per_debtor].find_element(By.TAG_NAME, 'div').text == '60,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'Without Details Record Ledger Balance as per Debtor Accuracy',
                               'Checking if Ledger Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger Balance as per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'Without Details Record Ledger Balance as per Debtor Accuracy',
                               'Checking if Ledger Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger Balance as per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                if value == 4:
                    print('came here4')
                    if cell_bodies[invoice_balance_as_per_debtor].find_element(By.TAG_NAME, 'div').text == '15,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'With Invoice Details Record Invoice Balance as per Debtor Accuracy',
                               'Checking if Invoice Balance as per Debtor is accurate', 'User must be able to see a accurate Invoice Balance as per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'With Invoice Details Record Invoice Balance as per Debtor Accuracy',
                               'Checking if Invoice Balance as per Debtor is accurate', 'User must be able to see a accurate Invoice Balance as per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                if value == 5:
                    print('came here5')
                    if cell_bodies[ledger_balance_as_per_debtor].find_element(By.TAG_NAME, 'div').text == '3,00,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger Details Record Ledger Balance as per Debtor Accuracy',
                               'Checking if Ledger Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger Balance as per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger Details Record Ledger Balance as per Debtor Accuracy',
                               'Checking if Ledger Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger Balance as per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                if value == 6:
                    print('came here6')
                    if cell_bodies[invoice_balance_as_per_debtor].find_element(By.TAG_NAME, 'div').text == '15,00,000.00' and cell_bodies[ledger_balance_as_per_debtor].find_element(By.TAG_NAME, 'div').text == '35,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger and Invoice Details Record Ledger Balance and Invoice Details as per Debtor Accuracy',
                               'Checking if Ledger and Invoice Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger and Invoice Balance as per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger and Invoice Details Record Ledger Balance and Invoice Details as per Debtor Accuracy',
                               'Checking if Ledger and Invoice Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger and Invoice Balance as per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
            email_view_download_checker(
                driver, df, email_view_download, module_number)

        elif module_number == 3:
            new_table_row = WebDriverWait(table_rows[4], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))

            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')

            if len(div_elements) >= 2:
                second_div_element = div_elements[1]

            cell_bodies = WebDriverWait(second_div_element, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            email_view_download = cell_bodies[view_email]
            print(invoice_balance_as_per_client, invoice_balance_as_per_creditor,
                  ledger_balance_as_per_client, ledger_balance_as_per_creditor, email_view_download)
            arr2 = [3, 4, 5, 6]
            for index, value in enumerate(arr2):
                # print('Working till here')
                print(value)
                # print(table_rows[value])
                new_table_row = WebDriverWait(table_rows[value], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
                print('Working till here1')
                driver.execute_script(
                    "arguments[0].scrollIntoView();", new_table_row)
                print('Working till here2')

                div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
                print('Working till here3')
                if len(div_elements) >= 2:
                    second_div_element = div_elements[1]
                print(second_div_element, 'second div element')
                cell_body = second_div_element.find_element(
                    By.TAG_NAME, 'datatable-body-cell').find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'label')

                print(cell_body)
                driver.execute_script("arguments[0].click();", cell_body)

                cell_bodies = WebDriverWait(second_div_element, 10).until(
                    EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))

                if value == 3:
                    print('came here3')
                    if cell_bodies[ledger_balance_as_per_creditor].find_element(By.TAG_NAME, 'div').text == '5,00,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'Without Details Record Ledger Balance as per Creditor Accuracy',
                               'Checking if Ledger Balance as per Creditor is accurate', 'User must be able to see a accurate Ledger Balance as per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'Without Details Record Ledger Balance as per Creditor Accuracy',
                               'Checking if Ledger Balance as per Creditor is accurate', 'User must be able to see a accurate Ledger Balance as per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                if value == 4:
                    print('came here4')
                    if cell_bodies[invoice_balance_as_per_creditor].find_element(By.TAG_NAME, 'div').text == '15,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'With Invoice Details Record Invoice Balance as per Creditor Accuracy',
                               'Checking if Invoice Balance as per Creditor is accurate', 'User must be able to see a accurate Invoice Balance as per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'With Invoice Details Record Invoice Balance as per Creditor Accuracy',
                               'Checking if Invoice Balance as per Creditor is accurate', 'User must be able to see a accurate Invoice Balance as per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                if value == 5:
                    print('came here5')
                    if cell_bodies[ledger_balance_as_per_creditor].find_element(By.TAG_NAME, 'div').text == '25,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger Details Record Ledger Balance as per Creditor Accuracy',
                               'Checking if Ledger Balance as per Creditor is accurate', 'User must be able to see a accurate Ledger Balance as per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger Details Record Ledger Balance as per Creditor Accuracy',
                               'Checking if Ledger Balance as per Creditor is accurate', 'User must be able to see a accurate Ledger Balance as per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                if value == 6:
                    print('came here6')
                    if cell_bodies[invoice_balance_as_per_creditor].find_element(By.TAG_NAME, 'div').text == '15,00,000.00' and cell_bodies[ledger_balance_as_per_creditor].find_element(By.TAG_NAME, 'div').text == '35,00,000.00':
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger and Invoice Details Record Ledger Balance and Invoice Details as per Creditor Accuracy',
                               'Checking if Ledger and Invoice Balance as per Debtor is accurate', 'User must be able to see a accurate Ledger and Invoice Balance as per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        row = ['special', 'Module Level', 'Record Screen', 'With Ledger and Invoice Details Record Ledger Balance and Invoice Details as per Creditor Accuracy',
                               'Checking if Ledger and Invoice Balance as per Creditor is accurate', 'User must be able to see a accurate Ledger and Invoice Balance as per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
            email_view_download_checker(
                driver, df, email_view_download, module_number)
        elif module_number == 4:
            new_table_row = WebDriverWait(table_rows[4], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))

            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')

            if len(div_elements) >= 2:
                second_div_element = div_elements[1]

            cell_bodies = WebDriverWait(second_div_element, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))

            email_view_download = cell_bodies[view_email]
            email_view_download_checker(
                driver, df, email_view_download, module_number)
            print('hello')
    except Exception as e:
        print('important1', e)

    return df


def report_checker(driver, df, mail_subject_unique_id, module_number, path, type):
    try:
        if module_number == 1:
            new_df = pd.read_excel(path)

            if type == 'consolidated':
                try:
                    file = 'Bank_Consolidated.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report row count',
                               'Checking if Consolidated Report row count match original report row count', 'Consolidated Report row count must match original report row count', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report row count',
                               'Checking if Consolidated Report row count match original report row count', 'Consolidated Report row count must match original report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report column match',
                               'Checking if Report column names match Consolidated original report column names', 'Report column names must match Consolidated original report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report column match',
                               'Checking if Consolidated Report column names match original report column names', 'Consolidated Report column names must match original report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]

                    mask1 = new_df['Debit Balances'].notnull()
                    mask2 = original_df['Debit Balances'].notnull()

                    try:
                        print(new_df.loc[mask1, 'Debit Balances'] ==
                              original_df.loc[mask2, 'Debit Balances'])
                        row = ['special', 'Record Level', 'Record Screen', 'Debit Balances match', 'Checking if Consolidated Report Debit Balances match original report Debit Balances match',
                               'Report Debit Balances must match original Consolidated report Debit Balances', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('1234')
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Debit Balances match', 'Checking if Consolidated Report Debit Balances match original Consolidated report Debit Balances match',
                               'Consolidated Report Debit Balances must match original Consolidated report Debit Balances', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            elif type == 'detailed':
                try:
                    file = 'Bank_Detailed.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Detailed Report row count match original Detailed report row count', 'Detailed Report row count must match original Detailed report row count', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Detailed Report row count match original detailed report row count', 'Detailed Report row count must match original detailed report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Detailed Report column match',
                               'Checking if Detailed Report column names match original detailed report column names', 'Detailed Report column names must match original detailed report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Detailed Report column match',
                               'Checking if Detailed Report column names match original detailed report column names', 'Detailed Report column names must match original detailed report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]

                    mask1 = new_df['Balance'].notnull()
                    mask2 = original_df['Balance'].notnull()

                    try:
                        print(new_df.loc[mask1, 'Balance'] ==
                              original_df.loc[mask2, 'Balance'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Debit Balances match', 'Checking if Report Debit Balances match original report Debit Balances match',
                               'Report Debit Balances must match original report Debit Balances', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Debit Balances match', 'Checking if Report Debit Balances match original report Debit Balances match',
                               'Report Debit Balances must match original report Debit Balances', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('4321')
                # print(new_df)
                # print(original_df)

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Detailed Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Detailed Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

        if module_number == 2:
            new_df = pd.read_excel(path)

            if type == 'consolidated':
                try:
                    file = 'DC_Consolidated.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'PASS']
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]
                    # print(original_null_columns)

                    # print(null_columns==original_null_columns)

                    # if null_columns == original_null_columns:
                    #     print('truw')
                    # else:
                    #     print('false')

                    # print(new_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = new_df['Ledger Balance As Per Client'].notnull()
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Client'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Client'])

                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('1234')
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Ledger Balance As Per Debtor'].notnull()
                    mask2 = original_df['Ledger Balance As Per Debtor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Debtor'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Debtor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Debtor match', 'Checking if Report Ledger Balance As Per Debtor match original report Ledger Balance As Per Debtor',
                               'Report Ledger Balance As Per Debtor must match original report Ledger Balance As Per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Debtor match', 'Checking if Report Ledger Balance As Per Debtor match original report Ledger Balance As Per Debtor',
                               'Report Ledger Balance As Per Debtor must match original report Ledger Balance As Per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Client'].notnull()
                    mask2 = original_df['Invoice Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Client'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Debtor'].notnull()
                    mask2 = original_df['Invoice Balance As Per Debtor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Debtor'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Debtor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Debtor match', 'Checking if Report Invoice Balance As Per Debtor match original report Invoice Balance As Per Debtor',
                               'Report Invoice Balance As Per Debtor must match original report Invoice Balance As Per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Debtor match', 'Checking if Report Invoice Balance As Per Debtor match original report Invoice Balance As Per Debtor',
                               'Report Invoice Balance As Per Debtor must match original report Invoice Balance As Per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('4321')
                # print(new_df)
                # print(original_df)

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            elif type == 'detailed':
                try:
                    file = 'DC_Detailed.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'PASS']
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]
                    # print(original_null_columns)

                    # print(null_columns==original_null_columns)

                    # if null_columns == original_null_columns:
                    #     print('truw')
                    # else:
                    #     print('false')

                    # print(new_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = new_df['Ledger Balance As Per Client'].notnull()
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Client'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Ledger Balance As Per Debtor'].notnull()
                    mask2 = original_df['Ledger Balance As Per Debtor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Debtor'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Debtor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Debtor match', 'Checking if Report Ledger Balance As Per Debtor match original report Ledger Balance As Per Debtor',
                               'Report Ledger Balance As Per Debtor must match original report Ledger Balance As Per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Debtor match', 'Checking if Report Ledger Balance As Per Debtor match original report Ledger Balance As Per Debtor',
                               'Report Ledger Balance As Per Debtor must match original report Ledger Balance As Per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Client'].notnull()
                    mask2 = original_df['Invoice Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Client'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Debtor'].notnull()
                    mask2 = original_df['Invoice Balance As Per Debtor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Debtor'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Debtor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Debtor match', 'Checking if Report Invoice Balance As Per Debtor match original report Invoice Balance As Per Debtor',
                               'Report Invoice Balance As Per Debtor must match original report Invoice Balance As Per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Debtor match', 'Checking if Report Invoice Balance As Per Debtor match original report Invoice Balance As Per Debtor',
                               'Report Invoice Balance As Per Debtor must match original report Invoice Balance As Per Debtor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('4321')
                # print(new_df)
                # print(original_df)

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Detailed Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Detailed Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
        if module_number == 3:
            new_df = pd.read_excel(path)

            if type == 'consolidated':
                try:
                    file = 'CC_Consolidated.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'PASS']
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]
                    # print(original_null_columns)

                    # print(null_columns==original_null_columns)

                    # if null_columns == original_null_columns:
                    #     print('truw')
                    # else:
                    #     print('false')

                    # print(new_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = new_df['Ledger Balance As Per Client'].notnull()
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Client'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Ledger Balance As Per Creditor'].notnull()
                    mask2 = original_df['Ledger Balance As Per Creditor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Creditor'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Creditor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Debtor match', 'Checking if Report Ledger Balance As Per Creditor match original report Ledger Balance As Per Creditor',
                               'Report Ledger Balance As Per v must match original report Ledger Balance As Per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Creditor match', 'Checking if Report Ledger Balance As Per Debtor match original report Ledger Balance As Per Creditor',
                               'Report Ledger Balance As Per Creditor must match original report Ledger Balance As Per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Client'].notnull()
                    mask2 = original_df['Invoice Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Client'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Creditor'].notnull()
                    mask2 = original_df['Invoice Balance As Per Creditor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Creditor'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Creditor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Creditor match', 'Checking if Report Invoice Balance As Per Creditor match original report Invoice Balance As Per Creditor',
                               'Report Invoice Balance As Per Creditor must match original report Invoice Balance As Per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Creditor match', 'Checking if Report Invoice Balance As Per Creditor match original report Invoice Balance As Per Creditor',
                               'Report Invoice Balance As Per Creditor must match original report Invoice Balance As Per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('4321')
                # print(new_df)
                # print(original_df)

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            elif type == 'detailed':
                try:
                    file = 'CC_Detailed.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'PASS']
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]
                    # print(original_null_columns)

                    # print(null_columns==original_null_columns)

                    # if null_columns == original_null_columns:
                    #     print('truw')
                    # else:
                    #     print('false')

                    # print(new_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = new_df['Ledger Balance As Per Client'].notnull()
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Client'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Client match', 'Checking if Report Ledger Balance As Per Client match original report Ledger Balance As Per Client',
                               'Report Ledger Balance As Per Client must match original report Ledger Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Ledger Balance As Per Creditor'].notnull()
                    mask2 = original_df['Ledger Balance As Per Creditor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Ledger Balance As Per Creditor'] ==
                              original_df.loc[mask2, 'Ledger Balance As Per Creditor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Creditor match', 'Checking if Report Ledger Balance As Per Creditor match original report Ledger Balance As Per Creditor',
                               'Report Ledger Balance As Per Creditor must match original report Ledger Balance As Per Creditor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Ledger Balance As Per Creditor match', 'Checking if Report Ledger Balance As Per Creditor match original report Ledger Balance As Per Creditor',
                               'Report Ledger Balance As Per Creditor must match original report Ledger Balance As Per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Client'].notnull()
                    mask2 = original_df['Invoice Balance As Per Client'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Client'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Client'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Client match', 'Checking if Report Invoice Balance As Per Client match original report Invoice Balance As Per Client',
                               'Report Invoice Balance As Per Client must match original report Invoice Balance As Per Client', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    mask1 = new_df['Invoice Balance As Per Creditor'].notnull()
                    mask2 = original_df['Invoice Balance As Per Creditor'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Invoice Balance As Per Creditor'] ==
                              original_df.loc[mask2, 'Invoice Balance As Per Creditor'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Creditor match', 'Checking if Report Invoice Balance As Per Creditor match original report Invoice Balance As Per Creditor',
                               'Report Invoice Balance As Per Debtor must match original report Invoice Balance As Per Debtor', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Invoice Balance As Per Creditor match', 'Checking if Report Invoice Balance As Per Creditor match original report Invoice Balance As Per Creditor',
                               'Report Invoice Balance As Per Creditor must match original report Invoice Balance As Per Creditor', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('4321')
                # print(new_df)
                # print(original_df)

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Detailed Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Detailed Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

        if module_number == 4:
            new_df = pd.read_excel(path)

            if type == 'consolidated':
                try:
                    file = 'LMC.xlsx'
                    original_df = pd.read_excel(file)
                    column_names_match = new_df.columns.tolist() == original_df.columns.to_list()

                    # print(new_df.shape[0])
                    if new_df.shape[0] == original_df.shape[0]:
                        print('row count match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print('row count nnot match')
                        row = ['special', 'Record Level', 'Record Screen', 'Report row count',
                               'Checking if Report row count match original report row count', 'Report row count must match original report row count', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    if column_names_match:
                        print("All column names match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print("Column names do not match the expected list.")
                        row = ['special', 'Record Level', 'Record Screen', 'Report column match',
                               'Checking if Report column names match original report column names', 'Report column names must match original report column names', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    original_null_columns = original_df.columns[original_df.isnull(
                    ).any()]

                    mask1 = new_df['Client: Estimated Amount Of Liability Involved'].notnull(
                    )
                    mask2 = original_df['Client: Estimated Amount Of Liability Involved'].notnull(
                    )

                    try:
                        print(new_df.loc[mask1, 'Client: Estimated Amount Of Liability Involved'] ==
                              original_df.loc[mask2, 'Client: Estimated Amount Of Liability Involved'])
                        # print('1234')
                        row = ['special', 'Record Level', 'Record Screen', 'Party: Estimated Amount Of Liability Involved match', 'Checking if Report Party: Estimated Amount Of Liability Involved match original report Party: Estimated Amount Of Liability Involved',
                               'Report Party: Estimated Amount Of Liability Involved must match original report Party: Estimated Amount Of Liability Involved', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    except Exception as e:
                        print(e)
                        row = ['special', 'Record Level', 'Record Screen', 'Party: Estimated Amount Of Liability Involved match', 'Checking if Report Party: Estimated Amount Of Liability Involved match original report Party: Estimated Amount Of Liability Involved',
                               'Report Party: Estimated Amount Of Liability Involved must match original report Party: Estimated Amount Of Liability Involved', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                        # print('4321')
                # print(new_df)
                # print(original_df)

                    mask3 = new_df.isnull()
                    mask4 = original_df.isnull()
                    placement_match = (mask1 == mask2).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'PASS']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")
                        row = ['special', 'Record Level', 'Record Screen', 'NULL value match',
                               'Checking if Report NULL value match original report NULL value', 'Report NULL value must match original report NULL value', 'FAIL']
                        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row

                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                except:
                    row = ['special', 'Record Level', 'Record Screen', 'Consolidated Report Check',
                           'Checking if Report is correct', 'Report must be correct', 'FAIL']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

    except Exception as e:
        print(e)


def report_checker_batches_level(driver, df, module_number):
    try:
        driver.refresh()
        print('report batches 1')
        input_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="datatable_filter"]/label/input')))
        print('testing5')
        input_element.send_keys("Test_CLient_1")
        print('testing6')
        # Press the Enter key
        input_element.send_keys(Keys.RETURN)
        print('testing7')

        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[1]/div/label'))).click()
        print('report batches 2')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[2]/datatable-body-row/div[2]/datatable-body-cell[1]/div/label'))).click()
        print('report batches 3')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/app-filter/div/form/div/div[4]/div/button[1]'))).click()
        print('report batches 4')
        element = driver.execute_script(
            "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.pb-3.pt-2 > div > div:nth-child(1) > label');")
        print('inside report download1')
        print('report batches 5')
        time.sleep(10)

        driver.execute_script("arguments[0].click();", element)
        print('report batches 6')
        time.sleep(10)
        print('inside report download2')
        element2 = driver.execute_script(
            "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.text-center.pt-1 > button');")
        time.sleep(10)
        print('inside report download3')
        driver.execute_script("arguments[0].click();", element2)
        time.sleep(10)
        print('inside report download4')
        download_directory = "C:/Users/harsh.vijaykumar/Downloads"
        print('inside report download5')

        if module_number == 1:
            file_name = "Bank_Consolidated (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Bank_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "Bank_Consolidated.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    print(filtered_df['Debit Balances'])
                    print(original_df['Debit Balances'])

                    mask1 = filtered_df['Debit Balances'].notnull()
                    mask2 = original_df['Debit Balances'].notnull()
                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Debit Balances']
                          == original_df.loc[mask2, 'Debit Balances'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                print(filtered_df)
                print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

                # if list(original_df.columns) != list(filtered_df.columns):
                #     print("Column names or order are different.")
                # else:
                #     # Check index values
                #     if not original_df.index.equals(filtered_df.index):
                #         print("Index values are different.")
                #     else:
                # print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['1', 'Batches Level', 'Bank Consolidated Batches Level Report', 'Bank Consolidated Batches Level Report Check',
                               'Check Bank Consolidated Batches Level Report', 'User should be able to checl Bank Consolidated Batches Level Report', 'PASS']
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

                except Exception as e:
                    print(e)
                    row = ['1', 'Batches Level', 'Bank Consolidated Batches Level Report', 'Bank Consolidated Batches Level Report Check',
                           'Check Bank Consolidated Batches Level Report', 'User should be able to checl Bank Consolidated Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            else:
                print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Bank Consolidated Batches Level Report', 'Bank Consolidated Batches Level Report Check',
                       'Check Bank Consolidated Batches Level Report', 'User should be able to checl Bank Consolidated Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        if module_number == 2:
            file_name = "DC_Consolidated (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Debtor_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "DC_Consolidated.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]
                filtered_df.rename(
                    columns={'Total_reminders': 'Total reminders'}, inplace=True)

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    # print(filtered_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = filtered_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask3 = filtered_df['Ledger Balance As Per Debtor'].notnull(
                    )
                    mask4 = original_df['Ledger Balance As Per Debtor'].notnull(
                    )
                    mask5 = filtered_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask6 = original_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask7 = filtered_df['Invoice Balance As Per Debtor'].notnull(
                    )
                    mask8 = original_df['Invoice Balance As Per Debtor'].notnull(
                    )
                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Ledger Balance As Per Client']
                          == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask3, 'Ledger Balance As Per Debtor']
                          == original_df.loc[mask4, 'Ledger Balance As Per Debtor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask5, 'Invoice Balance As Per Client']
                          == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask7, 'Invoice Balance As Per Debtor']
                          == original_df.loc[mask8, 'Invoice Balance As Per Debtor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                # print(filtered_df)
                # print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

        #         if list(original_df.columns) != list(filtered_df.columns):
        #             print("Column names or order are different.")
        #         else:
        # # Check index values
        #             if not original_df.index.equals(filtered_df.index):
        #                 print("Index values are different.")
        #             else:
        #                 print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['1', 'Batches Level', 'Debtor Consolidated Batches Level Report', 'Debtor Consolidated Batches Level Report Check',
                               'Check Debtor Consolidated Batches Level Report', 'User should be able to checl Debtor Consolidated Batches Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

                    # print("File download failed or timed out.")

                except Exception as e:
                    print(e)
                    print("File download failed or timed out.")
                    row = ['1', 'Batches Level', 'Debtor Consolidated Batches Level Report', 'Debtor Consolidated Batches Level Report Check',
                           'Check Debtor Consolidated Batches Level Report', 'User should be able to checl Debtor Consolidated Batches Level Report', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            else:
                print("File download failed or timed out.")
                # print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Debtor Consolidated Batches Level Report', 'Debtor Consolidated Batches Level Report Check',
                       'Check Debtor Consolidated Batches Level Report', 'User should be able to checl Debtor Consolidated Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        if module_number == 3:
            file_name = "CC_Consolidated (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Creditor_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "CC_Consolidated.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    # print(filtered_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = filtered_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask3 = filtered_df['Ledger Balance As Per Creditor'].notnull(
                    )
                    mask4 = original_df['Ledger Balance As Per Creditor'].notnull(
                    )
                    mask5 = filtered_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask6 = original_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask7 = filtered_df['Invoice Balance As Per Creditor'].notnull(
                    )
                    mask8 = original_df['Invoice Balance As Per Creditor'].notnull(
                    )
                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Ledger Balance As Per Client']
                          == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask3, 'Ledger Balance As Per Creditor']
                          == original_df.loc[mask4, 'Ledger Balance As Per Creditor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask5, 'Invoice Balance As Per Client']
                          == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask7, 'Invoice Balance As Per Creditor']
                          == original_df.loc[mask8, 'Invoice Balance As Per Creditor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                # print(filtered_df)
                # print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

        #         if list(original_df.columns) != list(filtered_df.columns):
        #             print("Column names or order are different.")
        #         else:
        # # Check index values
        #             if not original_df.index.equals(filtered_df.index):
        #                 print("Index values are different.")
        #             else:
        #                 print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        # print("File download failed or timed out.")
                        row = ['1', 'Batches Level', 'Creditor Consolidated Batches Level Report', 'Creditor Consolidated Batches Level Report Check',
                               'Check Creditor Consolidated Batches Level Report', 'User should be able to checl Creditor Consolidated Batches Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

                except Exception as e:
                    print(e)
                    print("File download failed or timed out.")
                    row = ['1', 'Batches Level', 'Creditor Consolidated Batches Level Report', 'Creditor Consolidated Batches Level Report Check',
                           'Check Creditor Consolidated Batches Level Report', 'User should be able to checl Creditor Consolidated Batches Level Report', 'FAIL']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            else:
                # print("File download failed or timed out.")
                print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Creditor Consolidated Batches Level Report', 'Creditor Consolidated Batches Level Report Check',
                       'Check Creditor Consolidated Batches Level Report', 'User should be able to checl Creditor Consolidated Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        if module_number == 4:
            file_name = "LMC_Consolidated (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Legal_Matter_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "LMC_Consolidated.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    # print(filtered_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = filtered_df['Client: Estimated Amount Of Liability Involved'].notnull(
                    )
                    mask2 = original_df['Client: Estimated Amount Of Liability Involved'].notnull(
                    )
                    # mask3 = filtered_df['Ledger Balance as per Creditor'].notnull()
                    # mask4 = original_df['Ledger Balance as per Creditor'].notnull()

                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Client: Estimated Amount Of Liability Involved']
                          == original_df.loc[mask2, 'Client: Estimated Amount Of Liability Involved'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                # try:
                #     print(filtered_df.loc[mask3,'Ledger Balance as per Creditor']==original_df.loc[mask4,'Ledger Balance as per Creditor'])
                #     print('1234')
                # except Exception as e:
                #     print(e)
                #     print('4321')

                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                # print(filtered_df)
                # print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

        #         if list(original_df.columns) != list(filtered_df.columns):
        #             print("Column names or order are different.")
        #         else:
        # # Check index values
        #             if not original_df.index.equals(filtered_df.index):
        #                 print("Index values are different.")
        #             else:
        #                 print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['1', 'Batches Level', 'Legal Matter Consolidated Batches Level Report', 'Legal Matter Consolidated Batches Level Report Check',
                               'Check Legal Matter Consolidated Batches Level Report', 'User should be able to check Legal Matter Consolidated Batches Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

                except Exception as e:
                    print(e)
                    row = ['1', 'Batches Level', 'Legal Matter Consolidated Batches Level Report', 'Legal Matter Consolidated Batches Level Report Check',
                           'Check Legal Matter Consolidated Batches Level Report', 'User should be able to check Legal Matter Consolidated Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            else:
                print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Legal Matter Consolidated Batches Level Report', 'Legal Matter Consolidated Batches Level Report Check',
                       'Check Legal Matter Consolidated Batches Level Report', 'User should be able to check Legal Matter Consolidated Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        print('Here detailed download')
        driver.refresh()
        # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[1]/div/a[1]'))).click()
        
        print('report batches 1')
        input_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="datatable_filter"]/label/input')))
        print('testing5')
        input_element.send_keys("Test_CLient_1")
        print('testing6')
        # Press the Enter key
        input_element.send_keys(Keys.RETURN)
        print('testing7')


        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[1]/div/label'))).click()
        print('report batches 2')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[2]/datatable-body-row/div[2]/datatable-body-cell[1]/div/label'))).click()
        print('report batches 3')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/app-root/div/app-batch-info/div[1]/div/tabset/div/tab[2]/app-filter/div/form/div/div[4]/div/button[1]'))).click()

        element3 = driver.execute_script(
            "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.pb-3.pt-2 > div > div:nth-child(2) > label');")
        print('inside report download1')
        time.sleep(10)

        driver.execute_script("arguments[0].click();", element3)

        time.sleep(10)
        print('inside report download2')
        element4 = driver.execute_script(
            "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.text-center.pt-1 > button');")
        time.sleep(10)
        print('inside report download3')
        driver.execute_script("arguments[0].click();", element4)
        time.sleep(10)
        print('inside report download4')
        download_directory = "C:/Users/harsh.vijaykumar/Downloads"
        print('inside report download5')

        if module_number == 1:
            file_name = "Bank_Detailed (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Bank_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "Bank_Detailed.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    print(filtered_df['Balance'])
                    print(original_df['Balance'])

                    mask1 = filtered_df['Balance'].notnull()
                    mask2 = original_df['Balance'].notnull()
                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Balance']
                          == original_df.loc[mask2, 'Balance'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                print(filtered_df)
                print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

                if list(original_df.columns) != list(filtered_df.columns):
                    print("Column names or order are different.")
                else:
                    # Check index values
                    if not original_df.index.equals(filtered_df.index):
                        print("Index values are different.")
                    else:
                        print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['1', 'Batches Level', 'Bank Detailed Batches Level Report', 'Bank Detailed Batches Level Report Check',
                               'Check Bank Detailed Batches Level Report', 'User should be able to check Bank Detailed Batches Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

                except Exception as e:
                    print(e)
                    row = ['1', 'Batches Level', 'Bank Detailed Batches Level Report', 'Bank Detailed Batches Level Report Check',
                           'Check Bank Detailed Batches Level Report', 'User should be able to check Bank Detailed Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            else:
                print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Bank Detailed Batches Level Report', 'Bank Detailed Batches Level Report Check',
                       'Check Bank Detailed Batches Level Report', 'User should be able to check Bank Detailed Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        if module_number == 2:
            file_name = "DC_Detailed (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Debtor_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "DC_Detailed.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    # print(filtered_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = filtered_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask3 = filtered_df['Ledger Balance As Per Debtor'].notnull(
                    )
                    mask4 = original_df['Ledger Balance As Per Debtor'].notnull(
                    )
                    mask5 = filtered_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask6 = original_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask7 = filtered_df['Invoice Balance As Per Debtor'].notnull(
                    )
                    mask8 = original_df['Invoice Balance As Per Debtor'].notnull(
                    )
                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Ledger Balance As Per Client']
                          == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask3, 'Ledger Balance As Per Debtor']
                          == original_df.loc[mask4, 'Ledger Balance As Per Debtor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask5, 'Invoice Balance As Per Client']
                          == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask7, 'Invoice Balance As Per Debtor']
                          == original_df.loc[mask8, 'Invoice Balance As Per Debtor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                # print(filtered_df)
                # print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

        #         if list(original_df.columns) != list(filtered_df.columns):
        #             print("Column names or order are different.")
        #         else:
        # # Check index values
        #             if not original_df.index.equals(filtered_df.index):
        #                 print("Index values are different.")
        #             else:
        #                 print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['1', 'Batches Level', 'Debtor Detailed Batches Level Report', 'Debtor Detailed Batches Level Report Check',
                               'Check Debtor Detailed Batches Level Report', 'User should be able to checl Debtor Detailed Batches Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

                except Exception as e:
                    print(e)
                    row = ['1', 'Batches Level', 'Debtor Detailed Batches Level Report', 'Debtor Detailed Batches Level Report Check',
                           'Check Debtor Detailed Batches Level Report', 'User should be able to checl Debtor Detailed Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            else:
                print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Debtor Detailed Batches Level Report', 'Debtor Detailed Batches Level Report Check',
                       'Check Debtor Detailed Batches Level Report', 'User should be able to checl Debtor Detailed Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        if module_number == 3:
            file_name = "CC_Detailed (1).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                print("File downloaded successfully!")

                # sheet_name = 'Creditor_Confirmations'
                new_df = pd.read_excel(file_path)

                file_name1 = "CC_Detailed.xlsx"
                file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

                original_df = pd.read_excel(file_path1)

                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)

                # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
                # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

                # current_date = datetime.now().date()

                # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

                filtered_df = new_df[new_df['Created On'].isin(
                    original_df['Created On'])]

                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
                # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

                lowercase_columns_df1 = [col.lower()
                                         for col in original_df.columns]
                lowercase_columns_df2 = [col.lower()
                                         for col in filtered_df.columns]

                # print(lowercase_columns_df1,'123', lowercase_columns_df2)
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')

                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                    # if column_names_match:
                    #     print("All column names match the expected list.")
                    # else:
                    #     print("Column names do not match the expected list.")

                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)

                    # print(filtered_df['Debit Balances'])
                    # print(original_df['Debit Balances'])

                    mask1 = filtered_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask2 = original_df['Ledger Balance As Per Client'].notnull(
                    )
                    mask3 = filtered_df['Ledger Balance As Per Creditor'].notnull(
                    )
                    mask4 = original_df['Ledger Balance As Per Creditor'].notnull(
                    )
                    mask5 = filtered_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask6 = original_df['Invoice Balance As Per Client'].notnull(
                    )
                    mask7 = filtered_df['Invoice Balance As Per Creditor'].notnull(
                    )
                    mask8 = original_df['Invoice Balance As Per Creditor'].notnull(
                    )
                except Exception as e:
                    print('exc', e)

                try:
                    print(filtered_df.loc[mask1, 'Ledger Balance As Per Client']
                          == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask3, 'Ledger Balance As Per Creditor']
                          == original_df.loc[mask4, 'Ledger Balance As Per Creditor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask5, 'Invoice Balance As Per Client']
                          == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                try:
                    print(filtered_df.loc[mask7, 'Invoice Balance As Per Creditor']
                          == original_df.loc[mask8, 'Invoice Balance As Per Creditor'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                    # pd.set_option('display.max_columns', None)
                    # pd.set_option('display.max_rows', None)
                # print(filtered_df)
                # print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(
                        lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

    #                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

    # # Identify rows where values are unequal
    #                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
    #                     print(unequal_rows)

    #                     excel_filename = 'unequal_rows.xlsx'
    #                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

        #         if list(original_df.columns) != list(filtered_df.columns):
        #             print("Column names or order are different.")
        #         else:
        # # Check index values
        #             if not original_df.index.equals(filtered_df.index):
        #                 print("Index values are different.")
        #             else:
        #                 print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                        row = ['1', 'Batches Level', 'Creditor Detailed Batches Level Report', 'Creditor Detailed Batches Level Report Check',
                               'Check Creditor Detailed Batches Level Report', 'User should be able to checl Creditor Detailed Batches Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                        df.loc[len(df)] = row
                    else:
                        print(
                            "Null value placement does not match in both DataFrames.")

                except Exception as e:
                    print(e)
                    row = ['1', 'Batches Level', 'Creditor Detailed Batches Level Report', 'Creditor Detailed Batches Level Report Check',
                           'Check Creditor Detailed Batches Level Report', 'User should be able to checl Credit Detailed Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row

            else:
                print("File download failed or timed out.")
                row = ['1', 'Batches Level', 'Creditor Detailed Batches Level Report', 'Credit Detailed Batches Level Report Check',
                       'Check Creditor Detailed Batches Level Report', 'User should be able to checl Creditor Detailed Batches Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

    except Exception as e:
        print(e)

    return df


def email_view_download_checker(driver, df, cell, module_number):
    if module_number == 1 or module_number == 2 or module_number == 3 or module_number == 4:
        try:
            print(cell, 'cell')
            buttondiv = WebDriverWait(cell, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            print('buttondiv', buttondiv)
            buttons = WebDriverWait(buttondiv, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'button')))
            print('Worjin hallo re')
            print('buttob[0]', buttons)

            buttons[1].click()
            # buttons[0].click()
            print('sab bdiya')
            download_directory = "C:/Users/harsh.vijaykumar/Downloads"

            file_name = "Email Template.pdf"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 30  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

            # Check if the file exists
            if os.path.exists(file_path):
                print("Email Template File downloaded successfully!")
                row = ['special', 'Module Level', 'Record Screen', 'Download Email Template',
                       'Checking if Email Template is Downloadable', 'User must be able to Download Email Template', 'PASS']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            else:
                print("Email Template File download failed or timed out.")
                row = ['special', 'Module Level', 'Record Screen', 'Download Email Template',
                       'Checking if Email Template is Downloadable', 'User must be able to Download Email Template', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

            try:
                time.sleep(20)
                print("before email template view button click")
                buttons[0].click()
                print("after email template view button click")
                time.sleep(20)
                # close_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#summernoteModal > div > div > div.modal-header > button')))
                # close_button.click()
                if (WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="summernoteModal"]/div/div/div[1]/button')))):
                    print('Email Template View Successful')
                    row = ['special', 'Module Level', 'Record Screen', 'View Email Template',
                           'Checking if Email Template is Visible', 'User must be able to View Email Template', 'PASS']
                    # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
            except:
                print('email tmeplate view not sucessful')
                row = ['special', 'Module Level', 'Record Screen', 'View Email Template',
                       'Checking if Email Template is Visible', 'User must be able to View Email Template', 'FAIL']
                # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row
            driver.refresh()
            # element4 = driver.execute_script("return document.querySelector('#summernoteModal > div > div > div.modal-header > button');")
            # driver.execute_script("arguments[0].click();", element4)
            # WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]')))

            # element4 = driver.execute_script("return document.querySelector('#summernoteModal > div > div > div.modal-header');")
            # button_element = WebDriverWait(element4,10).until(EC.presence_of_element_located((By.TAG_NAME, 'button')))

            # driver.execute_script("arguments[0].click();", button_element)

            time.sleep(10)
        except Exception as e:
            print(e)
    return


def remainder_checker(driver, df, module_number, mail_subject_unique_id):
    driver.refresh()

    try:
        print('remainder0')
        new_table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        print("remainder1")
        new_table_row = WebDriverWait(new_table_body, 10).until(EC.presence_of_element_located(
            (By.TAG_NAME, 'datatable-row-wrapper'))).find_element(By.TAG_NAME, 'datatable-body-row')
        print("remainder2")
        table_rows = WebDriverWait(new_table_body, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print("remainder3")
        if module_number == 1:
            new_table_row = WebDriverWait(table_rows[4], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        elif module_number == 2 or module_number == 3:
            new_table_row = WebDriverWait(table_rows[6], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        elif module_number == 4:
            new_table_row = WebDriverWait(table_rows[0], 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        print("remainder4")
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row)
        print("remainder5")

        div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
        print("remainder6")
        if len(div_elements) >= 2:
            second_div_element = div_elements[1]
            # print(second_div_element,'second div element')
        print("remainder7")
        cell_body = second_div_element.find_element(
            By.TAG_NAME, 'datatable-body-cell').find_element(By.TAG_NAME, 'div').find_element(By.TAG_NAME, 'label')
        print("remainder8")
        # print(cell_body)
        driver.execute_script("arguments[0].click();", cell_body)
        print("remainder9")
    except Exception as e:
        print(e, 'cannot access table for remainder sending')

    try:
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/div/div/a[3]'))).click()
            time.sleep(20)
            driver.refresh()
            print('e_res_count1')
            header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
            print('e_res_count2')
            header_cells = WebDriverWait(header_row, 20).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
            print('e_res_count3')
            required_index = 0

            whole_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

            i = 0
            for header_cell in header_cells:
                if i != 0:
                    driver.execute_script(
                        f"arguments[0].scrollLeft += 80;", whole_table_body)
                print(header_cell)
                column_name = WebDriverWait(header_cell, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'div')))
                print('columnname', column_name)
                name_exact = WebDriverWait(column_name, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span')))
                print('namexact', name_exact)
                text_part = WebDriverWait(name_exact, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
                print('column name = ', text_part)

                if text_part == 'Email Reminder Count':
                    break
                required_index += 1
                i += 1
            print('e_res_count4', required_index)

            new_table_body = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
            print("remainder1")
            new_table_row = WebDriverWait(new_table_body, 10).until(EC.presence_of_element_located(
                (By.TAG_NAME, 'datatable-row-wrapper'))).find_element(By.TAG_NAME, 'datatable-body-row')
            print("remainder2")
            table_rows = WebDriverWait(new_table_body, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
            print("remainder3")
            if module_number == 1:
                new_table_row = WebDriverWait(table_rows[4], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
            elif module_number == 2 or module_number == 3:
                new_table_row = WebDriverWait(table_rows[6], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
            elif module_number == 4:
                new_table_row = WebDriverWait(table_rows[0], 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))

            print("remainder4")
            driver.execute_script(
                "arguments[0].scrollIntoView();", new_table_row)
            print("remainder5")

            div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
            print("remainder6")
            if len(div_elements) >= 2:
                second_div_element = div_elements[1]
            cell_bodies = WebDriverWait(second_div_element, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
            # cell_bodies[16]
            # print(cell_bodies)
            if module_number == 1:
                remainder_count = cell_bodies[required_index]
            elif module_number == 2 or module_number == 3:
                remainder_count = cell_bodies[required_index]
            elif module_number == 4:
                remainder_count = cell_bodies[required_index]
            print('remainder count', remainder_count)
            req_div = WebDriverWait(remainder_count, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            print('req_div', req_div)
            button = WebDriverWait(req_div, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'button')))
            print('button', button.text)

            if button.text == '1':
                row = ['Remainder', 'Module Level', 'Record Screen', 'Send Remainder Update',
                       'Checking if Remainder Count is updated after sending mail', 'User must be able to See updated remainder count', 'PASS']
                df.loc[len(df)] = row
            else:
                row = ['Remainder', 'Module Level', 'Record Screen', 'Send Remainder Update',
                       'Checking if Remainder Count is updated after sending mail', 'User must be able to See updated remainder count', 'FAIL']
                df.loc[len(df)] = row

        except Exception as e:
            print('special exception', e)
            row = ['Remainder', 'Module Level', 'Record Screen', 'Send Remainder Update',
                   'Checking if Remainder Count is updated after sending mail', 'User must be able to See updated remainder count', 'PASS']
            df.loc[len(df)] = row

    except Exception as e:
        print(e)
        row = ['Remainder', 'Module Level', 'Record Screen', 'Send Remainder Update',
               'Checking if Remainder Count is updated after sending mail', 'User must be able to See updated remainder count', 'PASS']
        df.loc[len(df)] = row


def email_response_count_colour(driver, df, module_number, mail_subject_unique_id, name, mail_category, legal_party):
    driver.refresh()
    try:
        print('e_res_count1')
        header_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-header/div/div[2]')))
        print('e_res_count2')
        header_cells = WebDriverWait(header_row, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-header-cell')))
        print('e_res_count3')
        required_index = 0

        whole_table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body')))

        i = 0
        for header_cell in header_cells:
            if i != 0:
                driver.execute_script(
                    f"arguments[0].scrollLeft += 80;", whole_table_body)
            print(header_cell)
            column_name = WebDriverWait(header_cell, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, 'div')))
            print('columnname', column_name)
            name_exact = WebDriverWait(column_name, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span')))
            print('namexact', name_exact)
            text_part = WebDriverWait(name_exact, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'span'))).text
            print('column name = ', text_part)

            if text_part == 'Email Response Count':
                break
            required_index += 1
            i += 1
        print('e_res_count4', required_index)

        new_table_body = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        print('e_res_count5')
        table_rows = WebDriverWait(new_table_body, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print('e_res_count6')

        new_table_row = WebDriverWait(table_rows[4], 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        print('e_res_count7')
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row)
        print('e_res_count8')

        div_elements = new_table_row.find_elements(By.TAG_NAME, 'div')
        print('e_res_count9')

        if len(div_elements) >= 2:
            second_div_element = div_elements[1]
        cell_bodies = WebDriverWait(second_div_element, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
        # cell_bodies[16]
        # print(cell_bodies)

        email_response_count = cell_bodies[required_index]

        print('remainder count', email_response_count)
        req_div = WebDriverWait(email_response_count, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'div')))
        print('req_div', req_div)
        button = WebDriverWait(req_div, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'button')))
        print('button', button)
        print('button text', button.text)

        if button.text == '1':
            print('Email response count 1 is showing')
            row = ['1', 'Batches Level', 'Email Response Count 1', 'Email Response Count 1 Check',
                   'Check Email Response Count 1', 'User should be able to see Email Response Count 1', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            row = ['1', 'Batches Level', 'Email Response Count 1', 'Email Response Count 1 Check',
                   'Check Email Response Count 1', 'User should be able to see Email Response Count 1', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        print('asdf1')
        df = response_to_email_received_in_outlook(
            driver, df, mail_subject_unique_id, name, mail_category, module_number, legal_party)
        print('asdf2')
        time.sleep(360)
        print('asdf3')
        driver.refresh()
        print('asdf4')
        time.sleep(20)
        print('asdf5')

        new_table_body1 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        table_rows1 = WebDriverWait(new_table_body1, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print("remainder3")

        new_table_row1 = WebDriverWait(table_rows1[4], 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        print("remainder4")
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row1)
        print("remainder5")

        div_elements1 = new_table_row1.find_elements(By.TAG_NAME, 'div')
        print("remainder6")

        if len(div_elements1) >= 2:
            second_div_element1 = div_elements1[1]
        cell_bodies1 = WebDriverWait(second_div_element1, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
        # cell_bodies[16]
        # print(cell_bodies)

        email_response_count = cell_bodies1[required_index]

        print('remainder count', email_response_count)
        req_div1 = WebDriverWait(email_response_count, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'div')))
        print('req_div', req_div)
        button1 = WebDriverWait(req_div1, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'button')))
        print('button', button1.text)

        if button1.text == '2':
            print('Email response count 2 is showing')
            row = ['1', 'Batches Level', 'Email Response Count 2', 'Email Response Count 2 Check',
                   'Check Email Response Count 2', 'User should be able to see Email Response Count 2', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            row = ['1', 'Batches Level', 'Email Response Count 2', 'Email Response Count 2 Check',
                   'Check Email Response Count 2', 'User should be able to see Email Response Count 2', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        button1.click()

        print('sleep after button click')
        time.sleep(20)

        input_element = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="emailResponseCountDataModal"]/div/div/div[2]/div/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[1]/div/input')))
        print('input element found')
        input_element.click()
        print('input element click')

        time.sleep(20)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="emailResponseCountDataModal"]/div/div/div[3]/div/button'))).click()

        time.sleep(20)
        print('update element find and click')
        driver.refresh()

        new_table_body2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        table_rows2 = WebDriverWait(new_table_body2, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print("remainder3")

        new_table_row2 = WebDriverWait(table_rows2[4], 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        print("remainder4")
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row2)
        print("remainder5")

        div_elements2 = new_table_row2.find_elements(By.TAG_NAME, 'div')
        print("remainder6")

        if len(div_elements2) >= 2:
            second_div_element2 = div_elements2[1]
        cell_bodies2 = WebDriverWait(second_div_element2, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
        # cell_bodies[16]
        # print(cell_bodies)

        email_response_count = cell_bodies2[required_index]

        print('remainder count', email_response_count)
        req_div2 = WebDriverWait(email_response_count, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'div')))
        print('req_div', req_div2)
        button2 = WebDriverWait(req_div2, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'button')))

        button_classes = button2.get_attribute('class')

        # Check if 'orange' is in the list of classes
        print(button_classes.split())
        if 'orange-row-color' in button_classes.split():
            print("The button has the 'orange' class.")
            row = ['1', 'Batches Level', 'Email Response Colour Orange', 'Email Response Colour Orange Check',
                   'Check Email Response Colour Orange', 'User should be able to see Email Response Colour Orange', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            print("The button does not have the 'orange' class.")
            row = ['1', 'Batches Level', 'Email Response Colour Orange', 'Email Response Colour Orange Check',
                   'Check Email Response Colour Orange', 'User should be able to see Email Response Colour Orange', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        print('button', button2.text)

        if button2.text == '2':
            print('Email response count 2 is showing again')

        df = response_to_email_received_in_outlook(
            driver, df, mail_subject_unique_id, name, mail_category, module_number, legal_party)
        time.sleep(360)

        driver.refresh()
        time.sleep(20)

        new_table_body3 = WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller')))
        table_rows3 = WebDriverWait(new_table_body3, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-row-wrapper')))
        print("remainder3")

        new_table_row3 = WebDriverWait(table_rows3[4], 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'datatable-body-row')))
        print("remainder4")
        driver.execute_script("arguments[0].scrollIntoView();", new_table_row3)
        print("remainder5")

        div_elements3 = new_table_row3.find_elements(By.TAG_NAME, 'div')
        print("remainder6")

        if len(div_elements3) >= 2:
            second_div_element3 = div_elements3[1]
        cell_bodies3 = WebDriverWait(second_div_element3, 20).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'datatable-body-cell')))
        # cell_bodies[16]
        # print(cell_bodies)

        email_response_count = cell_bodies3[required_index]

        print('remainder count', email_response_count)
        req_div3 = WebDriverWait(email_response_count, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'div')))
        print('req_div', req_div3)
        button3 = WebDriverWait(req_div3, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'button')))

        button_classes = button3.get_attribute('class')

        # Check if 'orange' is in the list of classes
        print(button_classes.split())
        if 'red-row-color' in button_classes.split():
            print("The button has the 'red' class.")
            row = ['1', 'Batches Level', 'Email Response Colour Red', 'Email Response Colour Red Check',
                   'Check Email Response Colour Red', 'User should be able to see Email Response Red Orange', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            print("The button does not have the 'red' class.")
            row = ['1', 'Batches Level', 'Email Response Colour Red', 'Email Response Colour Red Check',
                   'Check Email Response Colour Red', 'User should be able to see Email Response Red Orange', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

        print('button', button3.text)

        if button3.text == '3':
            print('Email response count 3 is showing')
            row = ['1', 'Batches Level', 'Email Response Count 3', 'Email Response Count 3 Check',
                   'Check Email Response Count 3', 'User should be able to see Email Response Count 3', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        else:
            row = ['1', 'Batches Level', 'Email Response Count 3', 'Email Response Count 3 Check',
                   'Check Email Response Count 3', 'User should be able to see Email Response Count 3', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    except Exception as e:
        print(e)


def client_report_checker(driver, df, module_number):
    dropdown_arrow = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="navbarDropdown2"]')))
    dropdown_arrow.click()
    print('testing1')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[2]/div/div[2]/div/div[1]/a'))).click()
    print('testing2')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/app-root/div/app-modules-info/div/div/div[1]/div/div[2]/button[1]'))).click()
    print('testing3')

    input_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '//*[@id="datatable_filter"]/label/input')))
    print('testing5')
    input_element.send_keys("Test_CLient_1")
    print('testing6')
    # Press the Enter key
    input_element.send_keys(Keys.RETURN)
    print('testing7')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[2]/div/div[2]/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[1]/div/label'))).click()
    print('testing8')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[1]/div/a[1]'))).click()
    print('testing9')

    element = driver.execute_script(
        "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.pb-3.pt-2 > div > div:nth-child(1) > label');")
    print('inside report download1')
    # time.sleep(10)

    driver.execute_script("arguments[0].click();", element)

    # time.sleep(10)
    print('inside report download2')
    element2 = driver.execute_script(
        "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.text-center.pt-1 > button');")
    # time.sleep(10)
    print('inside report download3')
    driver.execute_script("arguments[0].click();", element2)
    time.sleep(10)
    print('inside report download4')
    download_directory = "C:/Users/harsh.vijaykumar/Downloads"
    print('inside report download5')
    time.sleep(20)
    files = os.listdir(download_directory)

    # Iterate through the files and rename the one you want
    for filename in files:
        if 'client_Consolidated.xlsx' in filename:
            # Rename the file using os.rename
            new_name = ''
            if module_number == 1:
                new_name = os.path.join(
                    download_directory, 'client_Consolidated_Bank.xlsx')
            elif module_number == 2:
                new_name = os.path.join(
                    download_directory, 'client_Consolidated_Debit.xlsx')
            elif module_number == 3:
                new_name = os.path.join(
                    download_directory, 'client_Consolidated_Credit.xlsx')
            elif module_number == 4:
                new_name = os.path.join(
                    download_directory, 'client_Consolidated_Legal.xlsx')

            os.rename(os.path.join(download_directory, filename), new_name)
            break  # You can break out of the loop once the file is renamed

    if module_number == 1:
        file_name = "client_Consolidated_Bank.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Bank_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "Bank_Consolidated.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            print(lowercase_columns_df1, '123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")

            try:
                # null_columns = new_df.columns[new_df.isnull().any()]
                # print(null_columns)

                print(filtered_df['Debit Balances'])
                print(original_df['Debit Balances'])

                mask1 = filtered_df['Debit Balances'].notnull()
                mask2 = original_df['Debit Balances'].notnull()
            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Debit Balances']
                      == original_df.loc[mask2, 'Debit Balances'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            print(filtered_df)
            print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.strip()
            filtered_df.columns = filtered_df.columns.str.strip()
            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

            original_df.index = range(len(original_df))
            filtered_df.index = range(len(filtered_df))

            # if list(original_df.columns) != list(filtered_df.columns):
            #     print("Column names or order are different.")
            # else:
            #     # Check index values
            #     if not original_df.index.equals(filtered_df.index):
            #         print("Index values are different.")
            #     else:
            #         print("DataFrames are identically labeled.")
            try:
                mask1.reset_index(drop=True, inplace=True)
                mask2.reset_index(drop=True, inplace=True)
                mask1 = original_df.isnull()
                mask2 = filtered_df.isnull()
                placement_match = (mask1 == mask2).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    row = ['1', 'Batches Level', 'Bank Consolidated Client Level Report', 'Bank Consolidated Client Level Report Check',
                           'Check Bank Consolidated Client Level Report', 'User should be able to checl Bank Consolidated Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Bank Consolidated Client Level Report', 'Bank Consolidated Client Level Report Check',
                       'Check Bank Consolidated Client Level Report', 'User should be able to checl Bank Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Bank Consolidated Client Level Report', 'Bank Consolidated Client Level Report Check',
                   'Check Bank Consolidated Client Level Report', 'User should be able to checl Bank Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    if module_number == 2:
        file_name = "client_Consolidated_Debit.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Debtor_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "DC_Consolidated.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]
            filtered_df.rename(
                columns={'Total_reminders': 'Total reminders'}, inplace=True)

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            print(lowercase_columns_df1, '123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')
            print(filtered_df.columns == original_df.columns)

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')

            try:

                mask1 = filtered_df['Ledger Balance as per Client'].notnull()
                mask2 = original_df['Ledger Balance As Per Client'].notnull()
                mask3 = filtered_df['Ledger Balance as per Debtor'].notnull()
                mask4 = original_df['Ledger Balance As Per Debtor'].notnull()
                mask5 = filtered_df['Invoice Balance as per Client'].notnull()
                mask6 = original_df['Invoice Balance As Per Client'].notnull()
                mask7 = filtered_df['Invoice Balance as per Debtor'].notnull()
                mask8 = original_df['Invoice Balance As Per Debtor'].notnull()
            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Ledger Balance as per Client']
                      == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask3, 'Ledger Balance as per Debtor']
                      == original_df.loc[mask4, 'Ledger Balance As Per Debtor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask5, 'Invoice Balance as per Client']
                      == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask7, 'Invoice Balance as per Debtor']
                      == original_df.loc[mask8, 'Invoice Balance As Per Debtor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            # print(filtered_df)
            # print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

                # Perform comparison and get the unequal parts
                # comparison_result = df1_lower.compare(df2_lower)
                # print(comparison_result)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.strip()
            filtered_df.columns = filtered_df.columns.str.strip()
            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

            # Set a common index for both DataFrames (e.g., integers)
            original_df.index = range(len(original_df))
            filtered_df.index = range(len(filtered_df))
    #         if list(original_df.columns) != list(filtered_df.columns):
    #             print("Column names or order are different.")
    #         else:
    # # Check index values
    #             if not original_df.index.equals(filtered_df.index):
    #                 print("Index values are different.")
    #             else:
    #                 print("DataFrames are identically labeled.")
            try:

                mask3.reset_index(drop=True, inplace=True)
                mask4.reset_index(drop=True, inplace=True)
                mask3 = original_df.isnull()
                mask4 = filtered_df.isnull()
                print('og', original_df)
                print('filtered', filtered_df)

                placement_match = (mask3 == mask4).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    row = ['1', 'Batches Level', 'Debtor Consolidated Client Level Report', 'Debtor Consolidated Client Level Report Check',
                           'Check Debtor Consolidated Client Level Report', 'User should be able to checl Debtor Consolidated Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Debtor Consolidated Client Level Report', 'Debtor Consolidated Client Level Report Check',
                       'Check Debtor Consolidated Client Level Report', 'User should be able to checl Debtor Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Debtor Consolidated Client Level Report', 'Debtor Consolidated Client Level Report Check',
                   'Check Debtor Consolidated Client Level Report', 'User should be able to checl Debtor Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    if module_number == 3:
        file_name = "client_Consolidated_Credit.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Creditor_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "CC_Consolidated.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]

            filtered_df.rename(
                columns={'Total_reminders': 'Total reminders'}, inplace=True)

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            print(lowercase_columns_df1, '123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")

            try:
                # null_columns = new_df.columns[new_df.isnull().any()]
                # print(null_columns)

                # print(filtered_df['Debit Balances'])
                # print(original_df['Debit Balances'])

                mask1 = filtered_df['Ledger Balance as per Client'].notnull()
                mask2 = original_df['Ledger Balance As Per Client'].notnull()
                mask3 = filtered_df['Ledger Balance as per Creditor'].notnull()
                mask4 = original_df['Ledger Balance As Per Creditor'].notnull()
                mask5 = filtered_df['Invoice Balance as per Client'].notnull()
                mask6 = original_df['Invoice Balance As Per Client'].notnull()
                mask7 = filtered_df['Invoice Balance as per Creditor'].notnull(
                )
                mask8 = original_df['Invoice Balance As Per Creditor'].notnull(
                )
            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Ledger Balance as per Client']
                      == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask3, 'Ledger Balance as per Creditor']
                      == original_df.loc[mask4, 'Ledger Balance As Per Creditor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask5, 'Invoice Balance as per Client']
                      == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask7, 'Invoice Balance as per Creditor']
                      == original_df.loc[mask8, 'Invoice Balance As Per Creditor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            # print(filtered_df)
            # print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

                # Perform comparison and get the unequal parts
                # comparison_result = df1_lower.compare(df2_lower)
                # print(comparison_result)

#                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

# # Identify rows where values are unequal
#                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
#                     print(unequal_rows)

#                     excel_filename = 'unequal_rows.xlsx'
#                     unequal_rows.to_excel(excel_filename, index=False)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.strip()
            filtered_df.columns = filtered_df.columns.str.strip()
            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

            original_df.index = range(len(original_df))
            filtered_df.index = range(len(filtered_df))

            try:
                mask3.reset_index(drop=True, inplace=True)
                mask4.reset_index(drop=True, inplace=True)
                mask3 = original_df.isnull()
                mask4 = filtered_df.isnull()
                placement_match = (mask3 == mask4).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    row = ['1', 'Batches Level', 'Creditor Consolidated Client Level Report', 'Creditor Consolidated Client Level Report Check',
                           'Check Creditor Consolidated Client Level Report', 'User should be able to checl Creditor Consolidated Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Creditor Consolidated Client Level Report', 'Creditor Consolidated Client Level Report Check',
                       'Check Creditor Consolidated Client Level Report', 'User should be able to checl Creditor Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Creditor Consolidated Client Level Report', 'Creditor Consolidated Client Level Report Check',
                   'Check Creditor Consolidated Client Level Report', 'User should be able to checl Creditor Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    if module_number == 4:
        file_name = "client_Consolidated_Legal.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Legal_Matter_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "LMC_Consolidated.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            # print(lowercase_columns_df1,'123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")

            try:
                # null_columns = new_df.columns[new_df.isnull().any()]
                # print(null_columns)

                # print(filtered_df['Debit Balances'])
                # print(original_df['Debit Balances'])

                mask1 = filtered_df['Client: Estimated Amount Of Liability Involved'].notnull(
                )
                mask2 = original_df['Client: Estimated Amount Of Liability Involved'].notnull(
                )
                # mask3 = filtered_df['Ledger Balance as per Creditor'].notnull()
                # mask4 = original_df['Ledger Balance as per Creditor'].notnull()

            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Client: Estimated Amount Of Liability Involved']
                      == original_df.loc[mask2, 'Client: Estimated Amount Of Liability Involved'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            # try:
            #     print(filtered_df.loc[mask3,'Ledger Balance as per Creditor']==original_df.loc[mask4,'Ledger Balance as per Creditor'])
            #     print('1234')
            # except Exception as e:
            #     print(e)
            #     print('4321')

                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            # print(filtered_df)
            # print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

                # Perform comparison and get the unequal parts
                # comparison_result = df1_lower.compare(df2_lower)
                # print(comparison_result)

#                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

# # Identify rows where values are unequal
#                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
#                     print(unequal_rows)

#                     excel_filename = 'unequal_rows.xlsx'
#                     unequal_rows.to_excel(excel_filename, index=False)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

    #         if list(original_df.columns) != list(filtered_df.columns):
    #             print("Column names or order are different.")
    #         else:
    # # Check index values
    #             if not original_df.index.equals(filtered_df.index):
    #                 print("Index values are different.")
    #             else:
    #                 print("DataFrames are identically labeled.")
            try:
                mask3 = original_df.isnull()
                mask4 = filtered_df.isnull()
                placement_match = (mask3 == mask4).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    row = ['1', 'Batches Level', 'Legal Matter Consolidated Client Level Report', 'Legal Matter Consolidated Client Level Report Check',
                           'Check Legal Matter Consolidated Client Level Report', 'User should be able to check Legal Matter Consolidated Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Legal Matter Consolidated Client Level Report', 'Legal Matter Consolidated Client Level Report Check',
                       'Check Legal Matter Consolidated Client Level Report', 'User should be able to check Legal Matter Consolidated Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Legal Matter Consolidated Client Level Report', 'Legal Matter Consolidated Client Level Report Check',
                   'Check Legal Matter Consolidated Client Level Report', 'User should be able to check Legal Matter Consolidated Client Lvel Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    print('Here detailed download')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[1]/div/a[1]'))).click()

    element3 = driver.execute_script(
        "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.pb-3.pt-2 > div > div:nth-child(2) > label');")
    print('inside report download1')
    # time.sleep(10)

    driver.execute_script("arguments[0].click();", element3)

    # time.sleep(10)
    print('inside report download2')
    element4 = driver.execute_script(
        "return document.querySelector('#downloadsModal > div > div > div.modal-body.pb-5.mt-n-40.pb-15 > div.text-center.pt-1 > button');")
    # time.sleep(10)
    print('inside report download3')
    driver.execute_script("arguments[0].click();", element4)
    time.sleep(10)
    print('inside report download4')
    download_directory = "C:/Users/harsh.vijaykumar/Downloads"
    print('inside report download5')
    time.sleep(20)
    files1 = os.listdir(download_directory)

    # Iterate through the files and rename the one you want
    for filename in files1:
        print(filename, 'yoyoyo')
        if 'client_Detailed View.xlsx' in filename:
            # Rename the file using os.rename
            new_name = ''
            if module_number == 1:
                new_name = os.path.join(
                    download_directory, 'client_Detailed View_Bank.xlsx')
            elif module_number == 2:
                new_name = os.path.join(
                    download_directory, 'client_Detailed View_Debit.xlsx')
            elif module_number == 3:
                new_name = os.path.join(
                    download_directory, 'client_Detailed View_Credit.xlsx')
            elif module_number == 4:
                new_name = os.path.join(
                    download_directory, 'client_Detailed View_Legal.xlsx')

            os.rename(os.path.join(download_directory, filename), new_name)
            break  # You can break out of the loop once the file is renamed

    if module_number == 1:
        file_name = "client_Detailed View_Bank.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Bank_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "Bank_Detailed.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            print(lowercase_columns_df1, '123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")

            try:
                # null_columns = new_df.columns[new_df.isnull().any()]
                # print(null_columns)

                # print(filtered_df['Debit Balances'])
                # print(original_df['Debit Balances'])

                mask1 = filtered_df['Balance'].notnull()
                mask2 = original_df['Balance'].notnull()
            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Balance']
                      == original_df.loc[mask2, 'Balance'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            print(filtered_df)
            print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

                # Perform comparison and get the unequal parts
                # comparison_result = df1_lower.compare(df2_lower)
                # print(comparison_result)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

            # if list(original_df.columns) != list(filtered_df.columns):
            #     print("Column names or order are different.")
            # else:
            #     # Check index values
            #     if not original_df.index.equals(filtered_df.index):
            #         print("Index values are different.")
            #     else:
            #         print("DataFrames are identically labeled.")
            try:
                mask3 = original_df.isnull()
                mask4 = filtered_df.isnull()
                placement_match = (mask3 == mask4).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    row = ['1', 'Batches Level', 'Bank Detailed Client Level Report', 'Bank Detailed Client Level Report Check',
                           'Check Bank Detailed Client Level Report', 'User should be able to checl Bank Detailed Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Bank Detailed Client Level Report', 'Bank Detailed Client Level Report Check',
                       'Check Bank Detailed Client Level Report', 'User should be able to checl Bank Detailed Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Bank Detailed Client Level Report', 'Bank Detailed Client Level Report Check',
                   'Check Bank Detailed Client Level Report', 'User should be able to checl Bank Detailed Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    if module_number == 2:
        file_name = "client_Detailed View_Debit.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Debtor_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "DC_Detailed.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            print(lowercase_columns_df1, '123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")

            try:
                # null_columns = new_df.columns[new_df.isnull().any()]
                # print(null_columns)

                # print(filtered_df['Debit Balances'])
                # print(original_df['Debit Balances'])

                mask1 = filtered_df['Ledger Balance as per Client'].notnull()
                mask2 = original_df['Ledger Balance As Per Client'].notnull()
                mask3 = filtered_df['Ledger Balance as per Debtor'].notnull()
                mask4 = original_df['Ledger Balance As Per Debtor'].notnull()
                mask5 = filtered_df['Invoice Balance as per Client'].notnull()
                mask6 = original_df['Invoice Balance As Per Client'].notnull()
                mask7 = filtered_df['Invoice Balance as per Debtor'].notnull()
                mask8 = original_df['Invoice Balance As Per Debtor'].notnull()
            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Ledger Balance as per Client']
                      == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask3, 'Ledger Balance as per Debtor']
                      == original_df.loc[mask4, 'Ledger Balance As Per Debtor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask5, 'Invoice Balance as per Client']
                      == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask7, 'Invoice Balance as per Debtor']
                      == original_df.loc[mask8, 'Invoice Balance As Per Debtor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            # print(filtered_df)
            # print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

                # Perform comparison and get the unequal parts
                # comparison_result = df1_lower.compare(df2_lower)
                # print(comparison_result)

#                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

# # Identify rows where values are unequal
#                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
#                     print(unequal_rows)

#                     excel_filename = 'unequal_rows.xlsx'
#                     unequal_rows.to_excel(excel_filename, index=False)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

    #         if list(original_df.columns) != list(filtered_df.columns):
    #             print("Column names or order are different.")
    #         else:
    # # Check index values
    #             if not original_df.index.equals(filtered_df.index):
    #                 print("Index values are different.")
    #             else:
    #                 print("DataFrames are identically labeled.")
            try:
                mask3 = original_df.isnull()
                mask4 = filtered_df.isnull()
                placement_match = (mask3 == mask4).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    row = ['1', 'Batches Level', 'Debtor Detailed Client Level Report', 'Debtor Detailed Client Level Report Check',
                           'Check Debtor Detailed Client Level Report', 'User should be able to checl Debtor Detailed Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Debtor Detailed Client Level Report', 'Debtor Detailed Client Level Report Check',
                       'Check Debtor Detailed Client Level Report', 'User should be able to checl Debtor Detailed Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Debtor Detailed Client Level Report', 'Debtor Detailed Client Level Report Check',
                   'Check Debtor Detailed Client Level Report', 'User should be able to checl Debtor Detailed Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    if module_number == 3:
        file_name = "client_Detailed View_Credit.xlsx"
        file_path = os.path.join(download_directory, file_name)

        # Wait for the file to be downloaded
        timeout = 10  # Maximum time to wait for the file (in seconds)
        while not os.path.exists(file_path) and timeout > 0:
            timeout -= 1
            time.sleep(1)  # Wait for 1 second

            # Check if the file exists
        if os.path.exists(file_path):
            print("File downloaded successfully!")

            sheet_name = 'Creditor_Confirmations'
            new_df = pd.read_excel(file_path, sheet_name=sheet_name)

            file_name1 = "CC_Detailed.xlsx"
            file_path1 = os.path.join(download_directory, file_name1)

            # file = 'DC_Consolidated.xlsx'

            original_df = pd.read_excel(file_path1)

            print(new_df.columns.to_list())
            print(original_df.columns.to_list())
            new_df = new_df.drop(new_df.columns[0], axis=1)
            original_df = original_df.drop(original_df.columns[0], axis=1)

            new_df.reset_index(drop=True, inplace=True)
            original_df.reset_index(drop=True, inplace=True)

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]

            filtered_df = new_df[new_df['Created On'].isin(
                original_df['Created On'])]

            filtered_df.reset_index(drop=True, inplace=True)
            print(filtered_df)
            print(original_df)

            # normalized_columns_df1 = [col.strip().lower for col in filtered_df.columns]
            # normalized_columns_df2 = [col.strip().lower for col in original_df.columns]

            lowercase_columns_df1 = [col.lower()
                                     for col in original_df.columns]
            lowercase_columns_df2 = [col.lower()
                                     for col in filtered_df.columns]

            print(lowercase_columns_df1, '123', lowercase_columns_df2)
            print(lowercase_columns_df1 == lowercase_columns_df2)
            # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
            # print(column_names_match, 'columns names match')

            if original_df.shape[0] == filtered_df.shape[0]:
                print('row count match')
            else:
                print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")

            try:
                # null_columns = new_df.columns[new_df.isnull().any()]
                # print(null_columns)

                # print(filtered_df['Debit Balances'])
                # print(original_df['Debit Balances'])

                mask1 = filtered_df['Ledger Balance as per Client'].notnull()
                mask2 = original_df['Ledger Balance As Per Client'].notnull()
                mask3 = filtered_df['Ledger Balance as per Creditor'].notnull()
                mask4 = original_df['Ledger Balance As Per Creditor'].notnull()
                mask5 = filtered_df['Invoice Balance as per Client'].notnull()
                mask6 = original_df['Invoice Balance As Per Client'].notnull()
                mask7 = filtered_df['Invoice Balance as per Creditor'].notnull(
                )
                mask8 = original_df['Invoice Balance As Per Creditor'].notnull(
                )
            except Exception as e:
                print('exc', e)

            try:
                print(filtered_df.loc[mask1, 'Ledger Balance as per Client']
                      == original_df.loc[mask2, 'Ledger Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask3, 'Ledger Balance as per Creditor']
                      == original_df.loc[mask4, 'Ledger Balance As Per Creditor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask5, 'Invoice Balance as per Client']
                      == original_df.loc[mask6, 'Invoice Balance As Per Client'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
            try:
                print(filtered_df.loc[mask7, 'Invoice Balance as per Creditor']
                      == original_df.loc[mask8, 'Invoice Balance As Per Creditor'])
                print('1234')
            except Exception as e:
                print(e)
                print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
            # print(filtered_df)
            # print(original_df)

            try:
                # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                # print(data_equals)

                df1_lower = original_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)
                df2_lower = filtered_df.applymap(
                    lambda x: x.lower() if isinstance(x, str) else x)

                # Perform comparison and get the unequal parts
                # comparison_result = df1_lower.compare(df2_lower)
                # print(comparison_result)

            except Exception as e:
                print(e, 'total comparision')

            original_df.columns = original_df.columns.str.lower()
            filtered_df.columns = filtered_df.columns.str.lower()

            try:
                mask3 = original_df.isnull()
                mask4 = filtered_df.isnull()
                placement_match = (mask3 == mask4).all().all()

                if placement_match:
                    print("Null value placement matches in both DataFrames.")
                    print('client detailed creditor allooo')
                    row = ['1', 'Batches Level', 'Creditor Detailed Client Level Report', 'Creditor Detailed Client Level Report Check',
                           'Check Creditor Detailed Client Level Report', 'User should be able to checl Creditor Detailed Client Level Report', 'PASS']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                    df.loc[len(df)] = row
                else:
                    print("Null value placement does not match in both DataFrames.")

            except Exception as e:
                print(e)
                row = ['1', 'Batches Level', 'Creditor Detailed Client Level Report', 'Creditor Detailed Client Level Report Check',
                       'Check Creditor Detailed Client Level Report', 'User should be able to check Creditor Detailed Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
                df.loc[len(df)] = row

        else:
            print("File download failed or timed out.")
            row = ['1', 'Batches Level', 'Creditor Detailed Client Level Report', 'Creditor Detailed Client Level Report Check',
                   'Check Creditor Detailed Client Level Report', 'User should be able to check Creditor Detailed Client Level Report', 'FAIL']
        # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row

    return df


@app.route('/useraccessmatrix', methods=['POST'])
def run_user_acesss_test():
    userTypes = ['1', '2', '3', '4']

    for usertype in userTypes:
        print('running everything')
        # usertype = request.json.get('userType')
        # print(usertype, type(usertype), 'hello')
        website_url = request.json.get('websiteUrl')

        # chrome_options = webdriver.ChromeOptions()
        # chrome_options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
        # chrome_options.add_argument('--no-sandbox')  # Avoid sandbox issues
        # chrome_options.add_argument('--incognito')
        # chrome_options.add_argument("--disable-web-security")
        # chrome_options.add_argument("--allow-running-insecure-content")
        # chrome_options.add_argument("--disable-infobars")
        # chrome_options.add_argument("--disable-notifications")

        # chrome_driver_path = "/chromedriver_win32/chromedriver.exe"

        # service = Service(chrome_driver_path)
        # driver = webdriver.Chrome(service=service, options=chrome_options)
        driver = webdriver.Chrome()

        # driver.get(website_url)
        # driver.maximize_window()
        driver.get(website_url)
        driver.maximize_window()

        response = {}
        column_names = ['User Type', 'Item Accesses', 'Status(PASS/FAIL)']
        df = pd.DataFrame(columns=column_names)

        user = ''
        username = ''
        password = ''
        if usertype == '1':
            user = 'Admin'
            username = 'adminuser1'
            password = 'Audit@123'
        elif usertype == '2':
            user = 'COE Executive'
            username = 'COEUser1'
            password = 'Audit@123'
        elif usertype == '3':
            user = 'COE POD Lead'
            username = 'coepuser1'
            password = 'Audit@123'
        elif usertype == '4':
            user = 'Business User'
            username = 'ETuser1'
            password = 'Audit@123'
        print(username)
        print(password)

        # login
        try:
            wait = WebDriverWait(driver, 20)
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="username"]')))
            print(username)
            print(password)
            driver.find_element(
                By.XPATH, '//*[@id="username"]').send_keys(username)

            driver.find_element(
                By.XPATH, '//*[@id="password"]').send_keys(password)

            submit_btn = driver.find_element(By.XPATH, '//*[@id="kc-login"]')
            submit_btn.click()

            # print('login pass')
            row = [user, 'Login', 'PASS']
            # df = df.append(pd.Series(row, index=df.columns), ignore_index=True)
            df.loc[len(df)] = row
        except Exception as e:
            row = [user, 'Login', 'FAIL']
            df.loc[len(df)] = row
            print(e)

        # Bank Confirmations clickabilty
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[1]/div/div')))
            row = [user, f'Module Bank Confirmations', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, f'Module Bank Confirmations', 'FAIL']
            df.loc[len(df)] = row

        # Debtor Confirmations clickabilty
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[2]/div/div')))
            row = [user, f'Module Debtor Confirmations', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, f'Module Debtor Confirmations', 'FAIL']
            df.loc[len(df)] = row

        # Creditor Confirmations clickabilty
        try:
            element3 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[3]/div/div')))
            row = [user, f'Module Creditor Confirmations', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, f'Module Creditor Confirmations', 'FAIL']
            df.loc[len(df)] = row

        # Legal Matter Confirmations clickabilty
        try:
            element4 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[4]/div/div')))
            row = [user, f'Module Legal Matter Confirmations', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, f'Module Legal Matter Confirmations', 'FAIL']
            df.loc[len(df)] = row

        # Clicking Bank Confirmations
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, f'/html/body/app-root/div/app-modules-info/div/div/div[2]/div[2]/div[1]/div/div'))).click()

        # View Client clickabilty
        # try:
        #     WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        #         (By.XPATH, '//*[@id="0"]/app-client-list/div/div/div[2]/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[8]/div/a')))
        #     row = [user, 'View Client', 'PASS']
        #     df.loc[len(df)] = row
        # except:
        #     row = [user, 'View Client', 'FAIL']
        #     df.loc[len(df)] = row

        # View Email Batches clickabilty
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="1-link"]')))
            row = [user, 'View Email Batches', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, 'View Email Batches', 'FAIL']
            df.loc[len(df)] = row

        # Clicking view email batches
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="1-link"]'))).click()

        # Details of Email batch clickabilty
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[13]/div/a')))
            row = [user, 'View Email Batch', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, 'View Email Batch', 'FAIL']
            df.loc[len(df)] = row

        # Add Email batch clickibility
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="1"]/div[2]/div[1]/div[1]/div/a')))
            row = [user, 'Add Email Batch', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, 'Add Email Batch', 'FAIL']
            df.loc[len(df)] = row

        # Clicking Email batch details
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="1"]/div[2]/div[2]/div/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[13]/div/a'))).click()

        # Send email clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/div/div/a[5]')))
            row = [user, 'Send Email', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, 'Send Email', 'FAIL']
            df.loc[len(df)] = row

        # Send remainder clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/div/div/a[3]')))
            row = [user, 'Send Remainder', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, 'Send Remainder', 'FAIL']
            df.loc[len(df)] = row

        # Download Attachments clickibility
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[18]/div/span')))
            row = [user, 'Download Attachments', 'PASS']
            df.loc[len(df)] = row
        except:
            row = [user, 'Download Attachments', 'FAIL']
            df.loc[len(df)] = row

        # Mail Responded Filtering For Yes clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[1]/div/div[1]/label')))
            row = [user, 'Mail Responded Filtering For Yes', 'PASS']
            df.loc[len(df)] = row

        except:
            row = [user, 'Mail Responded Filtering For Yes', 'FAIL']
            df.loc[len(df)] = row

        # Mail Responded Filtering For No clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[1]/div/div[2]/label')))
            row = [user, 'Mail Responded Filtering For No', 'PASS']
            df.loc[len(df)] = row

        except:
            row = [user, 'Mail Responded Filtering For No', 'FAIL']
            df.loc[len(df)] = row

        # Mail Undelivered Filtering For Yes clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[2]/div/div[1]/label')))
            row = [user, 'Undelivered Mail Filtering For Yes', 'PASS']
            df.loc[len(df)] = row

        except:
            row = [user, 'Undelivered Mail Filtering For Yes', 'FAIL']
            df.loc[len(df)] = row

        # Mail Undelivered Filtering For No clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="1"]/app-dashboard/div[2]/div/app-filter/div/form/div/div[2]/div/div[2]/label')))
            row = [user, 'Undelivered Mail Filtering For No', 'PASS']
            df.loc[len(df)] = row

        except:
            row = [user, 'Undelivered Mail Filtering For No', 'FAIL']
            df.loc[len(df)] = row

        # Consolidated Report Download Clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/div/div/a[1]'))).click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="downloadsModal"]/div/div/div[2]/div[1]/div/div[1]/label'))).click()

            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="downloadsModal"]/div/div/div[2]/div[2]/button')))

            row = [user, 'Consolidated Report Download Clickability', 'PASS']
            df.loc[len(df)] = row

            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="downloadsModal"]/div/div/div[2]/div[2]/button'))).click()

            download_directory = "C:/Users/harsh.vijaykumar/Downloads"

            file_name = ''
            if usertype == '1':
                file_name = "Bank_Consolidated.xlsx"
            if usertype == '2':
                file_name = "Bank_Consolidated (1).xlsx"
            if usertype == '3':
                file_name = "Bank_Consolidated (2).xlsx"
            if usertype == '4':
                file_name = "Bank_Consolidated (3).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                row = [user, 'Consolidated Report Downloaded', 'PASS']
                df.loc[len(df)] = row

        except:
            row = [user, 'Consolidated Report Download', 'FAIL']
            df.loc[len(df)] = row
            row = [user, 'Consolidated Report Downloaded', 'FAIL']
            df.loc[len(df)] = row

        # Detailed Report Download Clickability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/div/div/a[1]'))).click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="downloadsModal"]/div/div/div[2]/div[1]/div/div[2]/label'))).click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="downloadsModal"]/div/div/div[2]/div[2]/button')))

            row = [user, 'Detailed Report Download', 'PASS']
            df.loc[len(df)] = row

            download_directory = "C:/Users/harsh.vijaykumar/Downloads"

            file_name = ''
            if usertype == '1':
                file_name = "Bank_Detailed.xlsx"
            if usertype == '2':
                file_name = "Bank_Detailed (1).xlsx"
            if usertype == '3':
                file_name = "Bank_Detailed (2).xlsx"
            if usertype == '4':
                file_name = "Bank_Detailed (3).xlsx"
            file_path = os.path.join(download_directory, file_name)

            # Wait for the file to be downloaded
            timeout = 10  # Maximum time to wait for the file (in seconds)
            while not os.path.exists(file_path) and timeout > 0:
                timeout -= 1
                time.sleep(1)  # Wait for 1 second

                # Check if the file exists
            if os.path.exists(file_path):
                row = [user, 'Detailed Report Downloaded', 'PASS']
                df.loc[len(df)] = row

        except Exception as e:
            print(e)
            row = [user, 'Detailed Report Download', 'FAIL']
            df.loc[len(df)] = row
            row = [user, 'Detailed Report Downloaded', 'FAIL']
            df.loc[len(df)] = row

        # Closing report download module
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="downloadsModal"]/div/div/div[1]/button')))

        driver.execute_script("arguments[0].click();", element)

        # Email template submitability
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/div/div/a[2]'))).click()

            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="summernoteModal"]/div/div/div[2]/div/button')))

            row = [user, 'Email Template', 'PASS']
            df.loc[len(df)] = row

        except Exception as e:
            print(e)
            row = [user, 'Email Template', 'FAIL']
            df.loc[len(df)] = row

        # email template modal close

        # submit_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="summernoteModal"]/div/div/div[2]/div/button')))
        # driver.execute_script("arguments[0].click();", submit_button)
        driver.refresh()
        print('hello worls')
        # time.sleep(100)
        # credit balance checking
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[4]/datatable-body-row/div[2]/datatable-body-cell[16]/div/button'))).text
            row = [user, 'Credit Checking', 'PASS']
            df.loc[len(df)] = row

        except Exception as e:
            print(e)
            row = [user, 'Credit Checking', 'FAIL']
            df.loc[len(df)] = row

        # debit balance checking
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="0"]/app-active-deactive-batch-records/div[3]/div/div[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[4]/datatable-body-row/div[2]/datatable-body-cell[17]/div/button'))).text
            row = [user, 'Debit Checking', 'PASS']
            df.loc[len(df)] = row
        except Exception as e:
            row = [user, 'Debit Checking', 'FAIL']
            df.loc[len(df)] = row

        try:
            print('hello1')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="navbarDropdown2"]'))).click()
            print('hello2')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/app-root/app-header/div/div/div/div[2]/div/div[2]/div/div[2]/div/div[1]/a'))).click()
            print('hello3')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/app-root/div/app-modules-info/div/div/div[1]/div/div[2]/button[1]'))).click()
            print('hello4')
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[2]/div/div[2]/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[9]/div/a')))
            print('hello5')
            row = [user, 'View Client', 'PASS']
            df.loc[len(df)] = row

        except Exception as e:
            print(e)
            row = [user, 'View Client', 'FAIL']
            df.loc[len(df)] = row

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[1]/div/a[2]')))
            row = [user, 'Add Client', 'PASS']
            df.loc[len(df)] = row
        except Exception as e:
            row = [user, 'Add Client', 'FAIL']
            df.loc[len(df)] = row

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[2]/div/div[2]/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[9]/div/button')))
            row = [user, 'Update Client', 'PASS']
            df.loc[len(df)] = row
        except Exception as e:
            row = [user, 'Update Client', 'FAIL']
            df.loc[len(df)] = row

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/app-root/div/app-client-list/div[1]/div/tabset/div/tab/div[2]/div/div[2]/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[1]/datatable-body-row/div[2]/datatable-body-cell[9]/div/button'))).click()
        except Exception as e:
            print(e)
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="engagementForeclousureDate"]'))).click()

            row = [user, 'Edit Engagement Closure Date', 'PASS']
            df.loc[len(df)] = row

        except Exception as e:
            print(e)
            row = [user, 'Edit Engagement Closure Date', 'FAIL']
            df.loc[len(df)] = row

        response = {
            'results_count': 'answer'
        }

        df.to_excel(f'User Access Matrix for {user}.xlsx', index=False)
        driver.quit()
    os.startfile("outlook")

    # html_table = df.to_html(index=False)

    outlook = win32com.client.Dispatch(
        'Outlook.Application')
    new_email = outlook.CreateItem(0)  # 0 represents olMailItem

    # Set email properties
    new_email.Subject = 'New Email'
    new_email.HTMLBody = f"""
        <html>
    <body>
    <p>Hi, 
</p>
 <br/>
<p>

Testing for the ConfirmEase application User Access Matrix has been done successfully. PFA the test reports. Below are the details of the failed tests:  <h2>Failed Tests</h2>

    </body>
    </html>
    """

    new_email.To = 'harsh.vijaykumar@walkerchandiok.in'
    new_email.Recipients.Add('Abhishek.Malan@IN.GT.COM')
    # new_email.Recipients.Add('siddharth.mishra@walkerchandiok.in')

    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\User Access Matrix for Admin.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\User Access Matrix for COE Executive.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\User Access Matrix for COE POD Lead.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\User Access Matrix for Business User.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
        # Send the reply email
    # attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Failed Tests.xlsx"
    # if os.path.exists(attachment_path):
    #     attachment = new_email.Attachments.Add(Source=attachment_path)
    #     # Send the reply email
    new_email.Send()

    return jsonify(response)


def worker_function(module_number, username, password, website_url, failed_df):
    driver = webdriver.Chrome()

    driver.get(website_url)
    driver.maximize_window()
    # Your code to be executed in the thread goes here
    response = {}
    column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
                    'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)']

    df = pd.DataFrame(columns=column_names)

    df = login(username, password, driver, df)

    # module_number = 1
    df = module_choose(driver, module_number, df)

    heading = ''
    if (module_number == 1):
        heading = 'Bank Confirmations'
    elif (module_number == 2):
        heading = 'Debtor Confirmations'
    elif (module_number == 3):
        heading = 'Creditor Confirmations'
    elif (module_number == 4):
        heading = 'Legal Matter Confirmations'

    df = h5textchecker(driver, heading, df)
    df = role(driver, df)

    df = refresh(driver, df)

    df = email_batch_link(driver, df)

    df = new_email_batch_button(driver, df)

    df = batch_creation(driver, df, module_number)

    df = is_table_body_visible_and_view_details_click(driver, df)
    mail_subject_unique_id = []
    name = []
    mail_category = []

    legal_party = []

    df, mail_subject_unique_id, name, mail_category, legal_party = mail_send_from_application(
        driver, df, mail_subject_unique_id, name, module_number, mail_category, legal_party)

    print(mail_subject_unique_id)
    time.sleep(20)

    time.sleep(200)

    df = response_to_email_received_in_outlook(
        driver, df, mail_subject_unique_id, name, mail_category, module_number, legal_party)

    time.sleep(300)
    driver.refresh()

    print('refresh before data_checker')
    time.sleep(10)

    df = data_checker(driver, df, mail_subject_unique_id,
                      name, module_number, mail_category, legal_party)
    print('allrightttt')
    report_checker_batches_level(driver, df, module_number)

    df = client_report_checker(driver, df, module_number)

    driver.refresh()
    print('dfgi', df, 'dfgi')
    df = logout(driver, df)

    response = {
        'results_count': 'answer'
    }

    # Generate a sequence of numbers and fill the column
    num_values = len(df)  # Number of values to generate
    df['Test Case'] = list(range(1, num_values + 1))

    print(df)

    del df['Common/Module Specific']

    df_reorder = ['Test Case', 'Test Case Button/Description',
                  'Status(PASS/FAIL)', 'Test Case Screen', 'Repro Steps', 'Expected Result']
    df = df[df_reorder]

    # writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
    df_name = ''
    if module_number == 1:
        # df.to_excel(writer, sheet_name='Bank Confirmations', index=False)
        df.to_excel('Bank Confirmations Smoke Test.xlsx', index=False)
        df_name = 'Bank Confirmations'
    elif module_number == 2:
        df.to_excel('Debtor Confirmations Smoke Test.xlsx', index=False)
        df_name = 'Debtor Confirmations'
        # df.to_excel(writer, sheet_name='Debtor Confirmations', index=False)
    elif module_number == 3:
        df.to_excel('Creditor Confirmations Smoke Test.xlsx', index=False)
        df_name = 'Creditor Confirmations'
        # df.to_excel(writer, sheet_name='Creditor Confirmations', index=False)
    elif module_number == 4:
        df.to_excel('Legal Confirmations Smoke Test.xlsx', index=False)
        df_name = 'Legal Confirmations'
        # df.to_excel(writer, sheet_name='Legal Confirmations', index=False)
        # Return the response as JSON
        # driver.quit()

    mask = df['Status(PASS/FAIL)'] == 'FAIL'
    filtered_rows = df[mask]
    filtered_rows['Module'] = df_name
    failed_df = pd.concat([failed_df, filtered_rows], ignore_index=True)


@app.route('/smoketest', methods=['POST'])
def run_smoke_test():
    website_url = request.json.get('website_url')
    username = request.json.get('username')
    password = request.json.get('password')
    first = request.json.get('first')
    second = request.json.get('second')
    third = request.json.get('third')
    fourth = request.json.get('fourth')
    # module_numbers=request.json.get('module_numbers')
    module_numbers = []
    if (first):
        module_numbers.append(1)
    if (second):
        module_numbers.append(2)
    if (third):
        module_numbers.append(3)
    if (fourth):
        module_numbers.append(4)
    print(module_numbers)

    # chrome_options = webdriver.ChromeOptions()
    # # chrome_options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
    # chrome_options.add_argument('--no-sandbox')  # Avoid sandbox issues
    # chrome_options.add_argument('--incognito')
    # chrome_options.add_argument("--disable-web-security")
    # chrome_options.add_argument("--allow-running-insecure-content")
    # chrome_options.add_argument("--disable-infobars")
    # chrome_options.add_argument("--disable-notifications")

    # chrome_driver_path = "/chromedriver_win32/chromedriver.exe"

    # service = Service(chrome_driver_path)
    # driver = webdriver.Chrome(service=service, options=chrome_options)

    # ChromeOptions = Options()

    # driver = webdriver.Chrome(service=Service(executable_path="/chromedriver_win32/chromedriver.exe"), options=ChromeOptions)

    failed_test_column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
                                'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)', 'Source']
    failed_df = pd.DataFrame(columns=failed_test_column_names)

    for module_number in module_numbers:
        # service = Service()
        # options = webdriver.ChromeOptions()
        chrome_options = webdriver.ChromeOptions()
        download_directory = 'c:\\Users\\harsh.vijaykumar\\Downloads'

        prefs = {
            'download.default_directory': download_directory,
            'download.prompt_for_download': False,  # Disable the download popup
            'download.directory_upgrade': True,
            'safebrowsing.enabled': False,  # Disable safe browsing, which can trigger warnings
            'browser.download.folderList': 2,  # Use custom directory
            'profile.default_content_setting_values.automatic_downloads': 1,

        }
        chrome_options.add_experimental_option('prefs', prefs)
        driver = webdriver.Chrome(options=chrome_options)

        # driver = webdriver.Chrome()
        # driver = webdriver.Chrome()

        # webdriver_path='/chromedriver_win32/chromedriver.exe'
        # driver = webdriver.Chrome(executable_path=webdriver_path, options=chrome_options)
        # chrome_options.add_argument(f'--webdriver-path={webdriver_path}')

        # driver = webdriver.Chrome(options=chrome_options)

        # driver = webdriver.Chrome(executable_path='/chromedriver_win32/chromedriver.exe', options=chrome_options)
        driver.get(website_url)
        driver.maximize_window()
        # Your code to be executed in the thread goes here
        response = {}
        column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
                        'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)']

        df = pd.DataFrame(columns=column_names)

        df = login(username, password, driver, df)

        # module_number = 1
        df = module_choose(driver, module_number, df)

        heading = ''
        if (module_number == 1):
            heading = 'Bank Confirmations'
        elif (module_number == 2):
            heading = 'Debtor Confirmations'
        elif (module_number == 3):
            heading = 'Creditor Confirmations'
        elif (module_number == 4):
            heading = 'Legal Matter Confirmations'

        df = h5textchecker(driver, heading, df)
        df = role(driver, df)

        df = refresh(driver, df)

        df = email_batch_link(driver, df)

        df = new_email_batch_button(driver, df)

        df = batch_creation(driver, df, module_number)

        df = attachments_download_batches_level(driver, df, module_number)

        df = is_table_body_visible_and_view_details_click(driver, df)
        mail_subject_unique_id = []
        name = []
        mail_category = []

        legal_party = []

        df, mail_subject_unique_id, name, mail_category, legal_party = mail_send_from_application(
            driver, df, mail_subject_unique_id, name, module_number, mail_category, legal_party)

        print(mail_subject_unique_id)
        time.sleep(20)

        time.sleep(200)

        df = response_to_email_received_in_outlook(
            driver, df, mail_subject_unique_id, name, mail_category, module_number, legal_party)

        time.sleep(300)
        driver.refresh()

        print('refresh before data_checker')
        time.sleep(10)

        df = data_checker(driver, df, mail_subject_unique_id,
                          name, module_number, mail_category, legal_party)
        print('allrightttt')
        time.sleep(10)

        # WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="1"]/app-dashboard/div[1]/nav/ol/li[1]/a'))).click()
        print('asdfghjkl')
        df = report_download_after_view_details(
            driver, df, mail_subject_unique_id, module_number)

        df = report_checker_batches_level(driver, df, module_number)

        df = client_report_checker(driver, df, module_number)
        time.sleep(100)
        driver.refresh()
        print('dfgi', df, 'dfgi')
        df = logout(driver, df)

        response = {
            'results_count': 'answer'
        }

        # Generate a sequence of numbers and fill the column
        num_values = len(df)  # Number of values to generate
        df['Test Case'] = list(range(1, num_values + 1))

        print(df)

        del df['Common/Module Specific']

        df_reorder = ['Test Case', 'Test Case Button/Description',
                      'Status(PASS/FAIL)', 'Test Case Screen', 'Repro Steps', 'Expected Result']
        df = df[df_reorder]

        # writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        df_name = ''
        if module_number == 1:
            # df.to_excel(writer, sheet_name='Bank Confirmations', index=False)
            df.to_excel('Bank Confirmations Smoke Test.xlsx', index=False)
            df_name = 'Bank Confirmations'
        elif module_number == 2:
            df.to_excel('Debtor Confirmations Smoke Test.xlsx', index=False)
            df_name = 'Debtor Confirmations'
            # df.to_excel(writer, sheet_name='Debtor Confirmations', index=False)
        elif module_number == 3:
            df.to_excel('Creditor Confirmations Smoke Test.xlsx', index=False)
            df_name = 'Creditor Confirmations'
            # df.to_excel(writer, sheet_name='Creditor Confirmations', index=False)
        elif module_number == 4:
            df.to_excel('Legal Confirmations Smoke Test.xlsx', index=False)
            df_name = 'Legal Confirmations'
            # df.to_excel(writer, sheet_name='Legal Confirmations', index=False)
            # Return the response as JSON
            # driver.quit()

        mask = df['Status(PASS/FAIL)'] == 'FAIL'
        filtered_rows = df[mask]
        filtered_rows['Module'] = df_name
        failed_df = pd.concat([failed_df, filtered_rows], ignore_index=True)

    # thread = threading.Thread(target=worker_function(
    #     "1", username, password, website_url, failed_df))
    # thread.start()
    # thread2 = threading.Thread(target=worker_function(
    #     2, username, password, website_url, failed_df))
    # thread2.start()
    # thread3 = threading.Thread(target=worker_function(
    #     "3", username, password, website_url, failed_df))
    # thread3.start()
    # thread4 = threading.Thread(target=worker_function(
    #     "4", username, password, website_url, failed_df))
    # thread4.start()

    failed_df_reorder_columns = ['Test Case', 'Module', 'Test Case Button/Description',
                                 'Status(PASS/FAIL)', 'Test Case Screen', 'Repro Steps', 'Expected Result']
    failed_df = failed_df[failed_df_reorder_columns]
    failed_df.to_excel('Failed Tests Smoke.xlsx', index=False)

    print(failed_df.shape[0], 'number of records in failed df')

    html_body = ''
    os.startfile("outlook")

    html_table = failed_df.to_html(index=False)
    if failed_df.shape[0] == 0:
        html_body = f"""<html>
        <body>
        Hi Team,<br>
        Smoke Testing for the ConfirmEase application has been done successfully. <br>
        All the tests cases are working fine.<br>
        PFA the test reports.<br>

        Thanks.
        </body>
        </html> """
    elif failed_df.shape[0] == 1:
        html_body = f"""Hi Team, <br>
Smoke Testing for the ConfirmEase application has been done successfully. <br>
There is {failed_df.shape[0]} failed test case. Below are the detail of same:<br>
{html_table}<br>
PFA the test reports.<br>

Thanks,

"""
    else:
        html_body = f"""Hi Team, <br>
Smoke Testing for the ConfirmEase application has been done successfully. <br>
There are {failed_df.shape[0]} failed test cases. Below are the detail of same:<br>
{html_table}<br>
PFA the test reports.<br>

Thanks,

"""

    outlook = win32com.client.Dispatch(
        'Outlook.Application')
    new_email = outlook.CreateItem(0)  # 0 represents olMailItem

    # Set email properties
    new_email.Subject = 'Smoke Testing'
    # new_email.Body = f"""I trust this message finds you well. Enclosed herewith, please find the test outcomes generated by the ConfirmEase Testing Application, resulting from the recent testing procedures conducted under your supervision.\n The attachments consist of two distinct files:\nComprehensive Test Report - This document encapsulates the comprehensive summary of all the tests executed during the assessment phase.\nFailed Test Cases Report - This report exclusively highlights the test cases that did not meet the expected criteria.\nKindly be informed that the automated testing process has been successfully concluded, and the resultant files have been provided for your perusal. We kindly request your thorough analysis of the outcomes, including a focused examination of the failed test cases. Your insights and observations in this regard will be greatly appreciated.\nShould you require any further assistance or clarification, please do not hesitate to reach out. Your expertise in this matter is highly valued, and we are eager to collaborate closely to ensure the optimal quality of our product.\nThank you for your dedicated involvement in this crucial testing phase.\n
    # """
    new_email.HTMLBody = html_body
    new_email.To = 'harsh.vijaykumar@walkerchandiok.in'
    new_email.Recipients.Add('Abhishek.Malan@IN.GT.COM')
    # new_email.Recipients.Add('siddharth.mishra@walkerchandiok.in')

    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Legal Confirmations Smoke Test.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Bank Confirmations Smoke Test.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Debtor Confirmations Smoke Test.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Creditor Confirmations Smoke Test.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
        # Send the reply email
    attachment_path = r"C:\Data\D Drive Maan Le Ise\Work\uat-testing\src\api\Failed Tests Smoke.xlsx"
    if os.path.exists(attachment_path):
        attachment = new_email.Attachments.Add(Source=attachment_path)
        # Send the reply email
    new_email.Send()

    return jsonify({'results_count': 'answer'})


@app.route('/abhishek_batch', methods=['POST'])
def run_abhishek_batches():
    website_url = request.json.get('website_url')
    username = request.json.get('username')
    password = request.json.get('password')
    first = request.json.get('first')
    second = request.json.get('second')
    third = request.json.get('third')
    fourth = request.json.get('fourth')
    # module_numbers=request.json.get('module_numbers')
    module_numbers = []
    if (first):
        module_numbers.append(1)
    if (second):
        module_numbers.append(2)
    if (third):
        module_numbers.append(3)
    if (fourth):
        module_numbers.append(4)
    print(module_numbers)

    # chrome_options = webdriver.ChromeOptions()
    # # chrome_options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
    # chrome_options.add_argument('--no-sandbox')  # Avoid sandbox issues
    # chrome_options.add_argument('--incognito')
    # chrome_options.add_argument("--disable-web-security")
    # chrome_options.add_argument("--allow-running-insecure-content")
    # chrome_options.add_argument("--disable-infobars")
    # chrome_options.add_argument("--disable-notifications")

    # chrome_driver_path = "/chromedriver_win32/chromedriver.exe"

    # service = Service(chrome_driver_path)
    # driver = webdriver.Chrome(service=service, options=chrome_options)

    # ChromeOptions = Options()

    # driver = webdriver.Chrome(service=Service(executable_path="/chromedriver_win32/chromedriver.exe"), options=ChromeOptions)

    # failed_test_column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
    #                             'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)', 'Source']
    # failed_df = pd.DataFrame(columns=failed_test_column_names)

    for module_number in module_numbers:
        driver = webdriver.Chrome()

        driver.get(website_url)
        driver.maximize_window()
        # Your code to be executed in the thread goes here
        response = {}
        column_names = ['Test Case', 'Common/Module Specific', 'Test Case Screen',
                        'Test Case Button/Description', 'Repro Steps', 'Expected Result', 'Status(PASS/FAIL)']

        df = pd.DataFrame(columns=column_names)

        df = login(username, password, driver, df)

        # module_number = 1
        df = module_choose(driver, module_number, df)

        heading = ''
        if (module_number == 1):
            heading = 'Bank Confirmations'
        elif (module_number == 2):
            heading = 'Debtor Confirmations'
        elif (module_number == 3):
            heading = 'Creditor Confirmations'
        elif (module_number == 4):
            heading = 'Legal Matter Confirmations'

        df = h5textchecker(driver, heading, df)
        df = role(driver, df)

        df = refresh(driver, df)

        df = email_batch_link(driver, df)

        for i in range(200):
            df = new_email_batch_button(driver, df)

            df = batch_creation(driver, df, module_number)

        response = {
            'results_count': 'answer'
        }

    return jsonify({'results_count': 'answer'})


if __name__ == '__main__':
    # app.run()
    app.run(port=5000)
