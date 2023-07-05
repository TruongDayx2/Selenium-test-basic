from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import pandas as pd
import subprocess
import math

def wait_time(item):
    while not item.is_displayed():
        time.sleep(1)

def privateAccess():
    global browser
    
    # browser.find_element(By.ID, 'details-button').click()
    # browser.find_element(By.ID,'proceed-link').click()
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-ssl-errors=yes')
    options.add_argument('--ignore-certificate-errors')
    

    browser = webdriver.Chrome(ChromeDriverManager().install(),options=options)
    # browser = webdriver.Chrome(options=options)
    # .Chrome(ChromeDriverManager().install())
    browser.get('https://10.0.18.122:9443/ProcessPortal/login.jsp')
    browser.maximize_window()

def loginPage():
    global strHtml
    with open('login.json') as f:
        d = json.load(f)
    inputUsername = browser.find_element(By.ID,'username')
    inputPassword = browser.find_element(By.ID,'password')

    inputUsername.send_keys(d['username'])
    inputPassword.send_keys(d['password'])

    browser.save_screenshot('login.png')
    strHtml = strHtml + "<tr><td>Đăng nhập</td><td><img style='height:150px;width:250px;' src='login.png' /></td></tr>"
    browser.find_element(By.CLASS_NAME,'ok').click()

def getBlackCustomer():
    print('Screen black customer')
    global strHtml
    browser.save_screenshot('menu.png')
    strHtml = strHtml + "<tr><td>Thanh điều hướng</td><td><img style='height:150px;width:250px;' src='menu.png' /></td></tr>"
    time.sleep(10)
    totalLink = browser.find_elements(By.CLASS_NAME,"ng-binding")
    for link in totalLink:
        # wait_time(link)
        try:
            if (link.get_attribute('title')== 'Quản lý khách hàng đen'):
                link.click()
                break 
        except NoSuchElementException:
            pass
    
    print('Screen black customer end')

def switchIframe():
    print('switch iframe:coach')
    iframes = browser.find_elements(By.TAG_NAME,'iframe')
    for iframe in iframes:
        try:
            if (iframe.get_attribute('title')=='Coach'):
                browser.switch_to.frame(iframe)
                break
        except NoSuchElementException:
            pass

def AddCustomer():
    print('Add customer')
    global strHtml
    df = pd.read_excel("input.xlsx")
    
    browser.save_screenshot('screenBlackCustomer.png')
    strHtml = strHtml + "<tr><td>Màn hình Black Customer</td><td><img style='height:150px;width:250px;' src='screenBlackCustomer.png' /></td></tr>"

    for index, row in df.iterrows():
        print(index)
        time.sleep(1)
        
        addBtn = browser.find_element(By.ID,'button-button-LOS_Search_Customer:Button1')
        addBtn.click()
                
                
        # wait_time(addBtn)
        
        # Fill data add customer
        cus_name = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_NAME')
        cus_name.clear()
        # if not math.isnan(row[0]):
        cus_name.send_keys(row[0])

        cus_phone = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_PHONE')
        cus_phone.clear()
        if not math.isnan(row[1]):
            cus_phone.send_keys('0'+str(row[1]))

        cus_email = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_EMAIL')
        cus_email.clear()
        # if not math.isnan(row[2]):
        cus_email.send_keys(row[2])

        browser.find_element(By.ID,'singleselect-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:RETRICTION_TYPE').click()
        if (row[3] =='Không cấp tín dụng'):
            browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:RETRICTION_TYPE[0]').click()
        elif (row[3] == 'Hạn chế cấp tín dụng'):
            browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:RETRICTION_TYPE[1]').click()
        elif (row[3] == 'Khách hàng đen'):
            browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:RETRICTION_TYPE[2]').click()

        cus_cif = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_CIF')
        cus_cif.clear()
        # if not math.isnan(row[4]):
        cus_cif.send_keys(row[4])

        cus_id = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_ID')
        cus_id.clear()
        if not math.isnan(row[5]):
            cus_id.send_keys(row[5])

        cus_bir = browser.find_element(By.ID,'datetimepicker-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_BIRTH')
        cus_bir.clear()
        # if not math.isnan(row[6]):
        cus_bir.send_keys(row[6])

        browser.find_element(By.ID,'singleselect-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TARGET_GROUP').click()
        if (row[7]=='Không'):
            browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TARGET_GROUP[0]').click()
        elif (row[7]=='Thành viên HĐQT, thành viên BKS, TGĐ, Phó TGĐ của BIDV'):
            browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TARGET_GROUP[1]').click()
        elif (row[7]=='Cha, mẹ, vợ/chồng, con của thành viên HĐQT, thành viên BKS, TGĐ, Phó TGĐ của BIDV'):
            browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TARGET_GROUP[2]').click()

        cus_ad = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_PERMANENT_ADDRESS')
        cus_ad.clear()
        # if not math.isnan(row[8]):
        cus_ad.send_keys(row[8])

        cus_a = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_CURRENT_ADDRESS')
        cus_a.clear()
        # if not math.isnan(row[9]):
        cus_a.send_keys(row[9])

        cus_a1 = browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TITLE_NAME')
        cus_a1.clear()
        # if not math.isnan(row[10]):
        cus_a1.send_keys(row[10])

        cus_aN = browser.find_element(By.ID,'textarea-textarea-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_NOTE')
        cus_aN.clear()
        # if not math.isnan(row[11]):
        cus_aN.send_keys(row[11])

        
        browser.save_screenshot('popUpAdd.png')
        strHtml = strHtml + "<tr><td>Màn hình điền thông tin khách hàng</td><td><img style='height:150px;width:250px;' src='popUpAdd.png' /></td></tr>"

        browser.find_element(By.ID,'button-button-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:SAVE').click()
        time.sleep(2)
        close = browser.find_element(By.ID,'button-button-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CLOSE')
        
        if close.is_displayed():
            browser.save_screenshot('pictureErr/popUpAdd'+row[4]+'.png')
            print('close',close)
            close.click()
        else:
            print('123')
        print(str(index) + 'good')

    # //

    time.sleep(2)
    browser.save_screenshot('success.png')
    strHtml = strHtml + "<tr><td>Thêm thành công</td><td><img style='height:150px;width:250px;' src='success.png' /></td></tr>"


def report():
    global strHtml
    report = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Report</title>
    </head>
    <body>
        <table>  
            <tr><th>Step</th><th>Img</th></tr>  
            {strHtml} 
        </table>
    </body>
    </html>
    """
    with open("report2.html", "w", encoding="utf-8") as file:
        file.write(report)

if __name__== "__main__":
    privateAccess()
    strHtml=""
    print('Login page')
    loginPage()
    time.sleep(20)
    
    getBlackCustomer()
    time.sleep(15)
    
    switchIframe()
    
    AddCustomer()
    time.sleep(5)
    # report()
    browser.close()
    # Export data
    subprocess.call(["python", "oracle.py"])

    print("Successful")
    