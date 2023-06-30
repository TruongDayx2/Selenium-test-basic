from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
import json


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
    strHtml = strHtml + "<tr><td>Đăng nhập</td><td><img style='height:150px;width:250px;' src='pictureMain1/login.png' /></td></tr>"
    browser.find_element(By.CLASS_NAME,'ok').click()

def getBlackCustomer():
    global strHtml
    browser.save_screenshot('menu.png')
    strHtml = strHtml + "<tr><td>Thanh điều hướng</td><td><img style='height:150px;width:250px;' src='pictureMain1/menu.png' /></td></tr>"
    totalLink = browser.find_elements(By.CLASS_NAME,"ng-binding")
    for link in totalLink:
        try:
            if (link.get_attribute('title')== 'Quản lý khách hàng đen'):
                link.click()
                break 
        except NoSuchElementException:
            pass

def switchIframe():

    iframes = browser.find_elements(By.TAG_NAME,'iframe')
    for iframe in iframes:
        try:
            if (iframe.get_attribute('title')=='Coach'):
                browser.switch_to.frame(iframe)
                break
        except NoSuchElementException:
            pass

def AddCustomer():
    global strHtml
    
    browser.save_screenshot('screenBlackCustomer.png')
    strHtml = strHtml + "<tr><td>Màn hình Black Customer</td><td><img style='height:150px;width:250px;' src='pictureMain1/screenBlackCustomer.png' /></td></tr>"

    browser.find_element(By.ID,'button-button-LOS_Search_Customer:Button1').click()
    # Fill data add customer
    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_NAME').send_keys('Truong')
    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_PHONE').send_keys('0396134431')
    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_EMAIL').send_keys('truong1@gmail.com')
    browser.find_element(By.ID,'singleselect-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:RETRICTION_TYPE').click()
    browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:RETRICTION_TYPE[0]').click()

    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_CIF').send_keys('orcl1')
    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_ID').send_keys('123456789')
    browser.find_element(By.ID,'datetimepicker-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_BIRTH').send_keys('06/06/2010')
    browser.find_element(By.ID,'singleselect-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TARGET_GROUP').click()
    browser.find_element(By.ID,'singleselect-option-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TARGET_GROUP[0]').click()

    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_PERMANENT_ADDRESS').send_keys('Thu Duc')
    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_CURRENT_ADDRESS').send_keys('Thu Duc')
    browser.find_element(By.ID,'text-input-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_TITLE_NAME').send_keys('Tu do')
    browser.find_element(By.ID,'textarea-textarea-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:CUS_NOTE').send_keys('Truong test selenium')
    # //
    browser.save_screenshot('popUpAdd.png')
    strHtml = strHtml + "<tr><td>Màn hình điền thông tin khách hàng</td><td><img style='height:150px;width:250px;' src='pictureMain1/popUpAdd.png' /></td></tr>"

    browser.find_element(By.ID,'button-button-LOS_Search_Table_Result_Customer:LOS_POPUP_RESTRICTED_CUST1:SAVE').click()
    time.sleep(2)
    browser.save_screenshot('success.png')
    strHtml = strHtml + "<tr><td>Thêm thành công</td><td><img style='height:150px;width:250px;' src='pictureMain1/success.png' /></td></tr>"


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
    with open("report.html", "w", encoding="utf-8") as file:
        file.write(report)

if __name__== "__main__":
    privateAccess()
    strHtml=""
    print('Login page')
    loginPage()
    time.sleep(20)
    print('Screen black customer')
    getBlackCustomer()
    time.sleep(10)
    print('switch iframe:coach')
    switchIframe()
    print('Add customer')
    AddCustomer()
    time.sleep(5)
    report()
    browser.close()
    