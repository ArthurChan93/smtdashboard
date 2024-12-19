import selenium
from selenium import webdriver
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from pywinauto import application
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#
execution_count = 0
success_count = 0
failure_count = 0
email_interval = 14,400  # 每小時發送一次郵件

def send_email(execution_count, success_count, failure_count):
    sender_email = "arthurchan@ese.com.hk"
    receiver_email = "arthurchan@ese.com.hk"
    password = "s1winner**82793"

    subject = "Selenium- Autoupload-Monthly report& C66 report Streamlit"
    body = f"已執行{execution_count}次，成功{success_count}次；失敗{failure_count}次"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('arthurchan@ese.com.hk', 587)
        server.starttls()
        server.login(sender_email, password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        print(f"郵件已發送給 {receiver_email}")
    except Exception as e:
        print(f"發送郵件失敗: {e}")

def automate_task():
    global execution_count, success_count, failure_count
    execution_count += 1
    print(f"第{execution_count}次執行")

    try:
        # 設置無頭模式
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        # 開啟瀏覽器視窗
        driver = webdriver.Chrome(options=chrome_options)

        # 前往SMT monthly report網頁
        driver.get("https://github.com/ArthurChan93/smtdashboard") 

        driver.find_elements(By.CSS_SELECTOR,".HeaderMenu-link--sign-in")[0].click() # 點擊登入按鈕

        time.sleep(2)  # 等待登入頁面載入

        # 使用name屬性在LOGIN輸入
        username_input = driver.find_element(By.NAME, "login")
        username_input.send_keys("sing0017@gmail.com")

        # 使用name屬性在password輸入
        username_input = driver.find_element(By.NAME, "password")
        username_input.send_keys("chansingsing93*")

        # LOGIN GIT HUB
        driver.find_elements(By.CSS_SELECTOR,"input.btn")[0].click() 

        # 導航到SMT monthly report存儲庫並上傳文件
        driver.get('https://github.com/ArthurChan93/smtdashboard/upload/main')
        time.sleep(2)  # 等待登入頁面載入
        upload_input = driver.find_element(By.ID, 'upload-manifest-files-input')
        upload_input.send_keys('D:\\ArthurChan\\OneDrive - Electronic Scientific Engineering Ltd\\Monthly report(one drive)\\Monthly_report_for_edit.xlsm')
        time.sleep(6)

        commit_button = driver.find_element(By.CSS_SELECTOR, '.js-blob-submit')
        commit_button.click()
        time.sleep(5)

        driver.get('https://github.com/ArthurChan93/smtdashboard/upload/main')
        time.sleep(2)  # 等待登入頁面載入
        upload_input = driver.find_element(By.ID, 'upload-manifest-files-input')
        upload_input.send_keys('D:\\ArthurChan\\OneDrive - Electronic Scientific Engineering Ltd\\Monthly report(one drive)\\Machine_Import_data.xlsm')
        time.sleep(6)

        commit_button = driver.find_element(By.CSS_SELECTOR, '.js-blob-submit')
        commit_button.click()
        time.sleep(5)

        # 導航到C66 report存儲庫並上傳文件(hk stock information)
        driver.get('https://github.com/ArthurChan93/C66_Dashboard/upload/main')
        time.sleep(2)  # 等待登入頁面載入
        upload_input2 = driver.find_element(By.ID, 'upload-manifest-files-input')
        upload_input2.send_keys('D:\\ArthurChan\\OneDrive - Electronic Scientific Engineering Ltd\\C66\\HK_Stock_Summary.xlsm')
        time.sleep(6)

        commit_button = driver.find_element(By.CSS_SELECTOR, '.js-blob-submit')
        commit_button.click()
        time.sleep(5)

        # 導航到C66 report存儲庫並上傳文件(AR summary all region)
        driver.get('https://github.com/ArthurChan93/C66_Dashboard/upload/main')
        upload_input3 = driver.find_element(By.ID, 'upload-manifest-files-input')
        upload_input3.send_keys('D:\\ArthurChan\\OneDrive - Electronic Scientific Engineering Ltd\\C66 REPORT\\C66_All_AR_Summary-new_version.xlsm')
        time.sleep(6)

        commit_button = driver.find_element(By.CSS_SELECTOR, '.js-blob-submit')
        commit_button.click()
        time.sleep(5)

        urls = ['https://smtsalesdashboard.streamlit.app/Invoice_Summary','https://smtc66salesdashboard.streamlit.app/Service_Income']

        for i in range(2):
            driver.execute_script("window.open('"+ str(urls[i]) +"');")
            time.sleep(10)
            driver.switch_to.window(driver.window_handles[1]) # switch to new tab1
            time.sleep(5)
            driver.switch_to.window(driver.window_handles[1]) # switch to new tab2
            time.sleep(5)
        
        # 關閉所有瀏覽器窗口
        driver.quit()

        success_count += 1
    except Exception as e:
        print(f"執行失敗: {e}")
        failure_count += 1

# 每一段時間自動執行一次
start_time = time.time()
while True:
    automate_task()
    time.sleep(7200)  # 等待2小時
    if time.time() - start_time >= email_interval:
        send_email(execution_count, success_count, failure_count)
        start_time = time.time()

# 在VS Code啟動時自動執行此代碼
if __name__ == "__main__":
    automate_task()
