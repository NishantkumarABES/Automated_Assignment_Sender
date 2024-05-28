from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time

if name == "__main__":
    contact_name = ph_number
    docx_file_path = file_path
    options  = webdriver.ChromeOptions()
    options.add_argument('--profile-directory=default')
    options.add_argument(r"user-data-dir=C:\Users\admin\AppData\Local\Google\Chrome\User Data")
    service = Service(executable_path=r"D:\User Data\Desktop\nishant\PYTHON\selenium\chromedriver.exe")
    driver = webdriver.Chrome(service=service,options=options)
    driver.get(f'https://web.whatsapp.com/send?phone={contact_name}')

    while True:
        try:
            attachment_box = driver.find_element("xpath",'//div[@title="Attach"]')
            attachment_box.click()
            time.sleep(3)
            break
        except: pass

    time.sleep(5)
    doc_button = driver.find_element("xpath",'//input[@accept="*"]')
    doc_button.send_keys(docx_file_path)
            

    time.sleep(5)
    send_button = driver.find_element("xpath",'//span[@data-icon="send"]')
    send_button.click()
    time.sleep(2)
    
    contact_name = "97711072123"
    driver.get(f'https://web.whatsapp.com/send?phone={contact_name}')
    driver.quit()


