from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
import pandas as pd
from time import sleep
from datetime import datetime


# prepare data
data_blast =  pd.read_excel("data.xlsx", sheet_name=0)

# define path driver
driver = webdriver.Chrome(executable_path='./chromedriver.exe')

# open whatsapp web in web browser
driver.get("https://web.whatsapp.com/")

# scan barcode
input('Press Enter to continue after scan qr code ')

log_col = ['Number','Status']
log_row = []
max_try = 20

# iterate data to send message by phone number
for index, row in data_blast.iterrows():
    phone_number = str(row['Number'])
    message = row['Text']
    print(phone_number)

    driver.get("""https://web.whatsapp.com/send?phone={phone_number}""".format(phone_number=phone_number))

    loop_check_wa = True
    invalid_phone_number = False
    curr_try = 0
    err = False

    sleep(2)

    while(loop_check_wa) :
        curr_try = curr_try + 1
        # print('Check number ({})'.format(curr_try))
        try :
            wait = WebDriverWait(driver, 2)
            status = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@dir="ltr"][@spellcheck="true"]')))
            loop_check_wa = False
        except TimeoutException as to:
            if (curr_try != max_try) :
                try :
                    driver.find_element(By.XPATH, '//div[text()="Phone number shared via url is invalid."]')
                    invalid_phone_number = True
                    loop_check_wa = False
                except NoSuchElementException as ne:
                    loop_check_wa = True
            else :
                loop_check_wa = False
                err = True
                print('Whatsapp Web did not open')
                log_row.append([phone_number, 'Whatsapp Web did not open'])

    if (invalid_phone_number) :
        log_row.append([phone_number, 'Not a Phone Number'])
        continue

    if (err) :
        driver.close()
        continue

    sleep(2)
    if (message != '') :
        loop_find_chat = True
        curr_try = 0

        while(loop_find_chat) :
            curr_try = curr_try + 1
            # print('Check chat box ({})'.format(curr_try))
            try :
                wait = WebDriverWait(driver, 2)
                status = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@dir="ltr"][@spellcheck="true"]')))
                loop_find_chat = False
            except TimeoutException as to:
                if (curr_try != max_try) :
                    loop_find_chat = True
                else :
                    err = True
                    loop_find_chat = False
                    print('Failed to find chat box')
                    log_row.append([phone_number, 'Failed to find chat box'])

        if (err) :
            driver.close()
            continue

        chat_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@dir="ltr"][@spellcheck="true"]')
        chat_box.clear()
        for t in message.split('\n') :
            chat_box.send_keys(t)
            chat_box.send_keys(Keys.SHIFT, Keys.ENTER)

        loop_find_send = True
        curr_try = 0

        while(loop_find_send) :
            curr_try = curr_try + 1
            # print('Check send button ({})'.format(curr_try))
            try :
                wait = WebDriverWait(driver, 2)
                status = wait.until(EC.element_to_be_clickable((By.XPATH, '//button/span[@data-icon="send"]')))
                loop_find_send = False
            except TimeoutException as to:
                if (curr_try != max_try) :
                    loop_find_send = True
                else :
                    err = True
                    loop_find_send = False
                    print('Failed to find send button')
                    log_row.append([phone_number, 'Failed to fing send button'])

        if (err) :
            driver.close()
            continue

        send_button = driver.find_element_by_xpath('//button/span[@data-icon="send"]')
        send_button.click()
        sleep(2)

    print('Message has been sent successfully')
    log_row.append([phone_number, 'Message has been sent successfully'])
    
    
local_dt = datetime.now()
nama_out = "Log-{}-{:02d}-{:02d} {:02d}-{:02d}-{:02d}.xlsx".format(local_dt.year, local_dt.month, local_dt.day, local_dt.hour, local_dt.minute, local_dt.second)
pd_out = pd.DataFrame(log_row, columns=log_col)
pd_out.to_excel(nama_out, index = False)