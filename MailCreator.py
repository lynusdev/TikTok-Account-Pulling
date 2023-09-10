from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import warnings
from random import randint, choice
import sys
import os
from anycaptcha import AnycaptchaClient
from anycaptcha import FunCaptchaProxylessTask
from time import sleep, time
from datetime import timedelta
from concurrent.futures import ThreadPoolExecutor
from threading import current_thread
from os import system
from threading import Thread
from colorama import Fore, init

system(f"title {os.path.basename(__file__)[:-3]}")
warnings.filterwarnings("ignore")
init(convert=True)

CERROR = Fore.RED
CWARNING = Fore.YELLOW
CSUCCESS = Fore.LIGHTGREEN_EX
CNEUTRAL = Fore.WHITE
CTHREAD = Fore.CYAN

timeout = 30
completed = 0
created = 0
unavailable = 0
failed = 0
max_workers = 5

proxies = [ "IP:PORT",
            "IP:PORT",
            "IP:PORT" ]

with open("./data/names.txt", "r") as f:
    names = f.read().splitlines()

def stats():
    while queue == True:
        system(f"title {os.path.basename(__file__)[:-3]} - Tasks: {completed}/{len(cleaned_combo)} ^| Created: {created} ^| Unavailable: {unavailable} ^| Failed: {failed} ^| Time: {str(timedelta(seconds=(time() - start_time))).split('.')[0]}")
        sleep(0.01)

def solve_captcha(url):
    thread_number = int(current_thread().name.split('_')[1])+1
    if thread_number < 10:
        thread_number = "0" + str(thread_number)
    while True:
        anycaptcha_api_key = "API_KEY"
        site_key = "B7D8911C-5CC8-A9A3-35B0-554ACEE604DA"
        client = AnycaptchaClient(anycaptcha_api_key)
        task = FunCaptchaProxylessTask(url, site_key)

        job = client.createTask(task,typecaptcha="funcaptcha")
        job.join()
        result = job.get_solution_response()
        if result.find("ERROR") != -1:
            print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't get captcha: [{result}], retrying...")
        else:
            return result
        
def retry_login(login):
    global created
    global failed
    global completed
    thread_number = int(current_thread().name.split('_')[1])+1
    if thread_number < 10:
        thread_number = "0" + str(thread_number)
    email_address = login.split(":")[0]
    password = login.split(":")[1]
    with open(f"{combo_name}-failed.txt", "a+") as file:
        file.write(f'{email_address}:{password}\n')
    failed += 1
    completed += 1
    print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't verify email creation: [{email_address}], please try manually logging in...")
    return

def create_mail(account):
    global created
    global unavailable
    global failed
    global completed
    thread_number = int(current_thread().name.split('_')[1])+1
    if thread_number < 10:
        thread_number = "0" + str(thread_number)
    email_address = account.split(":")[2]
    password = f"{randint(100000, 999999)}A!"
    last_proxy = ""
    tried_to_create = False

    print(f"{CTHREAD}[Thread {thread_number}] {CNEUTRAL}Creating email {email_address}")

    options = webdriver.ChromeOptions()
    options.add_argument("--lang=en")
    options.add_argument("--mute-audio")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-site-isolation-trials")
    options.add_argument("--disable-application-cache")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-login-animations")
    options.add_argument("--disable-notifications")
    options.add_argument("--incognito")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)

    while True:
        try:
            while True:
                while True:
                    proxy = choice(proxies)
                    if proxy != last_proxy:
                        last_proxy = proxy
                        break
                options.add_argument(f"--proxy-server={proxy}")
                driver = webdriver.Chrome("./data/chromedriver.exe", options=options)
                driver.set_page_load_timeout(20)
                driver.set_window_size(384, 780)
                driver.set_window_position(0, 7)
                try:
                    driver.get("https://outlook.live.com/owa/?nlp=1&signup=1")
                    runner = 0
                    while True:
                        try:
                            driver.find_element(By. ID, "MemberName")
                            break
                        except:
                            if "This site can’t be reached" in driver.page_source:
                                print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Proxy not working, rotating proxy...")
                                raise ValueError('Proxy not working!')
                            elif "The request is blocked." in driver.page_source:
                                print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Proxy is blocked, rotating proxy...")
                                raise ValueError('Proxy is blocked!')
                            elif "There's a temporary problem" in driver.page_source:
                                print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Microsoft temporary problem, rotating proxy...")
                                raise ValueError('Proxy is blocked!')
                            elif runner > 15:
                                raise ValueError('Proxy not working!')
                            runner += 1
                            sleep(1)
                    WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "MemberName"))).send_keys(email_address.split("@")[0])
                    Select(driver.find_element(By.ID, "LiveDomainBoxList")).select_by_value(email_address.split("@")[1])
                    WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "iSignupAction"))).click()
                    break
                except:
                    try:
                        driver.close()
                    except:
                        pass

            counter = 0
            while True:
                try:
                    driver.find_element(By. ID, "PasswordInput").send_keys(password)
                    break
                except:
                    if "Someone already has this email address." in driver.page_source:
                        if tried_to_create == True:
                            login = [email_address, password]
                            Pool.submit(retry_login, login)
                            print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Email {email_address} not available after trying to create, trying to log in later and moving on...")
                            return
                        else:
                            unavailable += 1
                            completed += 1
                            print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Email {email_address} not available, moving on...")
                            return
                    elif counter > 60:
                        raise ValueError("Error while trying to check for availability")
                    counter += 1
                    sleep(0.5)
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "iSignupAction"))).click()
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "FirstName"))).send_keys(choice(names))
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "LastName"))).send_keys(choice(names))
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "iSignupAction"))).click()
            country_ids = ['AF', 'AX', 'AL', 'DZ', 'AS', 'AD', 'AO', 'AI', 'AQ', 'AG', 'AR', 'AM', 'AW', 'AC', 'AU', 'AT', 'AZ', 'BS', 'BH', 'BD', 'BB', 'BY', 'BE', 'BZ', 'BJ', 'BM', 'BT', 'BO', 'BQ', 'BA', 'BW', 'BV', 'BR', 'IO', 'VG', 'BN', 'BG', 'BF', 'BI', 'CV', 'KH', 'CM', 'CA', 'KY', 'CF', 'TD', 'CL', 'CN', 'CX', 'CC', 'CO', 'KM', 'CG', 'CD', 'CK', 'CR', 'CI', 'HR', 'CU', 'CW', 'CY', 'CZ', 'DK', 'DJ', 'DM', 'DO', 'EC', 'EG', 'SV', 'GQ', 'ER', 'EE', 'ET', 'FK', 'FO', 'FJ', 'FI', 'FR', 'GF', 'PF', 'TF', 'GA', 'GM', 'GE', 'DE', 'GH', 'GI', 'GR', 'GL', 'GD', 'GP', 'GU', 'GT', 'GG', 'GN', 'GW', 'GY', 'HT', 'HM', 'HN', 'HK', 'HU', 'IS', 'IN', 'ID', 'IR', 'IQ', 'IE', 'IM', 'IL', 'IT', 'JM', 'XJ', 'JP', 'JE', 'JO', 'KZ', 'KE', 'KI', 'KR', 'XK', 'KW', 'KG', 'LA', 'LV', 'LB', 'LS', 'LR', 'LY', 'LI', 'LT', 'LU', 'MO', 'MG', 'MW', 'MY', 'MV', 'ML', 'MT', 'MH', 'MQ', 'MR', 'MU', 'YT', 'MX', 'FM', 'MD', 'MC', 'MN', 'ME', 'MS', 'MA', 'MZ', 'MM', 'NA', 'NR', 'NP', 'NL', 'AN', 'NC', 'NZ', 'NI', 'NE', 'NG', 'NU', 'NF', 'MK', 'MP', 'NO', 'OM', 'PK', 'PW', 'PS', 'PA', 'PG', 'PY', 'PE', 'PH', 'PN', 'PL', 'PT', 'PR', 'QA', 'RE', 'RO', 'RU', 'RW', 'XS', 'BL', 'KN', 'LC', 'MF', 'PM', 'VC', 'WS', 'SM', 'ST', 'SA', 'SN', 'RS', 'SC', 'SL', 'SG', 'XE', 'SX', 'SK', 'SI', 'SB', 'SO', 'ZA', 'GS', 'SS', 'ES', 'LK', 'SH', 'SD', 'SR', 'SJ', 'SZ', 'SE', 'CH', 'SY', 'TW', 'TJ', 'TZ', 'TH', 'TL', 'TG', 'TK', 'TO', 'TT', 'TA', 'TN', 'TR', 'TM', 'TC', 'TV', 'UM', 'VI', 'UG', 'UA', 'AE', 'UK', 'US', 'UY', 'UZ', 'VU', 'VA', 'VE', 'VN', 'WF', 'YE', 'ZM', 'ZW']
            Select(WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "Country")))).select_by_value(choice(country_ids))
            Select(WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "BirthMonth")))).select_by_value(str(randint(1, 12)))
            Select(WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "BirthDay")))).select_by_value(str(randint(1, 28)))
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "BirthYear"))).send_keys(str(randint(1970, 2002)))
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "iSignupAction"))).click()

            captcha_sent = False
            captcha_stuck = False
            runner = 0
            while True:
                try:
                    WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "enforcementFrame")))
                    enforcement_iframe = driver.find_element(By. ID, "enforcementFrame")
                    url = enforcement_iframe.get_attribute("src")
                    solution = solve_captcha(url)
                    driver.execute_script("""
                        var anyCaptchaToken = '"""+solution+"""';
                        var enc = document.getElementById('enforcementFrame');
                        var encWin = enc.contentWindow || enc;
                        var encDoc = enc.contentDocument || encWin.document;
                        let script = encDoc.createElement('SCRIPT');
                        script.append('function AnyCaptchaSubmit(token) { parent.postMessage(JSON.stringify({ eventId: "challenge-complete", payload: { sessionToken: token } }), "*") }');
                        encDoc.documentElement.appendChild(script);
                        encWin.AnyCaptchaSubmit(anyCaptchaToken);
                    """)
                    tried_to_create = True
                    print(f"{CTHREAD}[Thread {thread_number}] {CNEUTRAL}Sent captcha!")
                    runner = 0
                    while True:
                        try:
                            WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.ID, "idBtn_Back")))
                            captcha_sent = True
                            break
                        except:
                            if "Someone already has this email address." in driver.page_source:
                                unavailable += 1
                                completed += 1
                                print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Email {email_address} not available, moving on...")
                                return
                            else:
                                if "https://login.live.com/ppsecure/post.srf" in driver.current_url:
                                    if runner > 10:
                                        print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Authentication redirect didn't work, logging in...")
                                        driver.get("https://outlook.live.com/owa/?nlp=1")
                                        runner = 0
                                        while True:
                                            if "Mail" in driver.title:
                                                sleep(2)
                                                with open(f"{combo_name}-created.txt", "a+") as file:
                                                    file.write(f'{email_address}:{password}\n')
                                                created += 1
                                                completed += 1
                                                print(f"{CTHREAD}[Thread {thread_number}] {CSUCCESS}[SUCCESS] Created email: [{email_address}]!")
                                                try:
                                                    driver.close()
                                                except:
                                                    pass
                                                return
                                            elif runner == 10 or runner == 30 or runner == 50 or runner == 70:
                                                driver.get("https://outlook.live.com/mail/0/")
                                            elif runner == 20 or runner == 40 or runner == 60:
                                                driver.get("https://outlook.live.com/owa/?nlp=1")
                                            elif runner > 80:
                                                with open(f"{combo_name}-failed.txt", "a+") as file:
                                                    file.write(f'{email_address}:{password}\n')
                                                failed += 1
                                                completed += 1
                                                print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't fully create email: [{email_address}] and load UI, moving on...")
                                                return
                                            runner += 1
                                            sleep(1)
                                    runner += 1
                                    sleep(1)
                                elif runner == 20 or runner == 40:
                                    print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Captcha didn't work, retrying...")
                                    break
                                elif runner >= 60:
                                    print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Captcha stuck, retrying...")
                                    captcha_stuck = True
                                    break
                                runner += 1
                                sleep(1)

                    if captcha_sent == True:
                        break
                except:
                    if "Phone number" in driver.page_source:
                        print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Phone verification needed, rotating ip...")
                        driver.close()
                        break
                    elif runner > 20:
                        break
                    runner += 1
                    sleep(1)

            if captcha_stuck == True:
                continue
            elif captcha_sent == True:
                break
        except Exception as e:
            print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Unexpected error [{e}], retrying...")
            try:
                driver.close()
            except:
                pass
        
    runner = 0
    while True:
        try:
            driver.find_element(By. ID, "idBtn_Back")
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "idBtn_Back"))).click()
            break
        except:
            if "https://login.live.com/ppsecure/post.srf" in driver.current_url:
                if runner > 10:
                    print(f"{CTHREAD}[Thread {thread_number}] {CWARNING}[WARNING] Authentication redirect didn't work, logging in...")
                    driver.get("https://outlook.live.com/owa/?nlp=1")
                    break
                runner += 1
                sleep(1)
    runner = 0
    while True:
        if "Mail" in driver.title:
            sleep(2)
            with open(f"{combo_name}-created.txt", "a+") as file:
                file.write(f'{email_address}:{password}\n')
            created += 1
            completed += 1
            print(f"{CTHREAD}[Thread {thread_number}] {CSUCCESS}[SUCCESS] Created email: [{email_address}]!")
            return
        elif runner == 10 or runner == 30 or runner == 50 or runner == 70:
            driver.get("https://outlook.live.com/mail/0/")
        elif runner == 20 or runner == 40 or runner == 60:
            driver.get("https://outlook.live.com/owa/?nlp=1")
        elif runner > 80:
            with open(f"{combo_name}-failed.txt", "a+") as file:
                file.write(f'{email_address}:{password}\n')
            failed += 1
            completed += 1
            print(f"{CTHREAD}[Thread {thread_number}] {CERROR}[ERROR] Couldn't fully create email: [{email_address}] and load UI, moving on...")
            return
        runner += 1
        sleep(1)

combo_name = input("Enter name of checked combo: ")
try:
    with open(f"{combo_name}.txt", "r") as f:
        combo = f.read().splitlines()
    cleaned_combo = []
    for i, line in enumerate(combo):
        try:
            email_address = line.split(":")[2]
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
for i, account in enumerate(cleaned_combo):
    Pool.submit(create_mail, account)
    if i < max_workers:
        sleep(0.1)
Pool.shutdown(wait=True)
sleep(0.1)
queue = False
input(f"{CSUCCESS}Successfully completed all tasks, press ENTER to close...")