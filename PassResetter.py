import os
from os import system, chdir
import sys
import warnings
from ppadb.client import Client
from colorama import Fore, init
from time import sleep, time
from datetime import timedelta
from threading import Thread
import imaplib
import email

system(f"title {os.path.basename(__file__)[:-3]}")
warnings.filterwarnings("ignore")
init(convert=True)
original_path = os.getcwd()

apk = "15.0.1.apk"
package_name = "com.zhiliaoapp.musically"
proxy = "83.149.70.159:13082"

CERROR = Fore.RED
CWARNING = Fore.YELLOW
CSUCCESS = Fore.LIGHTGREEN_EX
CTHREAD = Fore.CYAN
CNEUTRAL = Fore.WHITE

completed = 0
reset = 0
failed = 0
suspended = 0

def stats():
    task_length = len(cleaned_combo)
    while queue == True:
        system(f"title {os.path.basename(__file__)[:-3]} - Tasks: {completed}/{task_length} ^| Reset: {reset} ^| Failed: {failed} ^| Suspended: {suspended} ^| Time: {str(timedelta(seconds=(time() - start_time))).split('.')[0]}")
        sleep(0.01)

def getPixelColor(x, y):
    offset=1080*y+x+4
    cmd = f"dd if='/sdcard/screen.dump' bs=4 count=1 skip={offset} 2>/dev/null | xxd -p"
    device.shell("screencap /sdcard/screen.dump")
    out = device.shell(cmd)
    return str(out).strip()[:-2]

def waitForPixel(x, y, hex, timeout):
    start_time = time()
    while True:
        if getPixelColor(x, y) == hex:
            return True
        elif (time() - start_time) > timeout:
            return False
        sleep(0.001)

def pressInstallButton():
    while True:
        if getPixelColor(885, 1993) == "0073dd":
            device.shell("input tap 338 2024")
            break
        sleep(0.01)

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
    print(f"{CSUCCESS}[SUCCESS] Read combo, correct lines: [{len(cleaned_combo)}/{len(combo)}]")
except:
    input(f"{CERROR}[ERROR] Couldn't read combo, press ENTER to close...")
    sys.exit()

queue = True
start_time = time()
Thread(target=stats).start()

print(f"{CNEUTRAL}Starting adb...")
chdir("data\\adb")
system("adb kill-server >null 2>&1")
system("adb devices >null 2>&1")
adb = Client(host='127.0.0.1', port=5037)
chdir(original_path)

devices = adb.devices()
if not len(devices) < 1:
    device = devices[0]
else:
    input(f"{CERROR}[ERROR] No devices found, press ENTER to close...")
    sys.exit()

device.shell(f"adb shell settings put global http_proxy {proxy}")

if device.is_installed(package_name) == False:
    print(f"{CNEUTRAL}Installing apk...")
    Thread(target=pressInstallButton).start()
    device.install(f"./data/{apk}")

for i, account in enumerate(cleaned_combo):
    email_address = account.split(":")[0]
    password = account.split(":")[1]
    reset_start_time = time()

    print(f"{CNEUTRAL}Resetting password of {email_address}")

    code_sent = False
    while True:
        if code_sent == True:
            print(f"{CWARNING}[WARNING] Error after code was sent, retrying later...")
            cleaned_combo.append(account)
            code_sent = False
            break
        code_sent = False

        device.shell(f"pm clear {package_name}")
        device.shell(f"monkey -p {package_name} -c android.intent.category.LAUNCHER 1")

        logged_in = False
        for i in range(5):
            try:
                imap = imaplib.IMAP4_SSL("outlook.office365.com")
                imap.login(email_address, password)
            except:
                print(f"{CWARNING}[WARNING] Couldn't log in [{i+1}], retrying...")
                sleep(0.5)
            else:
                print(f"{CNEUTRAL}IMAP logged in!")
                logged_in = True
                break
        if logged_in == False:
            print(f"{CERROR}[ERROR] IMAP couldn't log in {email_address}, saving and moving on...")
            with open(f"{combo_name}-failed.txt", "a+") as file:
                file.write(f'{email_address}:{password}\n')
            failed += 1
            completed += 1
            break
        else:
            selecting_worked = False
            for i in range(10):
                try:
                    imap = imaplib.IMAP4_SSL("outlook.office365.com")
                    imap.login(email_address, password)
                    imap.select("Inbox")
                    _, msgnums = imap.search(None, 'SUBJECT "is your verification code"')
                    if msgnums[0].split():
                        for msgnum in msgnums[0].split():
                            imap.store(msgnum, "+FLAGS", "\\Deleted")
                        imap.expunge()
                        print(f"{CNEUTRAL}Deleting {len(msgnums[0].split())} old 2FA codes from inbox...")
                    imap.select("Junk")
                    _, msgnums = imap.search(None, 'SUBJECT "is your verification code"')
                    if msgnums[0].split():
                        for msgnum in msgnums[0].split():
                            imap.store(msgnum, "+FLAGS", "\\Deleted")
                        imap.expunge()
                        print(f"{CNEUTRAL}Deleting {len(msgnums[0].split())} old 2FA codes from junk...")
                except:
                    print(f"{CWARNING}[WARNING] IMAP couldn't select at try [{i+1}], retrying...")
                    sleep(2)
                else:
                    selecting_worked = True
                    break
            if selecting_worked == False:
                print(f"{CERROR}[ERROR] IMAP couldn't select {email_address}, saving and moving on...")
                with open(f"{combo_name}-failed.txt", "a+") as file:
                    file.write(f'{email_address}:{password}\n')
                failed += 1
                completed += 1
                break

        while True:
            if getPixelColor(1008, 135) == "8a8b91":
                device.shell("input tap 1008 135")
                waitForPixel(515, 2036, "fe2c55", 10)
                device.shell("input tap 515 2036")
                waitForPixel(971, 2120, "bfbfbf", 10)
                break
            elif getPixelColor(971, 2120) == "bfbfbf":
                break

        device.shell("input tap 971 2120")
        if waitForPixel(550, 1120, "fe2c55", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 1, retrying...")
            continue
        device.shell("input tap 550 1120")
        if waitForPixel(729, 2126, "fe2c55", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 2, retrying...")
            continue
        device.shell("input tap 729 2126")
        if waitForPixel(792, 802, "161823", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 3, retrying...")
            continue
        device.shell("input tap 792 802")
        if waitForPixel(490, 780, "04498d", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 4, retrying...")
            continue
        device.shell("input tap 490 780")
        if waitForPixel(930, 1360, "e2e3e3", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 5, retrying...")
            continue
        device.shell(f"input text {email_address}")
        sleep(0.1)
        device.shell("input tap 930 1360")
        code_sent = True
        if waitForPixel(930, 1360, "e2e3e3", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 6, retrying...")
            continue

        code_received = False
        for i in range(50):
            if i % 2 == 0:
                imap.select("Inbox")
            else:
                imap.select("Junk")
            _, msgnums = imap.search(None, 'SUBJECT "is your verification code"')
            if msgnums[0].split():
                _, data = imap.fetch(msgnums[0].split()[-1], "(RFC822)")
                message = email.message_from_bytes(data[0][1]) 
                code = message.get('Subject').split(" is your verification code")[0]
                code_received = True
                break
            else:
                sleep(0.5)
        if code_received == False:
            print(f"{CWARNING}[WARNING] Code receive timout, retrying later...")
            cleaned_combo.append(account)
            break
        else:
            print(f"{CNEUTRAL}2FA Code: {code}")

        device.shell(f"input text {code}")
        # Wait for code confirm button to be clickable
        if waitForPixel(920, 1350, "fe2d56", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 8, retrying...")
            continue
        device.shell("input tap 930 1360")
        # Wait for password confirm button to be gray so input can be made
        if waitForPixel(920, 1350, "e2e3e3", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 7, retrying...")
            continue
        device.shell(f"input text {password}")
        # Wait for password confirm button to be clickable
        if waitForPixel(920, 1350, "fe2d56", 10) == False:
            print(f"{CWARNING}[WARNING] Wait for pixel timeout 8, retrying...")
            continue

        stay_start_time = time()
        counter_start_time = 5

        while True:
            if (time() - counter_start_time) >= 5:
                device.shell("input tap 910 1350")
                counter_start_time = time()
            if (time() - stay_start_time) >= 30:
                print(f"{CWARNING}[WARNING] Password reset timeout, retrying later...")
                cleaned_combo.append(account)
                break
            # Check if account is suspended
            if getPixelColor(1000, 270) == "fe695a":
                print(f"{CERROR}[ERROR] {email_address} is currently suspended, saving and moving on...")
                suspended += 1
                completed += 1
                break
            # Check if successfully reset password by checking account icon
            account_icon = getPixelColor(970, 2120)
            if account_icon == "bfbfbf" or account_icon == "5f5f5f":
                with open(f"{combo_name}-reset.txt", "a+") as file:
                    file.write(f'{email_address}:{password}\n')
                print(f"{CSUCCESS}[SUCCESS] Successfully reset password of {email_address} in {round((time() - reset_start_time), 2)} seconds!")
                reset += 1
                completed += 1
                break
            # Check if account deactivated, reactivate if so
            elif getPixelColor(1030, 2090) == "fe2c55":
                print(f"{CWARNING}[WARNING] Account deactivated, reactivating...")
                device.shell("input tap 800 2040")
        break

# Clear app data, delete proxy, reboot
device.shell(f"pm clear {package_name}")
device.shell("settings delete global http_proxy")
device.shell("settings delete global global_http_proxy_host")
device.shell("settings delete global global_http_proxy_port")
device.shell("settings delete global global_http_proxy_exclusion_list")
device.shell("settings delete global global_proxy_pac_url")
device.shell("reboot")
queue = False
input(f"{CSUCCESS}Successfully completed all tasks, press ENTER to close...")