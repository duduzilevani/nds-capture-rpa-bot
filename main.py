#Bot_NDSCapture.py

import os
import time
import shutil
import socket
import splunk
import getpass
import logging
import openpyxl
import pyautogui
import pandas as pd
from datetime import date
from datetime import datetime
from configurations import Configurations
from email_module.email_service import Email
from splunk_batch import SplunkBatchLogger

print('Process starting')
logging.debug(str(time.ctime(time.time())) + ' ' + 'Process starting...')

"""Splunk code and funtions"""


def login_nds(username, password):

    login = pyautogui.locateCenterOnScreen(path_pics1 + '\\Capture.png')

    if login is None:
        return 1

    pyautogui.click(login)

    time.sleep(5)
    pyautogui.typewrite(username)
    pyautogui.press('tab')
    time.sleep(2)
    pyautogui.typewrite(password)
    time.sleep(3)
    pyautogui.typewrite(['enter'])
    pyautogui.typewrite(['enter'])
    time.sleep(7)
    logging.debug(str(time.ctime(time.time())) + ' ' + 'Logging in to NDS successful ')


def data_frame(count2, df1, accNum, amtL, csrL, count1, accNum2L, amt2L, csr2L):

    logging.debug(str(time.ctime(time.time())) + ' ' + 'Cleaning data')
    while count2 < len(df1):
        if str(df1.iloc[count2, 1]) == 'nan':
            count2 += 1
        else:
            acc1 = (int(df1.iloc[count2, 1]))
            amt_1 = (df1.iloc[count2, 3])
            amt_1 = str('%.2f' % amt_1)
            amt_1 = str(amt_1.replace(".", ''))
            csr1 = str(df1.iloc[count2, 4])
            if csr1.endswith('.0'):
                csr1 = csr1.replace('.0', '')
            accNum.append(str(acc1))
            amtL.append(str(amt_1))
            csrL.append(str(csr1))
            count2 += 1

    while count1 < len(df1):
        if str(df1.iloc[count1, 6]) == 'nan':
            count1 += 1
        else:
            acc2 = (int(df1.iloc[count1, 6]))
            amt_2 = (df1.iloc[count1, 8])
            amt_2 = str('%.2f' % amt_2)
            amt_2 = str(amt_2.replace(".", ''))
            csr2 = str(df1.iloc[count1, 9])
            if csr2.endswith('.0'):
                csr2 = csr2.replace('.0', '')
            accNum2L.append(str(acc2))
            amt2L.append(str(amt_2))
            csr2L.append(str(csr2))
            count1 += 1


def debit_screen():
    logging.debug(str(time.ctime(time.time())) + ' ' + 'Navigating on the Debit screen')
    time.sleep(3.5)
    value = str(Configurations('FOLDERS', 'value').read_config())
    f, g = map(int, value.split(','))
    pyautogui.click(f, g)
    # pyautogui.click(value1)
    time.sleep(2)
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(3)

    time.sleep(2)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(0.1)
    pyautogui.press('tab')
    # Added below two lines 2023-05-18
    time.sleep(0.1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')


import os
import pandas as pd
import logging
import time

def main_data_frame():
    global file_, path  # use both globals

    # Build the full path explicitly
    full_path = os.path.join(path, file_)

    logging.debug(str(time.ctime(time.time())) + f" Reading workbook: {full_path}")
    print("DEBUG: about to read ->", repr(full_path))

    # Safety checks before reading
    if not os.path.isfile(full_path):
        logging.debug(str(time.ctime(time.time())) + f" Not a file: {full_path}")
        raise FileNotFoundError(f"Not a valid file: {full_path}")

    size = os.path.getsize(full_path)
    logging.debug(str(time.ctime(time.time())) + f" File size: {size} bytes")
    print("DEBUG: size(bytes) =", size)

    # If file is empty or tiny, treat it as bad
    if size < 100:  # you can adjust this threshold
        raise OSError(f"File too small to be a valid Excel file: {full_path}")

    # Force openpyxl engine (more reliable for xlsx)
    df = pd.read_excel(full_path, engine="openpyxl")

    logging.debug(str(time.ctime(time.time())) + ' Reading file into memory')
    capture(df)



def debit_navigation(d_accNum1L, d_amt1L, d_csr1L, pics3):
    debit_range = str(Configurations('BOT', 'debit_range').read_config())
    # print("Debit Navigation")
    counter = 0
    for x, y, z in zip(d_accNum1L, d_amt1L, d_csr1L):
        counter +=1
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Checking the length of account number')
        for inth in range(5):
            if len(str(x)) < 9:
                x = ('0' + str(x))
            else:
                x = x
        pyautogui.typewrite(x)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Click amount field')
        time.sleep(1)
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        time.sleep(2)
        amount = str(Configurations('FOLDERS', 'amount').read_config())
        x, q = map(int, amount.split(','))
        pyautogui.click(x,q)
        time.sleep(2)

        pyautogui.typewrite(y)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Account number is valid')
        time.sleep(0.5)
        pyautogui.hotkey('tab')
        pyautogui.hotkey('space')
        for d in range(int(debit_range)):
            pyautogui.hotkey('down')

        pyautogui.hotkey('space')
        pyautogui.hotkey('tab')
        pyautogui.typewrite(str(z))
        pyautogui.hotkey('tab')
        time.sleep(1)
        pyautogui.hotkey("enter")
        time.sleep(2)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Writing to Splunk')
        splunk_logger.add_event(
            status=1,
            start=START,
            finish=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            unique_id='233',
            heartbeat=1
        )


def credit_navigation(c_accNum2, c_amt2, c_csr2, pics2):
    credit_range = str(Configurations('BOT', 'credit_range').read_config())
    counting = 0
    for a, b, c in zip(c_accNum2, c_amt2, c_csr2):
        counting += 1
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Checking the length of the account number')
        for length in range(5):
            if len(str(a)) < 9:
                a = ('0' + str(a))
            else:
                a = a
        time.sleep(3)
        pyautogui.typewrite(a)
        time.sleep(0.5)
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        time.sleep(2)
        pyautogui.hotkey('tab')
        time.sleep(1)

        time.sleep(0.5)
        pyautogui.typewrite(b)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Credit screen - account number is valid')
        time.sleep(2)
        pyautogui.hotkey('tab')
        pyautogui.hotkey('space')
        time.sleep(1)
        for e in range(int(credit_range)):
            pyautogui.hotkey('down')
        pyautogui.hotkey('enter')
        pyautogui.hotkey('tab')
        pyautogui.hotkey('space')
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        pyautogui.typewrite(str(c))
        pyautogui.hotkey('enter')
        time.sleep(2)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Writing to Splunk')
        splunk_logger.add_event(
            status=1,
            start=START,
            finish=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            unique_id='233',
            heartbeat=1
        )


def conclude(requester1):
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.typewrite(str(requester1))
    time.sleep(3)
    logging.debug(str(time.ctime(time.time())) + ' ' + 'Updating the requester details on the comments field')
    try:
        aut = pyautogui.locateCenterOnScreen(path_pics1 + '\\aut.png', confidence=0.9
                                             ,
                                                       grayscale=True)
    except pyautogui.ImageNotFoundException:
        aut = None
    if aut:
        pyautogui.click(aut)
        time.sleep(3)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Moving files to successful folder')
        src_ = path1 + file_
        dst = moveto + file_
        shutil.move(src_, dst)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'moving to another excel')
    else:
        Email(maintenance, from_, 'Exception - BDS Reject', f"Good Day, <br><br>Please note BDS entry:{file_} was rejected by BDS.<br><br>Kind Regards<br>NDS Capture Bot").mail_send()
        move_exceptions(moveto2)


def sending_successful_email(df, p):
    cc = df.iloc[22, 7]
    from_who = str(Configurations('EMAILING', 'FROM').read_config())
    to = df.iloc[23, 2]
    subject = str(os.path.basename(p))
    body = 'Good day<br><br>Kindly note that your card payment request: ' + str(subject) + \
           ' has been successfully captured and has been sent for release. <br><br> Kind regards<br> Card Payment Team'
    logging.info("Sending successful email, " + str(time.ctime(time.time())))
    Email(to, from_who, subject, body, cc).mail_send()

def safe_get(df, row_idx, col_idx, default=None):
    """Safely get a value from DataFrame, return default if out of bounds."""
    try:
        return df.iloc[row_idx, col_idx]
    except IndexError:
        return default

def capture(df, default=None):
    date1 = str(safe_get(df, 22, 2, default="MISSING_DATE"))
    requester = safe_get(df, 23, 2, default="MISSING_REQUESTER")

    if date1 == "MISSING_DATE" or requester == "MISSING_REQUESTER":
        Email(maintenance, from_, f"Exception - NDS Capture Bot", f"Good Day<br><br> Please note there is an exception on {file_}. Please check the length of the file.<br><br>Kind Regards<br>NDS Capture Bot").mail_send()
        move_exceptions(moveto1)
    else:

        accNum1L = []
        accNum2L = []
        amt1L = []
        amt2L = []
        csr1L = []
        csr2L = []
        count = 0
        count1 = 0
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Slicing of data-frame and creating lists')

        data_frame(count, df, accNum1L, amt1L, csr1L, count1, accNum2L, amt2L, csr2L)
        debit_screen()
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Date validation')
        if date1 == str(date.today()):
            pyautogui.press('enter')
            time.sleep(2)
            logging.debug(str(time.ctime(time.time())) + ' ' + 'validate date')
        else:
            pyautogui.press('down')
            pyautogui.press('enter')
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')

        debit_navigation(accNum1L, amt1L, csr1L, path_pics1)
        click_credit = str(Configurations('FOLDERS', 'credit').read_config())
        x, yr = map(int, click_credit.split(','))
        pyautogui.click(x, yr)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Navigating on the credit screen')

        time.sleep(2)
        credit_navigation(accNum2L, amt2L, csr2L, path_pics1)
        time.sleep(1)
        transfer = str(Configurations('FOLDERS', 'transfer').read_config())
        ab, cq = map(int, transfer.split(','))
        pyautogui.click(ab, cq)
        time.sleep(2)

        conclude(requester, path_pics1)
        logging.debug(str(time.ctime(time.time())) + ' ' + 'Completed capturing file')


def move_exceptions(moving):
    src_ = path1 + file_
    dst = moveto + file_
    exception = moving + file_
    shutil.move(src_, exception)
    logging.debug(str(time.ctime(time.time())) + ' ' + 'Moving files to exceptions folder')
    splunk_logger.add_event(
        status=0,
        start=START,
        finish=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        unique_id='233',
        heartbeat=1
    )


if __name__ == "__main__":
    splunk_logger = SplunkBatchLogger(splunk, batch_size=5, index="acoe_bot_events")
    START = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # global passed
    passed = 0
    # global index
    index = 'acoe_bot_events'  # (Always the same non prod, it will change once in prod)
    # global uniqueID
    uniqueID = '233'  # (Obtain your u  nique id from the RDA Tracker)
    robot_user = getpass.getuser()
    bot_machine = socket.gethostname()
    host = bot_machine
    auth = str(Configurations('NDS', 'auth').read_config())
    url = str(Configurations('NDS', 'splunk_url').read_config())
    """End of splunk functions"""
    path1 = str(Configurations('FOLDERS', 'PATH1').read_config())
    moveto = str(Configurations('SUCCESSFUL', 'MOVETO').read_config())
    moveto1 = str(Configurations('EXCEPTIONS', 'MOVETO1').read_config())
    moveto2 = str(Configurations('EXCEPTIONS', 'moveto2').read_config())
    copyto = str(Configurations('SUCCESSFUL', 'CopyToDesktop').read_config())
    path_pics1 = str(Configurations('PICTURES', 'PICS').read_config()       )

    logging.debug(str(time.ctime(time.time())) + ' ' + 'NDS Bot Capturing Process started')
    nds = 'BDS.lnk'
    path = str(Configurations('EXCEL_FOLDERS', 'PATH').read_config())
    maintenance = str(Configurations('EMAIL', 'maintenance').read_config())

    from_ = str(Configurations('EMAIL', 'FROM').read_config())
    os.chdir(path)
    fileNames = os.listdir()
    print(len(fileNames))

    for file_ in fileNames:
        # Skip Windows thumbnail database and any obvious junk
        if file_.lower() == "thumbs.db":
            logging.debug(str(time.ctime(time.time())) + f" Skipping system file: {file_}")
            continue

        # Only process Excel files
        if not file_.lower().endswith((".xlsx", ".xls", ".xlsm")):
            logging.debug(str(time.ctime(time.time())) + f" Skipping non-Excel file: {file_}")
            continue

        # (Optional safety) skip directories, just in case
        if not os.path.isfile(file_):
            logging.debug(str(time.ctime(time.time())) + f" Skipping directory: {file_}")
            continue

        logging.debug(str(time.ctime(time.time())) + ' ' + 'Working through workbook ' + file_)

        image_path = f"{path_pics1}landing.png"
        print(image_path)

        try:
            location = pyautogui.locateOnScreen(image_path, confidence=0.8)
        except pyautogui.ImageNotFoundException:
            location = None

        if location:
            try:
                print(file_)  # file_ is still global-friendly here
                main_data_frame()  # uses global file_
            except Exception as errs:
                logging.debug(str(time.ctime(time.time())) + ' ' + str(errs) + ' Error While looping in workbooks')
                print('Exception: ' + str(errs))
                move_exceptions(moveto1)  # also uses global file_
        else:
            Email(
                maintenance,
                from_,
                'Exception on Bot',
                'Good Day<br><br> Please note there is an exception on the bot: '
                'ETS/NDS not visible on screen, please logon and restart application<br><br>'
                'Kind Regards<br>NDS Capture Bot'
            ).mail_send()
