RUNTIME = '12:28:00 pm'
TAG = 'NGPL'


import schedule
import subprocess
import getpass
import time
from datetime import datetime
import datetime as dt
import os
import getpass
import Delivery as de
import Receipt as re
import Segment as se
from requests import get
USER_NAME = getpass.getuser()

everymin = 5
bat_path = None
def check_tag(tag = ''):
    link = 'https://pipeline2.kindermorgan.com/Capacity/OpAvailPoint.aspx?code=' + tag
    print(link)
    status = get(link).status_code
    print(status)

    if status == 200:
        return True
    else:
        return False

def get_filefolder_path(name):
    currentpath = os.path.realpath(__file__)
    print(currentpath)
    # add_to_startup(currentpath)

    patharr = currentpath.split('\\')
    patharr[len(patharr)-1] = name
    ptr = ('\\').join(patharr)
    return ptr


def add_to_startup(file_path=""):
    if file_path == "":
        dir_path = os.path.dirname(os.path.realpath(__file__))
        file_path = os.path.realpath(__file__)
    bat_path = dir_path + '\\' + TAG +'.bat'
    existflag = os.path.exists(bat_path)
    if(existflag == False):
        with open(bat_path, "w+") as bat_file:
            bat_file.write(r'python %s' % file_path)
            bat_file.close()
    vbs_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
    vbs_full_pasth = vbs_path + '\\' + TAG + ".vbs"
    existflag = os.path.exists(vbs_full_pasth)
    if(existflag == False):
        with open(vbs_full_pasth, "w+") as vbs_file:
            vbs_file.write('Set WshShell = CreateObject("WScript.Shell") \n')
            vbs_file.write('WshShell.Run chr(34) & "%s" & Chr(34), 0 \n' % bat_path)
            
            vbs_file.write('Set WshShell = Nothing')
            vbs_file.close()
print(USER_NAME)

def job():
    print('time reached')
    de.main(TAG)
    re.main(TAG)
    se.main(TAG)
    
if (check_tag(TAG)):
        
    add_to_startup()

    settime = datetime.strptime(RUNTIME, '%I:%M:%S %p')
    timestring = settime.strftime("%H:%M:%S")
    print(timestring)
    schedule.every().day.at(timestring).do(job) 
    while True: 
        schedule.run_pending()
        time.sleep(1)
