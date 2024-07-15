import datetime
import time
import sys
import winreg as wrg 
# pip install APScheduler
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.schedulers.base import STATE_STOPPED,STATE_RUNNING,STATE_PAUSED
# pip install infi.systray
from infi.systray import SysTrayIcon

def timed_job():
    #curTime = datetime.now()
    #curHour = curTime.hour
    #curMin = curTime.minute
    curHour = time.strftime('%H') # %H = in 24 hour format %I = 12 hour format %p Locale’s equivalent of either AM or PM.
    #curMin = time.strftime('%M')
    curHourString = str(curHour)
    newEndHour = None
    if curHourString[0] == "0":
        curHour = int(curHourString[1])
    else:
        curHour = int(curHour)
    if curHour <= 13:
        newEndHour = curHour + 10
    elif curHour >= 14:
        newEndHour = curHour - 10
    # Store location of HKEY_CURRENT_USER 
    location = wrg.HKEY_LOCAL_MACHINE
    # Storing path in soft 
    # SOFTWARE\Microsoft\WindowsUpdate\UX\Settings
    soft = wrg.OpenKeyEx(location,r"SOFTWARE\\Microsoft\\WindowsUpdate\\UX\\Settings\\", access=wrg.KEY_READ) 
    # Get Current Values As A BackUp
    ucahs = wrg.QueryValueEx(soft,"UserChoiceActiveHoursStart") 
    ucahe = wrg.QueryValueEx(soft,"UserChoiceActiveHoursEnd") 
    ahs = wrg.QueryValueEx(soft,"ActiveHoursStart") 
    ahe = wrg.QueryValueEx(soft,"ActiveHoursEnd") 
    sahs = wrg.QueryValueEx(soft,"SmartActiveHoursState")
    if soft:
        wrg.CloseKey(soft)

    try:
        # Set Active Hours to Current hour for start +10 hours for end
        if (type(curHour) == int) and (type(newEndHour) == int) and (curHour <= 24 and curHour >=1) and (newEndHour <= 24 and newEndHour >=1):
            soft = wrg.OpenKeyEx(location,r"SOFTWARE\\Microsoft\\WindowsUpdate\\UX\\Settings\\", access=wrg.KEY_WRITE) 
            wrg.SetValueEx(soft, "UserChoiceActiveHoursStart", 0, wrg.REG_DWORD, curHour) 
            wrg.SetValueEx(soft, "UserChoiceActiveHoursEnd", 0, wrg.REG_DWORD, newEndHour)
            wrg.SetValueEx(soft, "ActiveHoursStart", 0, wrg.REG_DWORD, curHour) 
            wrg.SetValueEx(soft, "ActiveHoursEnd", 0, wrg.REG_DWORD, newEndHour)
            # 0 is manual active hours 1 is "smart" or automatic
            wrg.SetValueEx(soft, "SmartActiveHoursState", 0, wrg.REG_DWORD, 0)
            if soft:
                wrg.CloseKey(soft)
            print("Set Active Hour Start To: ", curHour, " Set Active Hour End To: ", newEndHour, " Set Ative Hours To Manual")
        else:
            #print(curHour)
            #print(newEndHour)
            #print("type of curHour: ",type(curHour))
            #print("type of newEndHour: ",type(newEndHour))
            #print(type(curHour) == "int")
            print("Error getting curent hour or calculating end hour, not changing value!")
    except ValueError:
        if type(ucahs) == "tuple" and type(ucahs[0]) == "int":
            print("Got original active hours and they seem valid trying to restore!")
            try:
                soft = wrg.OpenKeyEx(location,r"SOFTWARE\\Microsoft\\WindowsUpdate\\UX\\Settings\\", access=wrg.KEY_WRITE)
                wrg.SetValueEx(soft, "UserChoiceActiveHoursStart", 0, wrg.REG_DWORD, ucahs[0]) 
                wrg.SetValueEx(soft, "UserChoiceActiveHoursEnd", 0, wrg.REG_DWORD, ucahe[0])
                wrg.SetValueEx(soft, "ActiveHoursStart", 0, wrg.REG_DWORD, ahs[0]) 
                wrg.SetValueEx(soft, "ActiveHoursEnd", 0, wrg.REG_DWORD, ahe[0])
                wrg.SetValueEx(soft, "SmartActiveHoursState", 0, wrg.REG_DWORD, sahs[0])
                if soft:
                    wrg.CloseKey(soft)
            except Exception as err:
                print(f"Unexpected {err=} when attempting to restore original active hours, {type(err)=}")
        else:
            print("Could not find valid previous values not attempting to restore!")
    except PermissionError:
        print("We Don't Have Permissin To Edit Registry, Unable to Update Active Hours!")
    except Exception as err:
        print(f"Unexpected {err=}, {type(err)=}")

# Setup Schedule
sched = BlockingScheduler()
#@sched.scheduled_job('interval', seconds=3600)
sched.add_job(timed_job, 'interval', hours=1)
def sched_start(systray):
    try:
        if sched.state == STATE_STOPPED:
            sched.start()
        elif sched.state == STATE_PAUSED:
            sched.resume()
    #except SchedulerAlreadyRunningError:
    #    print("scheduler already running!")
    #    pass
    except RuntimeError:
        print("running under uWSGI with threads disabled")
        pass
    except:
        pass

def sched_stop(systray):
    sched.pause()

# Setup Systray
hover_text = "Windows Always Active"
def say_hello(systray):
    print ("Hello")

def shutdown_tray(systray):
    try:
        sched.shutdown(wait=True)
    except:
        pass
    while sched.state == STATE_RUNNING:
        time.sleep(2)
    try:
        systray.shutdown()
    except:
        pass
    #if sched.state == STATE_STOPPED:
    #    sys.exit()
    #sys.exit()


#menu_options = (("Say Hello", None, say_hello),("Say Hello", None, say_hello),)
menu_options = ()
systray = SysTrayIcon("icon.ico", hover_text, menu_options, on_quit=shutdown_tray)
systray.start()
timed_job()
sched_start(systray)
