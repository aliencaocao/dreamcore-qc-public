### WARNING: Please read README thoroughly before attempting to execute/modify/borrow codes from this script.
### This script is not meant to be ran standalone. It require many external files that are not uploaded to the repository. It is purely for demonstration purpose.
### This script can make changes to your Windows PC and may cause data loss or system instability. It CAN format your drives, change your partitions, reboot your computer and install/uninstall software. Do NOT execute any part of this script without knowing what you are doing.
### Important information such as API keys are redacted.
import csv
import ctypes
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import winreg as reg
from ctypes import wintypes
from datetime import datetime, timedelta
from pathlib import Path
from time import sleep
from tkinter import *
from tkinter import messagebox
from tkinter.font import nametofont
from tkinter.ttk import *

import colorama
import keyboard
import pandas as pd
import psutil
import pyrebase
import telegram
import wmi
from bleak import BleakScanner
from colorama import Fore
from telegram.ext import Updater

__build__ = 92
__version__ = '2.3.8'
__dev__ = True
__author__ = 'Billy Cao'
__email__ = 'aliencaocao@gmail.com'
frozen = getattr(sys, 'frozen', False)  # frozen -> running in exe
pd.set_option('mode.chained_assignment', None)  # silence some df slicing and renaming warning
kernel32 = ctypes.windll.kernel32

online = None
time_synced = False
allow_mode_switch = False
__debug_switch_count__ = 0  # when it reaches 5 enable debugger

exec_path = sys.executable
workingDir = os.getcwd()
desktop = Path.home() / 'Desktop'
restart_shortcut_path = os.path.join(os.environ['ProgramData'], r'Microsoft\Windows\Start Menu\Programs\Startup', 'startup.bat')
driversFolder = 'drivers'
_7z_return_code = {1: 'Warning (Non-fatal error)',
                   2: 'Fatal error',
                   7: 'Command line error',
                   8: 'Not enough memory for operation',
                   255: 'User stopped the process'}

item_failed = set()
item_failed_first_run = True
old_error_device = None
iotest_results = {'USB': None, 'Audio': None}

qc_person = None
order_no = None
note = None

firebase = pyrebase.initialize_app({"apiKey": "redacted",
                                    "authDomain": "redacted.firebaseapp.com",
                                    "databaseURL": "https://redacted.firebasedatabase.app",
                                    "storageBucket": "redacted.appspot.com"})
db = firebase.database()

bot_updater = Updater(token='redacted')
failed_to_send_bot_msgs = dict()
failed_items_telegram_msg = None
failed_items_msg_sent = None

verbose_logging = os.path.isfile('debug.txt')
if verbose_logging: print('INFO: Enabled debug verbose logging (can be disabled by removing debug.txt from launch directory)')


class Unbuffered:
    def __init__(self, stream):
        self.stream = stream

    def write(self, data):
        self.stream.write(data)
        data = data.replace(Fore.WHITE, '').replace(Fore.YELLOW, '').replace(Fore.LIGHTRED_EX, '').replace(Fore.RESET, '').replace(Fore.GREEN, '')  # remove all the colour formatting escape characters
        logFile.write(data)
        logFile.flush()  # force file write

    def flush(self):  # for Python 3 compatibility
        pass


def get_res_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
     Relative path will always get extracted into root!"""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    if os.path.isfile(os.path.join(base_path, relative_path)):
        return os.path.join(base_path, relative_path)
    else:
        raise FileNotFoundError(f'Embedded file {os.path.join(base_path, relative_path)} is not found!')


def exec_live_output(name, p):
    """Takes in a subprocess.Popen object and redirect its stdout live to sys.stdout
    Both stdout and stderr need to be subprocess.PIPE"""
    for line in p.stdout:
        try:
            logger.info(name + ': ' + line.decode('GB2312').replace('\n', '').strip())  # for chinese locale systems
        except UnicodeDecodeError:
            logger.info(name + ': ' + line.decode('ISO-8859-1').replace('\n', '').strip())
        root.update()


def if_running(processName):
    return processName.lower() in [p.name().lower() for p in psutil.process_iter()]


def _quit_tk():
    root.quit()
    root.destroy()


def _quit_debugger():
    global __debug_switch_count__, debugger_window
    __debug_switch_count__ = 0
    debugger_window.destroy()
    debugger_window.quit()


def delete_sub_key(rootKey, sub, task, silent=False):
    if verbose_logging:
        silent = False  # force logging
        logger.debug(f'Removing {task} registry data from {rootKey}/{sub}')  # too much spam, only enable for debug
    try:
        open_key = reg.OpenKey(rootKey, sub, 0, reg.KEY_ALL_ACCESS)
        num, _, _ = reg.QueryInfoKey(open_key)
        for i in range(num):
            child = reg.EnumKey(open_key, 0)
            delete_sub_key(open_key, child, task)
        try:
            reg.DeleteKey(open_key, '')
        except Exception as e:
            if not silent: logger.error(f'Error when removing {task} registry settings from {rootKey}/{sub}: {e}')
        finally:
            reg.CloseKey(open_key)
    except Exception as e:
        if mode != 2 and not silent: logger.error(f'Error when removing {task} registry settings from {rootKey}/{sub}: {e}')  # mode 2 is just after OOB, so some reg are cleared alr


if frozen:
    import pyi_splash

    pyi_splash.update_text('Initializing logging system...')
logger = logging.getLogger(__name__)  # only show logs generated by this script, block all other logs generated by imported stuff
logging.addLevelName(15, 'STATUS')  # between debug 10 and info 20
logger.setLevel(logging.DEBUG)
colorama.init()


class SYSTEMTIME(ctypes.Structure):  # https://docs.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-systemtime
    _fields_ = [('wYear', wintypes.WORD),
                ('wMonth', wintypes.WORD),
                ('wDayOfWeek', wintypes.WORD),
                ('wDay', wintypes.WORD),
                ('wHour', wintypes.WORD),
                ('wMinute', wintypes.WORD),
                ('wSecond', wintypes.WORD),
                ('wMilliseconds', wintypes.WORD)]


SystemTime = SYSTEMTIME()


def get_system_time():
    lpSystemTime = ctypes.pointer(SystemTime)
    kernel32.GetLocalTime(lpSystemTime)
    return datetime.strptime(f'{SystemTime.wYear}-{SystemTime.wMonth}-{SystemTime.wDay} {SystemTime.wHour}:{SystemTime.wMinute}:{SystemTime.wSecond}:{SystemTime.wMilliseconds}', '%Y-%m-%d %H:%M:%S:%f')


class real_time_formatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        return datetime.strftime(get_system_time(), '%Y-%m-%d %H:%M:%S:%f')[:-3]  # remove last 3 digits for milliseconds


class ColorLog(logging.Formatter):
    grey = Fore.WHITE
    yellow = Fore.YELLOW
    red = Fore.LIGHTRED_EX
    reset = Fore.RESET

    if verbose_logging:
        format = '{asctime}: {levelname} - {message} ({module}: {lineno} in {funcName})'
    else:
        format = '{asctime}: {levelname} - {message}'

    FORMATS = {logging.DEBUG: grey + format + reset,
               15: grey + format + reset,  # custom STATUS logging level
               logging.INFO: grey + format + reset,
               logging.WARNING: yellow + format + reset,
               logging.ERROR: red + format + reset,
               logging.CRITICAL: red + format + reset}

    def format(self, record):
        log_fmt = self.FORMATS[record.levelno]
        formatter = real_time_formatter(log_fmt, style='{')
        return formatter.format(record)


logFile = open(os.path.join(desktop, 'log.log'), 'a+', encoding='utf-8')
if os.stat(os.path.join(desktop, 'log.log')).st_size: logFile.write('\n\n')  # only append new lines if not empty (size != 0), file.read() related don't work (always not empty)
sys.stdout = Unbuffered(sys.stdout)  # redirect all prints to log file
sys.stderr = sys.stdout  # redirect all errors to log file
streamHandler = logging.StreamHandler(sys.stdout)
streamHandler.setFormatter(ColorLog())
logger.addHandler(streamHandler)  # stdout here already redirected to log file so no need file handler (they conflict)
logger.info(f'Dreamcore QC Software V{__version__}/Build {__build__} by {__author__} ({__email__})\nRunning on Windows Management Instrumentation (WMI) V{wmi.__version__}\nWorking Directory: {workingDir}\nExecutable path: {exec_path}')
if __dev__: logger.debug('ALERT: This is a developer build, some features may be disabled or broken!')
logger.info('Setting time zone to SGT (UTC +8)')
subprocess.call('tzutil /s "Singapore Standard Time"')

try:
    with open(os.path.join(os.environ['ProgramData'], 'mode.txt'), 'r+') as f:
        mode = f.read()
        if mode == '1':
            mode = 1
            logger.info('Entered Benchmark Mode')
        elif mode == '2':
            mode = 2
            logger.info('Entered OOB Mode')
        elif mode == '3':
            mode = 3
            logger.info('Entered Restore Point Mode')
        else:
            mode = 0
except FileNotFoundError:
    mode = 0
    with open(os.path.join(os.environ['ProgramData'], 'mode.txt'), 'w+') as f:
        f.write('0')

if frozen:
    pyi_splash.update_text('Initializing GUI...')
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # fix blurry text on high DPI screens
except:
    pass
wingdi = ctypes.CDLL("gdi32")
wingdi.AddFontResourceExW.argtypes = [ctypes.c_wchar_p, ctypes.wintypes.DWORD, ctypes.c_void_p]
wingdi.AddFontResourceExW(str(Path(__file__).with_name("Nunito.ttf")), 0x10, 0)  # 0x10 flag means extract to memory temporarily, will not install it on system


class ScrolledWindow(Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.parent = parent

        # creating scrollbars
        self.xscrlbr = Scrollbar(self.parent, orient=HORIZONTAL)
        self.xscrlbr.pack(fill=X, side=BOTTOM)
        self.yscrlbr = Scrollbar(self.parent)
        self.yscrlbr.pack(fill=Y, side=RIGHT)

        # creating canvas
        self.canv = Canvas(self.parent)
        # noinspection PyArgumentList
        self.canv.config(relief='flat', width=10, heigh=10, bd=2)
        self.canv.pack(side=LEFT, fill=BOTH, expand=True)

        # accociating scrollbar comands to canvas scroling
        self.xscrlbr.config(command=self.canv.xview)
        self.yscrlbr.config(command=self.canv.yview)

        # creating a frame to inserto to canvas
        self.scrollwindow = Frame(self.parent)
        self.canv.create_window(0, 0, window=self.scrollwindow, anchor=NW)
        self.canv.config(xscrollcommand=self.xscrlbr.set, yscrollcommand=self.yscrlbr.set, scrollregion=self.canv.bbox('all'))
        self.yscrlbr.lift(self.scrollwindow)
        self.xscrlbr.lift(self.scrollwindow)
        self.scrollwindow.bind('<Configure>', self._configure_window)
        self.scrollwindow.bind('<Enter>', self._bound_to_mousewheel)
        self.scrollwindow.bind('<Leave>', self._unbound_to_mousewheel)

    def _bound_to_mousewheel(self, event):
        self.canv.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbound_to_mousewheel(self, event):
        self.canv.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canv.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _configure_window(self, event):
        # update the scrollbars to match the size of the inner frame
        size = (self.scrollwindow.winfo_reqwidth(), self.scrollwindow.winfo_reqheight())
        # noinspection PyTypeChecker
        self.canv.config(scrollregion='0 0 %s %s' % size)
        if self.scrollwindow.winfo_reqwidth() != self.canv.winfo_width():
            # update the canvas's width to fit the inner frame
            self.canv.config(width=self.scrollwindow.winfo_reqwidth())
        if self.scrollwindow.winfo_reqheight() != self.canv.winfo_height():
            # update the canvas's width to fit the inner frame
            self.canv.config(height=self.scrollwindow.winfo_reqheight())


root = Tk()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.minsize(min(700, int(0.7 * screen_width)), 0)
root.maxsize(int(0.7 * screen_width), int(0.7 * screen_height))
root.resizable(False, False)
root.attributes("-topmost", True)
root.protocol("WM_DELETE_WINDOW", _quit_tk)
root.iconbitmap(get_res_path('icon.ico'))
tk_default_font = nametofont("TkDefaultFont")
tk_default_font.configure(family='Nunito', size=14)
root.bind('g', lambda event: debugger())

root.title(f'Dreamcore QC Software V{__version__} - Initializing')

status_bar = LabelFrame(root, padding=(10, 0))  # status bar must be at front else scrollable window will block it
status_bar.pack(expand=True, fill=X, side=BOTTOM, pady=(0, 10))
scroll_container_frame = Frame(root)
scroll_container_frame.pack(expand=True, fill=BOTH, side=TOP)
scrollableFrame = ScrolledWindow(scroll_container_frame)
frame = Frame(scrollableFrame.scrollwindow)
frame.pack(fill=BOTH, expand=True, anchor=CENTER, padx=(40, 0), pady=(15, 0))

status_bar.rowconfigure(0, weight=1)
status_bar.columnconfigure(1, weight=1)
testing_text = Label(status_bar, text='', anchor=W)
testing_text.grid(row=0, column=0, sticky=W)
status_text = Label(status_bar, text='Initializing', anchor=E)
status_text.grid(row=0, column=2, sticky=E)

ver_msg = Label(frame, text=f'Dreamcore QC Software V{__version__}\nRunning on Windows Management Instrumentation (WMI) V{wmi.__version__}\n')
ver_msg.pack(expand=True, fill=BOTH, side=TOP)

tk_style = Style()
tk_style.configure('green_text.TButton', foreground='green')
tk_style.configure('red_text.TButton', foreground='red')
tk_style.configure('green_text.TLabel', foreground='green')
tk_style.configure('red_text.TLabel', foreground='red')

if not os.path.isdir(driversFolder) and not mode:
    logger.error(f'Driver folder not found at {driversFolder}! Many functions including driver installation will be skipped, and you may face unexpected errors or crashes!')
    messagebox.showerror(title='Driver folder not found!', message=f'Driver folder not found at {driversFolder}! Many functions including driver installation will be skipped, and you may face unexpected errors or crashes!')

if mode:
    try:
        with open(os.path.join(os.environ['ProgramData'], 'tester.txt'), 'r') as f:
            line = f.read().splitlines()
            qc_person = line[0]
            order_no = line[1]
            note = line[2]
    except FileNotFoundError:  # need to pop up even in benchmark mode cuz its pretty serious error and could lead to tracking errors (esp if integrate with Monday or Sheets)
        messagebox.showerror(title='Failed to parse saved tester info', message=f'tester.txt is not found at {os.environ["ProgramData"]}. The program will use default tester information.')
        logger.error(f'tester.txt is not found at {os.environ["ProgramData"]}. The program will use default tester information as placeholder. Ignore this error if you manually entered benchmark or OOB mode via mode.txt.')
        qc_person = 'Default QC Person'
        order_no = '0000-00'
        note = 'default tester info used'

    logger.info(f'Order #{order_no} testing by {qc_person}. Note: {note}')
    testing_text['text'] = f'Order #{order_no} testing by {qc_person}'
    testing_info_label = Label(frame, text=f'Order #{order_no} testing by {qc_person}, note: {note}\n')
    testing_info_label.pack(expand=True, fill=BOTH, side=TOP)

logger.debug('GUI initialized. Initializing WMI')
if not ctypes.windll.shell32.IsUserAnAdmin():
    logger.error('You are not running QC software as Admin. Most of the functions will not work!')
    messagebox.showerror(title='Not running with admin!', message='You are not running QC software as Admin. Most of the functions will not work! Continuing will result in software crashing!')
if frozen:
    pyi_splash.update_text('Initializing WMI...')
pc = wmi.WMI()
logger.debug('WMI Initialized')
if mode != 1:
    if frozen:
        pyi_splash.update_text('Awaiting user confirmation...')
    messagebox.showwarning(title='Reminder', message='Do NOT close the software by closing the black console window. ONLY close it via the GUI window X button, else it will not perform proper clean up!')


def status(text, log=True):
    global status_text
    text = text.strip()
    status_text['text'] = text
    status_text.update()
    status_bar.pack(expand=True, fill=BOTH, side=BOTTOM, pady=(0, 10))
    text = text.replace('\n', ' ')
    if log:
        if text.lower() == 'ready':
            logger.log(15, text + '\n')  # custom STATUS logging level 15
        else:
            logger.log(15, text)
    root.title(f'Dreamcore QC Software V{__version__} - {text}')


def debugger():
    global __debug_switch_count__, verbose_logging, debugger_window
    __debug_switch_count__ += 1
    if __debug_switch_count__ >= 5:
        logger.debug('Debugger enabled')
        verbose_logging = True  # nothing to do with debugger functions itself, just make logging verbose
        debugger_window = Toplevel(root)
        debugger_window.title('Debugger')
        debugger_window.geometry('400x100')
        debugger_window.resizable(True, True)
        cmd_input_frame = Frame(debugger_window, padding=(10, 20))
        command_input = Entry(cmd_input_frame, font=tk_default_font)
        command_input.pack(expand=True, fill=BOTH, side=TOP)
        cmd_input_frame.pack(expand=True, fill=BOTH, side=TOP)
        command_input.focus_set()
        command_input.bind('<Return>', lambda event: exec(command_input.get(), globals()))
        debugger_window.protocol("WM_DELETE_WINDOW", _quit_debugger)
        root.attributes("-topmost", False)
        debugger_window.mainloop()  # blocking here until window is closed
        logger.debug('debugger closed')
        __debug_switch_count__ = 0  # reset switch count (not working??
        root.attributes("-topmost", True)


class BotTimerQueue:  # Use Windows OS uptime as a method to calculate the actual timestamp for unsent Telegram messages as we cannot rely on system time or time.time() since they are both inaccurate if the time is not synced (which is the case when there is no internet)
    def __init__(self):
        self.counter = 0
        self.timer_id_starts = {}  # simulates a locals()

    def start_new_timer(self):
        self.counter += 1
        self.timer_id_starts[self.counter] = kernel32.GetTickCount64()  # milliseconds
        return self.counter

    def end_timer(self, timer_id):
        return kernel32.GetTickCount64() - self.timer_id_starts[timer_id]


bot_timer_queue = BotTimerQueue()


def bot_send_msg(msg=None, tag=None):  # tag is for special uses like preserving sent message object in global, stored with msg in dict {msg: tag}
    global failed_items_telegram_msg, failed_items_msg_sent, failed_to_send_bot_msgs, failed_to_send_bot_msgs_copy

    def send(_msg, _tag=None, timer_id=None):
        global failed_items_telegram_msg, failed_items_msg_sent, failed_to_send_bot_msgs, failed_to_send_bot_msgs_copy
        _msg = str(_msg)
        if timer_id: timestamp = datetime.strftime(get_system_time() - timedelta(milliseconds=bot_timer_queue.end_timer(timer_id) + 4000), '%Y-%m-%d %H:%M:%S')  # no need include milliseconds in msg, add 4sec overhead as try/except timeout is about 4sec
        try:
            if timer_id:
                sent_msg = bot_updater.bot.sendMessage(chat_id='redacted', text=timestamp + ': ' + _msg)  # My chat: redacted, DC group: redacted
            else:
                sent_msg = bot_updater.bot.sendMessage(chat_id='redacted', text=_msg)
        except telegram.error.NetworkError as e:
            logger.warning(f'Failed to send message:\n{_msg}\nto Telegram due to internet error. The software will try again later. Detailed error:\n{e}')
            if _msg not in failed_to_send_bot_msgs.keys(): failed_to_send_bot_msgs.update({_msg: {'tag': _tag, 'timer': bot_timer_queue.start_new_timer()}})  # because time is different, so must add a substr check here to prevent duplicate msgs
            if _tag == 'failed_items': failed_items_msg_sent = False
        except Exception as e:
            logger.warning(f'Failed to send message:\n{_msg}\nto Telegram due to unknown error. The software will try again later. Detailed error:\n{e}')
            if _msg not in failed_to_send_bot_msgs.keys(): failed_to_send_bot_msgs.update({_msg: {'tag': _tag, 'timer': bot_timer_queue.start_new_timer()}})
            if _tag == 'failed_items': failed_items_msg_sent = False
        else:
            try:
                failed_to_send_bot_msgs.pop(_msg)
            except KeyError:  # if first time succeed, the failed message set will not contain this message, thus will have key error if trying to remove it
                pass
            if _tag == 'failed_items':
                failed_items_telegram_msg = sent_msg  # preserve sent message obj of failed items list for later editing
                failed_items_msg_sent = True
            return sent_msg

    if online:
        failed_to_send_bot_msgs_copy = failed_to_send_bot_msgs.copy()  # make a copy here for later use
        if failed_to_send_bot_msgs:
            logger.debug(f'Sending {len(failed_to_send_bot_msgs)} unsent messages to Telegram...')
            for msg_i, info_i in failed_to_send_bot_msgs_copy.items():
                send(msg_i, _tag=info_i['tag'], timer_id=info_i['timer'])  # not returning sent message objects here as they are not ordered (failed to send msg is a set which is unordered)
        if msg and msg not in failed_to_send_bot_msgs.keys():  # check if the message to be sent is not in the old failed msgs list
            sent_msg = send(msg, tag)
            return sent_msg
    else:
        if check_online(retry_telegram=False):
            bot_send_msg(msg, tag)  # call itself again if online now, will be redirected to if online part so it won't be looping
        elif msg:
            logger.warning(f'Failed to send message:\n{msg}\nto Telegram as software is offline. Will try again later.')
            failed_to_send_bot_msgs.update({msg: {'tag': tag, 'timer': bot_timer_queue.start_new_timer()}})


def check_online(timeout=3, retry_telegram=True):
    status('Checking internet connection')
    global online
    if subprocess.run(f'ping google.com -n 1 -w {timeout * 1000}', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode:
        online = False
        logger.warning('Internet connection test failed: ping google.com failed')
    else:
        online = True
        logger.info(f'{Fore.GREEN}Internet connection test passed{Fore.RESET}')
        if not time_synced:
            sync_time()  # need to sync time first else may have SSL error with firebase
        if not db.child('DClicense').shallow().get().val():
            messagebox.showerror(title='Unauthorised copy', message=f'This copy of QC software is unauthorised. Please contact developer {__author__} at {__email__}. The software will exit now.')
            logger.info('Program exit (If this console window does not close itself, you can manually close it now)')
            logFile.close()
            sys.exit(1)
        elif verbose_logging:
            logger.debug('License authenticated successfully.')
        if failed_to_send_bot_msgs and retry_telegram:
            bot_send_msg()

    status('Ready', log=False)
    return online


def sync_time():
    status('Syncing time')
    global online, time_synced

    if not online:
        check_online()

    if not online:
        logger.warning('You are not connected to Internet. Skipping time sync.')
        item_failed.add('Time sync - not connected to Internet')
    elif not time_synced:  # in check online it may already synced time so no need to sync again
        status('Syncing time', log=False)  # overrides the status set by other functions
        logger.info('Starting Windows Time Service')
        subprocess.run('net start w32time', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  # supress 'service already started' error
        subprocess.run('w32tm /config /manualpeerlist:"time.nist.gov" /syncfromflags:manual /reliable:yes /update', stdout=subprocess.DEVNULL)
        subprocess.run('w32tm /config /manualpeerlist:"time.windows.com" /syncfromflags:manual /reliable:yes /update', stdout=subprocess.DEVNULL)
        subprocess.run('w32tm /config /update', stdout=subprocess.DEVNULL)  # force w32tm service config refresh
        logger.info('Syncing time')
        tries = 1
        before_sync_epoch = time.time()
        before_sync_uptime = kernel32.GetTickCount64() / 1000
        try:
            r = subprocess.check_output('w32tm /resync /rediscover').decode('GB2312')
        except subprocess.CalledProcessError:
            r = ''
        after_sync_epoch = time.time()
        after_sync_uptime = kernel32.GetTickCount64() / 1000
        if abs(after_sync_epoch - before_sync_epoch) * 0.9 > abs(after_sync_uptime - before_sync_uptime):  # if the change in seconds since epoch is larger than change in uptime, means that the system time has changed a lot and time sync is likely successful, 10% extra strict to ensure system time changed a lot. This is to reduce false failures from Windows side
            r = 'successfully 成功'  # just a success placeholder
        while 'successfully' not in r and '成功' not in r and tries <= 10:  # no matter success or fail all return 0
            logger.warning(f'Time sync failed. Retrying (try no {tries})\nDetailed error:')
            print(r)
            root.update()
            before_sync_epoch = time.time()
            before_sync_uptime = kernel32.GetTickCount64() / 1000
            try:
                r = subprocess.check_output('w32tm /resync /rediscover').decode('GB2312')
            except subprocess.CalledProcessError:
                r = ''
            after_sync_epoch = time.time()
            after_sync_uptime = kernel32.GetTickCount64() / 1000
            if abs(after_sync_epoch - before_sync_epoch) * 0.9 > abs(after_sync_uptime - before_sync_uptime):
                r = 'successfully 成功'  # just a success placeholder
            tries += 1
        if tries <= 10:
            logger.info(f'{Fore.GREEN}Time synced successfully{Fore.RESET}')
            time_synced = True
            try:
                item_failed.remove('Time sync - not connected to Internet')
            except KeyError:
                pass
            try:
                item_failed.remove('Time sync failed')
            except KeyError:
                pass
        else:
            logger.error('Time sync failed after 10 tries. Please try manually syncing.')
            item_failed.add('Time sync failed')
    status('Ready')


def b_to_gb(b, dp=2):
    return round(float(b) / pow(1024, 3), dp)


if frozen:
    pyi_splash.update_text('Syncing time...')

if not time_synced: sync_time()
if not time_synced: logger.warning('Time on this system have not been synced with Internet, thus date and time of logs may not be accurate. Another attempt to sync the time will be done after driver installation.')

if mode != 1 and mode != 3:  # Do not auto check for update in Benchmark and Restore Point modes.
    if online:
        status('Checking for update')
        if frozen:
            pyi_splash.update_text('Checking for update...')
        try:
            latest_build = db.child('DCbuild').shallow().get().val()
        except Exception as e:
            logger.error(f'An error occurred when getting the latest version: {e}, skipping update check')
        else:
            if latest_build > __build__:  # not checking for unequal here as build might be higher than actual latest release due to testing version
                if messagebox.askyesno(title='Update available!', message=f'Latest version is build {latest_build}, you are running build {__build__}. Do you want to update?\nChange log:\n{db.child("DCchangelog").shallow().get().val()}'):
                    status('Downloading update')
                    logger.info(f'Downloading latest build: {latest_build}')
                    if frozen:
                        pyi_splash.update_text('Downloading update...')
                    try:
                        import gdown

                        link = 'https://drive.google.com/uc?id=' + db.child('DCupdate_link').shallow().get().val().split('/')[-2]
                        gdown.download(link, 'new.7z', quiet=False)
                    except Exception as e:
                        logger.error(f'An error occurred when downloading update: {e}')
                        messagebox.showerror(title='An error occurred when downloading', message=f'An error occurred when downloading update: {e}\nThe program will now exit.')
                        sys.exit(1)
                    else:
                        status('Awaiting user input')
                        logger.info(f'{Fore.GREEN}Update downloaded successfully.{Fore.RESET}')
                        if messagebox.askyesno(title='Update is ready', message=f'Update downloaded successfully. You can relaunch the software to automatically update, or manually extract the newly downloaded new.7z (password: dreamcore) to replace the current running exe. Relaunch now?'):
                            shutil.copy(get_res_path('7za.exe'), '7za.exe')
                            shutil.copy(get_res_path('self_update.bat'), 'self_update.bat')
                            subprocess.call(f'start self_update.bat "{workingDir}"', shell=True)
                            sys.exit()
                else:
                    logger.debug('Update is available but user chooses to not update.')
            else:
                logger.debug('You are running the latest version')
    else:
        logger.debug('Unable to connect to Internet. Skipping update check.')
else:
    logger.debug('Running in benchmark or restore point mode. Skipping update check.')


# Hardware classes (letter C at end to prevent overlapping variable name)
class cpuC:
    def __init__(self):
        status('Initializing CPU')
        brand_code = {'GenuineIntel': 'I', 'AuthenticAMD': 'A'}
        self.cpus = pc.Win32_Processor()
        self.count = len(self.cpus)
        self.names = [cpu.Name.strip() for cpu in self.cpus]
        brandNames = [cpu.Manufacturer.strip() for cpu in self.cpus]
        self.brandCode = [brand_code[brandName] for brandName in brandNames]
        self.thread_no = [cpu.ThreadCount for cpu in self.cpus]


class motherBoardC:
    def __init__(self):
        status('Initializing Motherboard')
        brand_name = {'Micro-Star International Co., Ltd.': 'MSI',
                      'ASUSTeK COMPUTER INC.': 'Asus',
                      'Gigabyte Technology Co., Ltd.': 'Gigabyte',
                      'ASRock': 'ASRock',
                      'Colorful Technology And Development Co.,LTD': 'Colorful'}
        mb = pc.Win32_BaseBoard()[0]  # 0 here assumes only 1 mobo so take first in list
        self.name = mb.Product.strip()
        self.brand = brand_name.get(mb.Manufacturer.strip(), mb.Manufacturer.strip())  # default back to original name
        intel_11_gen_chipsets = ['H410', 'B460', 'H470', 'Q470', 'Z490', 'H510', 'B560', 'H570', 'Q570', 'Z590']
        intel_12_gen_chipsets = ['H610', 'B660', 'H670', 'Q670', 'Z690']
        self.intel_gen = None
        if mb.Manufacturer.strip() not in brand_name.keys():
            logger.warning(f'Motherboard brand detected is "{mb.Manufacturer.strip()}" which is not recognized as one of Asus, MSI, Gigabyte or AsRock. Some motherboard-specific software and drivers may not install properly or may be skipped. If you are sure that it is one of the brands, please contact developer.')
            if mode != 1 and not allow_mode_switch:
                messagebox.showwarning(title='Motherboard not recognized!', message=f'Motherboard brand detected is "{mb.Manufacturer.strip()}" which is not recognized as one of Asus, MSI, Gigabyte or AsRock. Some motherboard-specific software and drivers may not install properly or may be skipped. If you are sure that it is one of the brands, please contact developer.')
        elif 'I' in cpu.brandCode:  # only decide generation if intel CPU and mobo is supported
            if any([chipset in self.name.upper() for chipset in intel_11_gen_chipsets]):
                self.intel_gen = 11
                logger.debug('Intel 11th gen detected')
            elif any([chipset in self.name.upper() for chipset in intel_12_gen_chipsets]):
                self.intel_gen = 12
                logger.debug('Intel 12th gen detected')
            elif not mode:  # only have this warning in driver mode (0)
                logger.warning('Unsupported Intel chipset detected. The software only supports 11th and 12th Intel 400 series and newer chipsets. Intel Chipset driver installation will be skipped.')
                messagebox.showwarning(title='Unsupported Intel chipset', message='Unsupported Intel chipset detected. The software only supports 11th and 12th Intel 400 series and newer chipsets. Intel Chipset driver installation will be skipped.')


class ramC:
    def __init__(self):
        status('Initializing RAM')
        rams = pc.Win32_PhysicalMemory()
        self.count = len(rams)
        self.brand = [ram.Manufacturer.strip() for ram in rams]
        self.size = [b_to_gb(ram.Capacity) for ram in rams]
        self.speed = [ram.Speed for ram in rams]
        self.capacity = round(sum(self.size))

    def __call__(self):
        logger.info(f'Detected {self.count} DRAM sticks. Total capacity: {self.capacity}GB')
        for i in range(self.count):
            logger.info(f'{i} {self.brand[i]} {self.size[i]}GB (Running at {self.speed[i]}Mhz)')


class gpuC:
    def __init__(self):
        status('Initializing GPU')
        wmi_gpus = pc.Win32_VideoController()  # some GPUs without driver will not appear under WMI, so need add extra layer of check below
        pnp_util_gpus = subprocess.check_output('pnputil /enum-devices /problem').decode('GB2312').split('\r\n')[2:-1]
        pnp_util_gpus_grouped, temp = [], []
        for i in pnp_util_gpus:
            if i:
                temp += [i]
            else:
                pnp_util_gpus_grouped += [temp]
                temp = []
        pnp_util_gpus = {i[0].split()[-1]: ' '.join(i[1].split(':')[1:]).strip() for i in pnp_util_gpus_grouped if i[0].split()[-1].startswith('PCI') and ' '.join(i[1].split(':')[1:]).strip() == 'Video Controller'}  # PCIE ID: Caption
        self.count = len(wmi_gpus) + len(pnp_util_gpus)
        self.names = [gpu.Name.strip() for gpu in wmi_gpus] + list(pnp_util_gpus.values())
        self.pnpid = [gpu.pnpdeviceid for gpu in wmi_gpus] + list(pnp_util_gpus.keys())
        vendor_id = {'8086': 'I', '10DE': 'N'}  # If it's not these 2, its definitely AMD. Some AMD device ID follows brand so can't confirm
        brandNameDict = {'I': 'Intel', 'A': 'AMD', 'N': 'Nvidia'}
        self.brandCode = [vendor_id.get(id[8:12], 'A') for id in self.pnpid]  # works without driver
        self.brandName = [brandNameDict[code] for code in self.brandCode]  # gpu.AdapterCompatibility also gives brand name but only after driver installation
        self.igpu = self.if_igpu_exist()
        self.dgpu = self.if_dgpu_exist()
        self.quadro = self.check_quadro() if 'N' in self.brandCode else False  # if there is no nvidia gpu, no need check for Quadro
        if self.count == 1 and 'A' in self.brandCode and self.igpu:  # if there is only 1 AMD GPU and the CPU has APU, it will always be treated as iGPU in self.if_igpu_exist(). However, sometimes when there is both AMD iGPU and AMD dGPU, the iGPU will be disabled, causing the dGPU to be detected as iGPU. The code below handles this edge case.
            if not self.if_amd_apu_active():  # no need else here as self.igpu is already True
                self.igpu = False  # the APU is disabled and the detected GPU is dGPU
                self.dgpu = True
                logger.debug('AMD APU is disabled. The one detected AMD GPU is a discrete GPU.')

    def __call__(self):
        logger.info('ID GPU')
        for i in range(self.count):
            logger.info(f'{i}  {self.names[i]}')
        logger.info(f'System has integrated GPU: {self.igpu}')
        logger.info(f'System has discrete GPU: {self.dgpu}')
        logger.info(f'Is Quadro: {self.quadro}')

    def if_igpu_exist(self):
        if 'I' in self.brandCode:
            return True
        elif 'A' in cpu.brandCode and any([re.search(r'\d\d\d\d[GHUC]|RADEON', name) for name in cpu.names]):
            if 'N' in self.brandCode:  # AMD iGPU + Nvidia GPU might cause iGPU to be disabled, thus need special handling here. Based on assumption that you cant have AMD iGPU + AMD GPU + Nvidia GPU
                if 'A' in self.brandCode:  # AMD iGPU is enabled in this case
                    return True
                else:  # AMD iGPU is disabled
                    return False
            else:
                return True  # if there is no nvidia GPU, AMD iGPU is 100% working
        else:
            return False

    def if_dgpu_exist(self):
        if 'N' in self.brandCode:  # have nvidia
            return True
        elif self.count == 1:  # only 1 gpu
            if self.brandCode[0] == 'I':  # and its Intel
                return False
            elif any([re.search(r'\d\d\d\d[GHUC]|RADEON', name) for name in cpu.names]):
                # AMD desktop G series all have igpu so it has to be igpu only, mobile cpus include 'with radeon graphics'
                return False
            else:
                return True  # single AMD GPU and CPU is not APU
        else:  # multiple gpu but no matter what, >=2 gpu -> at least 1 dGPU, no need take care of AMD iGPU + nvidia GPU here as nvidia is elimated in first if
            return True

    def if_amd_apu_active(self):
        amd_apu_dev_id = ['15D8', '15DD', '1636', '1638', '164C']  # Obtained from radeon 21.10.2\Packages\Drivers\Display\WT6A_INF\U0372545.inf search code names, need to update for future APUs
        return self.pnpid[0][17:21] in amd_apu_dev_id  # can use first item as this function is only ran if there is only 1 GPU detected

    def check_quadro(self):  # assume that for multi-GPU systems, its either all GeForce or all Quadro (Aman said so)
        try:
            valid = [folder for folder in os.listdir(driversFolder) if folder.startswith('nvidia_quadro ') and os.path.isdir(os.path.join(driversFolder, folder))]
        except FileNotFoundError:
            logger.warning('Driver folder missing! The software is unable to detect whether installed GPU is a Quadro, thus will install GeForce drivers!')
            return None
        else:
            if len(valid) != 1:
                logger.error(f'None or multiple folders starting with "nvidia_quadro" exists in the drivers folder! The software is unable to detect whether installed GPU is a Quadro, thus will install GeForce drivers!')
                return None
            else:
                with open(os.path.join(driversFolder, valid[0], 'ListDevices.txt'), 'r') as f:
                    data = [line.strip() for line in f.read().splitlines() if line.startswith('\t')]
                    DevID_to_name = {line[:8]: line.split('"')[-2] for line in data}
                return any([True if id[13:21] in DevID_to_name.keys() else False for id in self.pnpid])


class diskC:
    def __init__(self):
        status('Initializing Storage Devices\n(This may take some time)')
        # type_mapping = {0: 'Unknown',
        #                 1: 'No Root Directory',
        #                 2: 'Removable disk',
        #                 3: 'Local disk',
        #                 4: 'Network Drive',
        #                 5: 'Compact Disc',
        #                 6: 'RAM Disk'}
        drive_type_mapping = {0: 'Unspecified',
                              3: 'HDD',
                              4: 'SSD',
                              5: 'SCM'}

        disks = [disk for disk in pc.Win32_DiskDrive() if disk.InterfaceType.upper() != 'USB']  # ignore USB devices
        # self.disk_pnpids = [disk.PNPDeviceID for disk in disks]
        # self.serials = [disk.SerialNumber for disk in disks]
        self.driveCount = len(disks)
        # self.names = [disk.Name for disk in disks]  # gives \\.\PHYSICALDRIVE0, can get drive ID
        serial_to_name = {}
        for disk in disks:
            if disk.SerialNumber.strip():  # if SN not empty
                serial_to_name[disk.SerialNumber.strip()] = disk.Model.strip()
            else:  # else fallback to powershell
                logger.warning(f'Unable to retrieve serial number for disk: {disk.Model.strip()}. Falling back to use firmware names (might not be accurate)')
                r = subprocess.check_output(['powershell.exe', 'Get-PhysicalDisk | Format-Table -Property FriendlyName, MediaType']).decode('GB2312').split('\r\n')[3:-3]
                self.model = [' '.join(entry.split()[:-1]) for entry in r]
                self.diskType = {' '.join(entry.split()[:-1]): entry.split()[-1] for entry in r}
                break  # if fall back already, then will get every disk from powershell, so break here
        else:  # only runs this part if not break loop
            self.model = [disk.Model.strip() for disk in disks]
            serial_to_disktype = {disk.SerialNumber: disk.MediaType for disk in wmi.WMI(moniker='//./ROOT/Microsoft/Windows/Storage').MSFT_PhysicalDisk() if disk.BusType != 7 and 'XVDD' not in disk.Model.upper()}  # ignore USB devices (BusType 7) and MSFT_PhysicalDisk return xbox game installation generated virtual disks (with name 'XVDD') while WMIC does not, causing the next line to break
            self.diskType = {serial_to_name[serial]: drive_type_mapping[serial_to_disktype[serial]] for serial in serial_to_disktype}
        raw_disks = [name.strip() for name in subprocess.check_output(['powershell.exe', 'Get-Disk | Where-Object PartitionStyle -Eq "RAW" | select FriendlyName']).decode('GB2312').split('\r\n')[3:-3]]
        self.raw_disks = [disk for disk in disks if disk.Model.strip() in raw_disks]
        # self.parts = [disk.Partitions for disk in disks]  # number of partitions on each disk
        self.size = [b_to_gb(disk.Size) for disk in disks]
        self.totalDiskSize = round(sum(self.size), 2)
        partitions = [partition.Dependent for partition in pc.Win32_DiskDriveToDiskPartition() if partition.Antecedent in disks]
        volumes = [vol.Dependent for vol in pc.Win32_LogicalDiskToPartition() if vol.Antecedent in partitions]
        self.letters = [vol.Name for vol in volumes]  # C, D, E etc.
        self.volNames = [vol.VolumeName.strip() if vol.VolumeName else 'Unspecified' for vol in volumes]
        self.volCount = len(volumes)
        self.volSize = [b_to_gb(vol.Size) for vol in volumes]
        self.freeSpace = [b_to_gb(vol.FreeSpace) for vol in volumes]
        self.percentFree = [round(i / j * 100, 2) for i, j in zip(self.freeSpace, self.volSize)]
        # self.type = [type_mapping[vol.DriveType] for vol in volumes]
        self.fs = [vol.FileSystem for vol in volumes]
        self.totalFormattedCapacity = round(sum(self.volSize), 2)
        self.totalFormattedFreeSpace = round(sum(self.freeSpace), 2)

    def __call__(self):
        logger.info(f'Detected {self.volCount} {"volume" if self.volCount == 1 else "volumes"} in {self.driveCount} {"drive" if self.driveCount == 1 else "drives"}. {self.totalFormattedFreeSpace}GB out of {self.totalFormattedCapacity}GB free')
        logger.info('Drive detected:')
        for i in range(self.driveCount):
            logger.info(f'{i} {self.model[i]} ({self.diskType[self.model[i]]}): Total {self.size[i]}GB')
        logger.info('Volume detected:')
        for i in range(self.volCount):
            logger.info(f'{i} {self.letters[i]} {self.volNames[i]} ({self.fs[i]}): {self.freeSpace[i]}GB out of {self.volSize[i]}GB free ({round(100 - self.percentFree[i], 2)}% used)')

    def init_disks(self):
        status('Extending OS Drive partition')
        exec_live_output('OS Drive partition extension', subprocess.Popen(f'powershell -Command {get_res_path("extend_osdrive.ps1")}', stdout=subprocess.PIPE, stderr=subprocess.PIPE))
        status('Initializing and formatting new disks')
        if self.raw_disks:
            logger.info(f'Detected uninitialized disks: {", ".join([d.model for d in self.raw_disks])}')
            try:
                r = subprocess.check_output(['powershell.exe', 'Get-Disk | Where-Object PartitionStyle -Eq "RAW" | Initialize-Disk -PassThru | New-Partition -AssignDriveLetter -UseMaximumSize | Format-Volume -FileSystem NTFS -Confirm:$false']).decode('GB2312')
                disks_formatted = [i.split()[0] + ':' for i in r.split('\r\n')[3:-3]]  # returns volume letters
            except subprocess.CalledProcessError as e:
                logger.error(f'An error has occurred when initializing and formatting new disks. Please try to do it manually. The Disk Management window is opened for you. Detailed error: {e}')
                subprocess.call('start diskmgmt.msc', shell=True)
                item_failed.add('Disk initialization and formatting')
            else:
                self.__init__()  # refresh
                logger.info(f'Initialized and formatted {len(disks_formatted)} partitions: {", ".join(disks_formatted)}')
                if self.raw_disks:
                    logger.warning(f'There are still some drives detected as uninitialized: {[d.model for d in self.raw_disks]}')
                if disks_formatted:  # if no disks are newly formatting, do not create the file
                    with open(os.path.join(desktop, 'disks_to_burnin_test.txt'), 'w+') as f:
                        f.write(' '.join(disks_formatted))
                        logger.info(f'The following partitions will be BurnIn tested: system partition and {disks_formatted}')
        else:
            logger.info('All drives are already initialized and formatted.')
        status('Ready')

    @staticmethod
    def get_burnin_disks():
        try:
            with open(os.path.join(desktop, 'disks_to_burnin_test.txt'), 'r') as f:
                disks_to_test = [d.lower() for d in f.read().split()]
        except FileNotFoundError:  # if no new drive is selected
            disks_to_test = []
        with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State') as key:
            if reg.QueryValueEx(key, 'ImageState')[0] != 'IMAGE_STATE_COMPLETE':  # detect whether in Audit mode
                disks_to_test += [os.getenv('SystemDrive')]  # test boot drive if system is in Audit mode because if customer provide own OS drive then we should not test it
        return set(disks_to_test)


class tempsC:
    def __init__(self):
        status('Initializing HWINFO')
        self.hwinfoAvail = False
        self.hwinfopath = 'HWINFO'
        self.exeName = 'HWiNFO64.exe'
        self.logging = False
        self.logfile = None
        self.data = None
        self.parse_error_shown = False
        subprocess.run(f"taskkill /im {self.exeName} /f /t", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        delete_sub_key(reg.HKEY_CURRENT_USER, r'Software\HWiNFO64', task='HWINFO', silent=True)  # clean up registry in case corrupted setting causing sensor read error
        with reg.CreateKeyEx(reg.HKEY_CURRENT_USER, r'Software\HWiNFO64\Sensors', 0, reg.KEY_ALL_ACCESS) as key:
            reg.SetValueEx(key, 'AutoLoggingHotKey', 0, reg.REG_DWORD, 0x00030053)  # ctrl+alt+s
            reg.FlushKey(key)
        if os.path.isfile(os.path.join(self.hwinfopath, self.exeName)):
            status('Launching HWINFO')
            shutil.copy(get_res_path(self.exeName[:-4] + ".INI"), os.path.join(self.hwinfopath, self.exeName[:-4] + ".INI"))
            logger.debug('Cleaning up conflicting HWINFO log files')
            for file in [file for file in os.listdir(self.hwinfopath) if os.path.splitext(file)[1].lower() == '.csv' and file.startswith('HWiNFO_LOG_')]:  # not using temps.tempfiles as clean up may not remove manually generated csvs which interfere with determination of latest log file
                try:
                    os.remove(os.path.join(self.hwinfopath, file))
                except OSError as e:
                    logger.warning(f'Failed to delete {file} due to error: {e}. This may cause HWINFO logging to not work properly!')
            subprocess.call(f'cd {self.hwinfopath} &'f'start {self.exeName} &''cd ..', shell=True)  # change launch dir so that log files are in sub folder
            while not if_running(self.exeName):  # wait till it started
                root.update()  # prevent freezing
            logger.info('HWINFO launched')
            self.hwinfoAvail = True
        else:  # allow pop up in all modes since it is crucial for mode 1 to work
            if order_no: bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: HWINFO not found.\nTester notes: {note}')
            messagebox.showerror(title='HWINFO not found', message=f'WARNING: HWINFO executable is not found at {self.hwinfopath}/{self.exeName}, please check the folder where the QC software is placed in and the name of the HWINFO exe is {self.exeName}. If you continue, the software may crash.')
            logger.warning(f'WARNING: HWINFO executable is not found at {self.hwinfopath}/{self.exeName}, please check the folder where the QC software is placed in and the name of the HWINFO exe is {self.exeName}. If you continue, the software may crash.')

    def start(self):
        if not self.hwinfoAvail:
            if order_no: bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: HWINFO not available\nTester notes: {note}')
            messagebox.showerror(title='HWINFO is not available!', message='HWINFO is not available. Please check log for details.')
            logger.warning('Attempt to start temps logging when HWINFO is not available. Please check log for details.')
        else:
            if self.logging:
                logger.info('Attempt to start HWINFO logging when it is already logging')
                return
            status('Starting HWINFO logging')
            self.dirFileCount = len(list(os.listdir(self.hwinfopath)))
            i = 0
            while len(list(os.listdir(self.hwinfopath))) != self.dirFileCount + 1:  # wait for the csv file to be created
                keyboard.press_and_release('ctrl+alt+s')
                i += 1
                if i > 10: logger.debug('Please press ctrl+alt+s on keyboard (ignore this message if it only appears for a few times, else please report to developer)')
                root.update()
                sleep(0.5)
            self.logging = True
            logger.info('HWINFO logging STARTED')
            self.tempfiles = [file for file in os.listdir(self.hwinfopath) if os.path.splitext(file)[1].lower() == '.csv' and file.startswith('HWiNFO_LOG_')]
            if len(self.tempfiles) != 1:
                tempfiles_digits = [int(file[11:-4]) for file in self.tempfiles]
                self.logfile = self.tempfiles[tempfiles_digits.index(max(tempfiles_digits))]  # pick latest file
                logger.info(f'{len(self.tempfiles)} CSV files detected under {self.hwinfopath} folder. Reading temperature from the latest file ({self.hwinfopath}/{self.logfile}).')
            else:
                self.logfile = self.tempfiles[0]
                logger.info(f'Parsing temperature from {self.hwinfopath}/{self.logfile}')
            status('Ready')

    def stop(self, kill=False, log=True):
        if self.logging:
            keyboard.press_and_release('ctrl+alt+s')
            self.logging = False
            if log: logger.info('HWINFO Logging STOPPED')
        else:
            if log: logger.info('HWINFO Logging was not running thus no need to stop.')
        if kill:
            if log: logger.info(f'Killing {self.exeName}')
            subprocess.run(f"taskkill /im {self.exeName} /f /t", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def read(self, verbose=True):
        if not self.hwinfoAvail:
            logger.warning('Attempt to read temps data when HWINFO is not available. Please check log for details.')
        else:
            if verbose: logger.debug(f'Reading temp csv file {self.hwinfopath}/{self.logfile}')
            if self.logfile:
                with open(os.path.join(self.hwinfopath, self.logfile), 'r', newline='') as f:
                    lines = f.read()
                if lines.find('°') == -1:  # when default encoding fails, use lib to guess (may also fail but it will be caught later)
                    logger.debug('Failed to parse file with system default encoding. Switched to encoding guessing mode')
                    import charset_normalizer
                    lines = str(charset_normalizer.from_path(os.path.join(self.hwinfopath, self.logfile)).best())
                csv_file = list(csv.reader([line.rstrip(',') for line in lines.replace('°', '').splitlines()]))  # rstrip remove extra column that HWINFO creates
                columns = csv_file[0]
                if not self.logging: csv_file = csv_file[:-2]  # remove bottom 2 row of columns for finished logs
                if len(csv_file) == 1:
                    logger.warning('Attempt to read temp data when it is still empty. The log file may be corrupted or just created such that no readings have been recorded yet. This is NOT a fatal or permanent error and you can try again later.')
                else:
                    lines = [line[:len(columns)] for line in csv_file[1:]]  # skip first row column names, only read up to columns that were present in first row
                    self.data = pd.DataFrame(lines, columns=columns)
                    cpu_col_names = {'I': ['CPU Package [C]', 'CPU Package Power [W]'], 'A': ['CPU (Tctl/Tdie) [C]', 'CPU Package Power [W]']}
                    if not gpu.dgpu:  # ignore gpu temp since igpu are not monitored
                        col_names = ['CPU Temp', 'CPU Power']
                    else:
                        if not gpu.igpu and gpu.dgpu:  # no igpu and only have dgpu -> take all GPU temp columns
                            gpu_data = self.data.filter(like='GPU Temperature [C]')
                            if len(gpu_data.columns) != gpu.count:
                                logger.warning(f'Please report this warning to developer together with system specs: GPU Temps column count mismatch: detected dGPU only thus expected GPU Temp columns to be same as no of GPUs. Got no of columns: {len(gpu_data.columns)}, no of GPUs: {gpu.count}')
                        else:  # both igpu and dgpu exist
                            if 'I' in cpu.brandCode:  # intel gpu don't have separate HWINFO section
                                gpu_data = self.data.filter(like='GPU Temperature [C]')
                                if len(gpu_data.columns) != gpu.count - 1:  # minus away the igpu
                                    logger.warning(f'Please report this warning to developer together with system specs: GPU Temps column count mismatch: detected Intel iGPU + {gpu.count - 1} dGPU thus expected GPU Temp columns to be no of GPUs-1. Got no of columns: {len(gpu_data.columns)}, no of GPUs: {gpu.count}')
                            else:  # AMD iGPU
                                gpu_data = self.data.filter(like='GPU Temperature [C]').iloc[:, 1:]  # AMD iGPU report itself in a new section, so need remove first col
                                if len(gpu_data.columns) != gpu.count - 1:  # minus away the igpu
                                    logger.warning(f'Please report this warning to developer together with system specs: GPU Temps column count mismatch: detected AMD iGPU + {gpu.count - 1} dGPU thus expected GPU Temp columns to be no of GPUs-1. Got no of columns: {len(gpu_data.columns)}, no of GPUs: {gpu.count}')
                        col_names = ['CPU Temp', 'CPU Power'] + [f'GPU Temp{i + 1}' for i in range(len(gpu_data.columns))]  # do not use gpu.count here to ensure stability, +1 to start from 1 for readability
                    try:
                        cpu_data = self.data[cpu_col_names[cpu.brandCode[0]]]  # use first CPU code since there is no way to mix Intel and AMD CPUs
                    except KeyError as e:
                        if not self.parse_error_shown:
                            logger.error(f'Error parsing temperature log, one or more of the required data is not recorded or may be corrupted. Please manually check the log csv file ({self.hwinfopath}/{self.logfile}).\nDetailed error:{e}')
                            self.parse_error_shown = True
                    else:
                        cpu_data = cpu_data.loc[:, ~cpu_data.columns.duplicated(keep='last')]  # keep last as it is usually the 'enhanced' sensor value
                        if not gpu.dgpu:
                            self.data = cpu_data
                        else:
                            self.data = pd.concat([cpu_data, gpu_data], axis=1)
                        self.data.columns = col_names
                        self.data = self.data.astype(float)
            else:
                logger.error('Attempt to read temps when HWINFO is not logging (may have detailed reasons in log above).')
        return self.data

    def maxTemp(self, part):
        if self.read(verbose=False) is not None:
            if part in list(self.data):
                return round(max(self.data[part].tolist()))
            else:
                logger.error(f'Attempted to get max temp of an unknown part. Got: {part}')

    def minTemp(self, part):
        if self.read(verbose=False) is not None:
            if part in list(self.data):
                return round(min(self.data[part].tolist()))
            else:
                logger.error(f'Attempted to get min temp of an unknown part. Got: {part}')

    def avgTemp(self, part):
        if self.read(verbose=False) is not None:
            if part in list(self.data):
                from statistics import mean
                return round(mean(self.data[part].tolist()))
            else:
                logger.error(f'Attempted to get avg temp of an unknown part. Got: {part}')

    def plot_temps(self):
        temp_plot_button['state'] = DISABLED
        power_plot_button['state'] = DISABLED
        root.attributes("-topmost", False)
        t = list(range(1, self.data.shape[0] + 1, 1))
        cpu_temp = self.data['CPU Temp'].tolist()
        plt.plot(t, cpu_temp, label='CPU Temp')
        if gpu.dgpu:
            for i in self.data.columns[2:]:
                plt.plot(t, self.data[i].tolist(), label=i)
        plt.plot(t, [80.0] * len(t), 'r:', label='80 degrees')
        plt.xlabel('Time (s)')
        plt.ylabel('Temps (C)')
        plt.ylim(20, 100)
        plt.xscale('linear')
        plt.yscale('linear')
        plt.title('Hardware Temperatures')
        plt.legend()
        plt.show(block=True)
        root.attributes("-topmost", True)
        temp_plot_button['state'] = NORMAL
        power_plot_button['state'] = NORMAL

    def plot_cpu_pwr(self):
        temp_plot_button['state'] = DISABLED
        power_plot_button['state'] = DISABLED
        root.attributes("-topmost", False)
        t = list(range(1, self.data.shape[0] + 1, 1))
        cpu_power = self.data['CPU Power'].tolist()
        plt.plot(t, cpu_power, label='CPU Power')
        plt.xlabel('Time (s)')
        plt.ylabel('Power (W)')
        plt.xscale('linear')
        plt.yscale('linear')
        plt.title('CPU Power')
        plt.legend()
        plt.show(block=True)
        root.attributes("-topmost", True)
        temp_plot_button['state'] = NORMAL
        power_plot_button['state'] = NORMAL


# class processC:
#     def __init__(self):
#         status('Initializing Processes')
#         self.processesIter = psutil.process_iter()
#
#     def refresh(self):  # just to refresh
#         self.processesIter = psutil.process_iter()
#
#     def if_running(self, process):
#         self.refresh()
#         return process.lower() in [p.name().lower() for p in self.processesIter]


class networkC:
    def __init__(self, post_driver=False):
        status('Initializing Network Adaptors')
        net_devices = subprocess.check_output('pnputil /enum-devices /class Net').decode('GB2312').split('\r\n')[2:-1]
        net_devices_grouped, temp = [], []
        for i in net_devices:  # Output from PnP Util is different in lines per device, so we need to split them into sublist by the space in middle
            if i:
                temp += [i]
            else:
                net_devices_grouped += [temp]
                temp = []  # reset sublist to empty
        self.net_devices = {i[0].split()[-1]: ' '.join(i[1].split(':')[1:]).strip() for i in net_devices_grouped if i[0].split()[-1].startswith('PCI')}  # PCIE ID: Caption

        net_devices_no_driver = subprocess.check_output('pnputil /enum-devices /problem').decode('GB2312').split('\r\n')[2:-1]  # cannot have '/class Net' here as no driver installed may not classify correctly
        net_devices_no_driver_grouped, temp = [], []
        for i in net_devices_no_driver:
            if i:
                temp += [i]
            else:
                net_devices_no_driver_grouped += [temp]
                temp = []
        self.net_devices_no_driver = {i[0].split()[-1]: ' '.join(i[1].split(':')[1:]).strip() for i in net_devices_no_driver_grouped if i[0].split()[-1].startswith('PCI') and ' '.join(i[1].split(':')[1:]).strip() in ['Network Controller', 'Ethernet Controller']}  # PCIE ID: Caption

        # WiFi devices processing
        self.wifi_devices_no_driver = {i: v for i, v in self.net_devices_no_driver.items() if v == 'Network Controller'}
        if post_driver and self.wifi_devices_no_driver:  # after driver phase, there should be no more wifi adaptors with name Network Controller
            logger.warning(f'There are still {len(self.wifi_devices_no_driver)} WiFi adaptors that have no driver installed. Please install the driver for them.')
        try:
            wifi_adaptors = subprocess.check_output(r'WMIC /NameSpace:\\root\WMI Path MSNdis_PhysicalMediumType where "NdisPhysicalMediumType=1 or NdisPhysicalMediumType=8 or NdisPhysicalMediumType=9"', stderr=subprocess.STDOUT).decode('GB2312').split('\r\r\n')  # mappings: (Get-NetAdapter | Get-Member PhysicalMediaType).Definition
            wifi_adaptors = [' '.join(device.split()[1:-1]) for device in wifi_adaptors[1:-2] if device]  # some device might be empty, causing bug in later if/else, so need another layer of check here, same for LAN adaptor below
        except subprocess.CalledProcessError:  # may return error if no wifi adaptor found (due to lack of driver)
            physical_wifi_devices = self.wifi_devices_no_driver.copy()  # so in this case will just use no driver list as all wifi adaptor should have no driver
        else:
            physical_wifi_devices = {i: v for i, v in self.net_devices.items() if v in wifi_adaptors}
            physical_wifi_devices.update(self.wifi_devices_no_driver)  # merge both have driver and no driver wifi devices as sometimes WMIC won't return error even if no wifi detected, and other times there may be a mixture of driver'ed and driverless wifi devices
        self.physical_wifi_adaptors = list(physical_wifi_devices.values())

        # LAN devices processing
        self.LAN_devices_no_driver = {i: v for i, v in self.net_devices_no_driver.items() if v == 'Ethernet Controller'}
        if post_driver and self.LAN_devices_no_driver:  # after driver phase, there should be no more wifi adaptors with name Network Controller
            logger.warning(f'There are still {len(self.LAN_devices_no_driver)} LAN adaptors that have no driver installed. Please install the driver for them.')
        try:
            LAN_adaptors = subprocess.check_output(r'WMIC /NameSpace:\\root\WMI Path MSNdis_PhysicalMediumType where "NdisPhysicalMediumType=0 or NdisPhysicalMediumType=17 or NdisPhysicalMediumType=18"', stderr=subprocess.STDOUT).decode('GB2312').split('\r\r\n')
            LAN_adaptors = [' '.join(device.split()[1:-1]) for device in LAN_adaptors[1:-2] if device]
        except subprocess.CalledProcessError:  # may return error if no LAN adaptor found (due to lack of driver)
            physical_LAN_devices = self.LAN_devices_no_driver.copy()  # so in this case will just use no driver list as all LAN adaptor should have no driver
        else:
            physical_LAN_devices = {i: v for i, v in self.net_devices.items() if v in LAN_adaptors}
            physical_LAN_devices.update(self.LAN_devices_no_driver)  # merge both have driver and no driver LAN devices, same reasons above
        self.physical_LAN_adaptors = list(physical_LAN_devices.values())

        self.vendor_id_dict = {'8086': 'Intel', '10EC': 'Realtek', '1D6A': 'Marvell(Aquantia)', '14C3': 'MediaTek'}
        self.wifi_vendorids = [pnpid[8:12] for pnpid in physical_wifi_devices.keys()]
        self.LAN_vendorids = [pnpid[8:12] for pnpid in physical_LAN_devices.keys()]
        self.profile_path = [get_res_path(path) for path in ['dreamcore_wifi5.xml', 'dreamcore_wifi.xml', 'dreamcore_wifi24.xml']]
        self.ssid = ['Dreamcore 5.0GHz', 'Dreamcore', 'Dreamcore 2.4GHz']  # must align with profile path sequence
        self.multiple_wifi_adaptor = len(self.physical_wifi_adaptors) > 1
        self.skip_wifi_test = False

        if not self.physical_LAN_adaptors:
            messagebox.showwarning(title='No LAN adaptor found', message=f'No LAN adaptor detected!')  # need to pop up as it is not normal to not have a LAN adaptor

        if self.multiple_wifi_adaptor and not mode:  # only need this in mode 0 where driver installation happens
            if order_no: bot_send_msg(f'Order {order_no} testing by {qc_person}\nWarning: multiple wifi adaptors detected, skipping WiFi test, but WiFi drivers will still install for all supported adaptors.\nTester notes: {note}')
            messagebox.showwarning(title='Multiple WiFi adaptors detected', message=f'{len(self.physical_wifi_adaptors)} WiFi adaptors detected. Only one is supported now. Skipping WiFi test, but WiFi drivers will still install for all supported adaptors.')
            logger.warning(f'{len(self.physical_wifi_adaptors)}) WiFi adaptors detected. Only one is supported now. Skipping WiFi test, but WiFi drivers will still install for all supported adaptors.')
            item_failed.add(f'WiFi test: multiple ({len(self.physical_wifi_adaptors)}) WiFi adaptors detected')
            self.skip_wifi_test = True

    def wifi_test(self):
        status('Preparing WiFi test')
        adaptor_status = subprocess.check_output('wmic nic get name,NetEnabled').decode('GB2312').split('\r\r\n')
        adaptor_status = {' '.join(device.split()[:-1]): True if device.split()[-1].upper() == 'TRUE' else False for device in adaptor_status[1:-2]}
        to_disable = [adaptor for adaptor in adaptor_status if adaptor_status[adaptor] and adaptor not in self.physical_wifi_adaptors]
        for adaptor in to_disable:
            logger.debug(f'Disabling network adaptor: {adaptor} (This is needed for WiFi testing, they will be enabled later)')
            try:
                subprocess.check_output(f'wmic path win32_networkadapter where name="{adaptor}" call disable')  # use check_output here to mute stdout from WMIC. It will not likely to return non 0 exit code (tested disabling an already disabled adaptor returns 0)
            except subprocess.CalledProcessError:
                logger.warning(f'Failed to disable network adaptor: {adaptor}')
            else:
                logger.debug(f'Disabled {adaptor}')
            root.update()
        for i in range(len(self.profile_path)):
            self.load_profile(self.ssid[i], self.profile_path[i])
        if self.skip_wifi_test:  # not adding to failed item list here as it is already added when respective reasons are detected
            logger.warning('WiFi test skipped. Detailed reason can be found in logs above.')
        else:
            for ssid in self.ssid:
                if self.connect(ssid): break  # Per dreamcore permission, no need to perform internet test after successfully connecting to WiFi
            else:
                item_failed.add(f'WiFi: Failed to connect to all of these WiFi: {self.ssid}')
            for adaptor in to_disable:
                logger.debug(f'Enabled network adaptor: {adaptor}')
                try:
                    subprocess.check_output(f'wmic path win32_networkadapter where name="{adaptor}" call enable')  # use check_output for same reason as above
                except subprocess.CalledProcessError:
                    logger.warning(f'Failed to enable network adaptor: {adaptor}')
                root.update()

    def load_profile(self, ssid, profile_path):
        status('Loading WiFi profile')
        logger.debug(f'Loading WiFi profile from {profile_path}')
        r = subprocess.call(fr'netsh wlan add profile filename="{profile_path}" user=all')  # will return non-0 if profile already exist, so need check again below
        if r:
            if subprocess.run(f'netsh wlan show profiles name="{ssid}"', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL):
                logger.warning(f'An error occurred when adding WiFi profile from {profile_path}. Return code: {r}')
                self.skip_wifi_test = True
                item_failed.add(f'WiFi: error loading WiFi profile')
            else:
                logger.info('WiFi profile loaded successfully')
        else:
            logger.info('WiFi profile loaded successfully')

    def connect(self, ssid):
        status(f'Connecting to WiFi: {ssid}')
        for i in range(5):
            r = subprocess.call(f'netsh wlan connect name="{ssid}" ssid="{ssid}"')  # success return 0 (false), fail return 1. Only works if there is only ONE valid wifi adaptor.
            if r:  # in some case it connects successfully but still return non 0 so we need check again
                connected_list = [i.split() for i in subprocess.check_output(f'netsh wlan show interface').decode('GB2312').split('\r\n')[3:-4] if i.split()[0].upper() == 'SSID']
                if connected_list and ' '.join(connected_list[0][2:]) == ssid:  # connected list will be empty if no wifi connected and thus will give index error if we try to slice it
                    logger.info('WiFi connected successfully')
                    break
                else:
                    logger.warning(f'Attempt {i + 1}: An error occurred when connecting to WiFi: {ssid}. Retrying...')
                    sleep(2)
            else:
                sleep(3)  # wait for it to connect
                logger.info(f'Connected to WiFi: {ssid}')
                status('Ready')
                break
            root.update()
        else:  # only failed connection here as successful will break which causes it to skip else block of for loop
            logger.error(f'Unable to connect to WiFi: {ssid} after 5 tries. There may be logs above detailing the error.')
            return False
        return True


class bluetoothC:
    def __init__(self):
        status('Initializing Bluetooth')
        self.passed = False
        self.timeout = 5.0

    def test(self):
        status('Bluetooth scanning')
        logger.info(f'Bluetooth scanning test started with timeout of {self.timeout} seconds.')

        async def run():
            if await BleakScanner.find_device_by_filter(lambda d, ad: True, timeout=self.timeout) is not None:
                self.passed = True

        import asyncio
        loop = asyncio.get_event_loop()
        try:
            loop.run_until_complete(run())
        except Exception as e:
            logger.error(f'Bluetooth scanning test failed due to error: {e}')
            self.passed = False

        if self.passed:
            logger.info(f'{Fore.GREEN}Bluetooth scanning test passed.{Fore.RESET}')
        else:
            logger.warning('Bluetooth scanning test failed as it found no device or encountered fatal error (logs above may record detailed error).')
            item_failed.add('Bluetooth scanning test')
        status('Ready')


def progressbar_step(progressbar, init=False, inputs=False, total_steps=1):
    if init:
        if inputs:
            total_steps = 12
        elif mode == 1:  # benchmark mode
            total_steps = 8
        elif mode == 2:  # OOB mode no need anything else
            total_steps = 6
        else:
            total_steps = 7
    progressbar['value'] += (100.0 / total_steps)
    root.update()


def gatherInfo(inputs=True, specs=True, post_driver=False):
    global os_name, cpu, mobo, ram, gpu, disk, temps, process, network, driver, bt, test_usb, _3dmark, SID, order_no_input, qc_person_input, note_input, inaccurate_spec_label, specs_label_frame, info_input_frame, submit_button
    if inputs: start_button.destroy()
    init_progressbar = Progressbar(frame, orient=HORIZONTAL, mode='determinate', length=600)
    init_progressbar.pack(expand=True, fill=BOTH, side=TOP)
    status('Gathering OS info')
    os_info = pc.Win32_OperatingSystem()[0]
    SID = subprocess.check_output('whoami /user').split()[-1].decode('GB2312')
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    cpu = cpuC()
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    mobo = motherBoardC()
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    ram = ramC()
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    gpu = gpuC()
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    disk = diskC()
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    network = networkC(post_driver=post_driver)
    progressbar_step(init_progressbar, init=True, inputs=inputs)
    if inputs:  # if launching at start then initialize these classes since they are not needed for specs display
        bt = bluetoothC()
        progressbar_step(init_progressbar, init=True, inputs=inputs)
        # process = processC()
        # progressbar_step(init_progressbar, init=True, inputs=inputs)
        temps = tempsC()
        progressbar_step(init_progressbar, init=True, inputs=inputs)
        driver = driversC()  # must be init after all hardware classes
        progressbar_step(init_progressbar, init=True, inputs=inputs)
        test_usb = usb_test()
        progressbar_step(init_progressbar, init=True, inputs=inputs)
        _3dmark = _3dmarkC()
        progressbar_step(init_progressbar, init=True, inputs=inputs)
    elif mode == 1:  # benchmark mode, input should be False
        # process = processC()
        # progressbar_step(init_progressbar, init=True, inputs=inputs)
        temps = tempsC()
        progressbar_step(init_progressbar, init=True, inputs=inputs)
        _3dmark = _3dmarkC()
        progressbar_step(init_progressbar, init=True, inputs=inputs)
    os_name = os_info.Caption
    os_version = ' '.join([os_info.Version, 'Build', os_info.BuildNumber])

    logger.info('Specifications:')
    logger.info(f'CPU ({cpu.count} detected): {", ".join(cpu.names)}')
    logger.info(f'Motherboard: {mobo.brand} {mobo.name}')
    logger.info(f'RAM ({ram.count} detected): {ram.capacity}GB')
    logger.info(f'GPU ({gpu.count} detected): {", ".join(gpu.names)}')
    logger.info(f'Storage: Total {disk.totalDiskSize}GB, formatted {disk.totalFormattedCapacity}GB')
    logger.info(f'OS: {os_name} {os_version}')
    print()
    ram()
    print()
    disk()
    print()
    logger.info(f'CPU Brand Code: {cpu.brandCode}')
    logger.info(f'GPU Brand Code: {gpu.brandCode}')
    print()
    gpu()

    if not mode:
        print()
        if network.physical_wifi_adaptors:
            logger.info(f'WiFi Adaptor: {network.physical_wifi_adaptors}, brand: {[network.vendor_id_dict[pnpid] for pnpid in network.wifi_vendorids]}')
        else:
            logger.info('No WiFi adaptor detected')
        if network.physical_LAN_adaptors:
            logger.info(f'LAN Adaptor: {network.physical_LAN_adaptors}, brand: {[network.vendor_id_dict[pnpid] for pnpid in network.LAN_vendorids]}')
        else:
            logger.warning('No LAN adaptor detected')
    if specs:
        specs_label_frame = LabelFrame(frame, text='Specifications', borderwidth=3, padding=(10, 10))
        specs_label_frame.pack(expand=True, fill=BOTH, side=TOP)
        Label(specs_label_frame, text=f'CPU ({cpu.count} detected): {", ".join(cpu.names)}').pack()
        Label(specs_label_frame, text=f'Motherboard: {mobo.brand} {mobo.name}').pack()
        Label(specs_label_frame, text=f'RAM ({ram.count} detected): {ram.capacity}GB running at {ram.speed[0]}Mhz').pack()
        Label(specs_label_frame, text=f'GPU ({gpu.count} detected): {", ".join(gpu.names)} (Brand: {", ".join(gpu.brandName)})').pack()
        Label(specs_label_frame, text=f'Storage: Total {disk.totalDiskSize}GB, formatted {disk.totalFormattedCapacity}GB').pack()
        Label(specs_label_frame, text=f'OS: {os_name} {os_version}').pack()
    if inputs:
        inaccurate_spec_label = Label(frame, text=f'Information above may not be accurate if drivers are not installed yet.\nThese information will be displayed again at the end of driver installation.\n')
        inaccurate_spec_label.pack(expand=True, fill=BOTH, side=TOP)
        info_input_frame = Frame(frame, padding=(10, 0))
        info_input_frame.pack(expand=True, fill=BOTH, side=TOP)
        labels_frame = Frame(info_input_frame, padding=(5, 0))
        labels_frame.pack(side=LEFT)
        inputs_frame = Frame(info_input_frame, padding=(5, 0))
        inputs_frame.pack(expand=True, fill=BOTH, side=RIGHT)
        Label(labels_frame, text='Order No: ').pack(fill=BOTH, side=TOP)
        order_no_input = Entry(inputs_frame, font=tk_default_font)
        order_no_input.pack(expand=True, fill=BOTH, side=TOP)
        order_no_input.bind('<Return>', lambda event: qc_person_input.focus_set())
        order_no_input.bind('<Down>', lambda event: qc_person_input.focus_set())
        Label(labels_frame, text='Tester\'s Name: ').pack(fill=BOTH, side=TOP)
        qc_person_input = Entry(inputs_frame, font=tk_default_font)
        qc_person_input.pack(expand=True, fill=BOTH, side=TOP)
        qc_person_input.bind('<Return>', lambda event: note_input.focus_set())
        qc_person_input.bind('<Up>', lambda event: order_no_input.focus_set())
        qc_person_input.bind('<Down>', lambda event: note_input.focus_set())
        Label(labels_frame, text='Notes (Optional): ').pack(fill=BOTH, side=TOP)
        note_input = Entry(inputs_frame, font=tk_default_font)
        note_input.pack(expand=True, fill=BOTH, side=TOP)
        note_input.bind('<Return>', lambda event: logTestingInfo())
        note_input.bind('<Up>', lambda event: qc_person_input.focus_set())
        note_input.delete(0, END)
        note_input.insert(0, f'{mobo.brand} {mobo.name}')
        submit_button = Button(frame, text='Submit', command=logTestingInfo)
        submit_button.pack(expand=True, fill=BOTH, side=TOP, pady=(15, 0))
        order_no_input.focus_set()

    init_progressbar.destroy()
    ver_msg.destroy()
    if inputs:
        status('Awaiting user input')
    else:
        status('Ready')


def logTestingInfo():
    global order_no, qc_person, note, begin_button
    order_no = order_no_input.get().strip()
    if not order_no:
        messagebox.showerror(title='Empty order no!', message='Please fill in the order no')
    # elif not all(c in '0123456789-' for c in order_no):
    #     messagebox.showerror(title='Wrong order no format!', message='Order no only accepts dashes, space and numbers. If there is any special note, please record in the notes box below.')
    else:
        qc_person = qc_person_input.get().strip()
        if not qc_person:
            messagebox.showerror(title='Empty QC person!', message='Please fill in the QC tester\'s name')
        else:
            logger.info(f'{order_no} testing by {qc_person}')
            note = note_input.get().strip()
            if note:
                logger.info(f'Tester notes: {note}')
            else:
                note = 'NIL'
            with open(os.path.join(os.environ["ProgramData"], 'tester.txt'), 'w+') as f:
                f.write(f'{qc_person}\n{order_no}\n{note}')
            inaccurate_spec_label.destroy()
            specs_label_frame.destroy()
            info_input_frame.destroy()
            submit_button.destroy()
            begin_button = Button(frame, text='Begin', command=init_qc, width=30)
            begin_button.pack(expand=True, fill=BOTH, side=TOP)
            begin_button.bind('<Return>', lambda event: init_qc())
            begin_button.focus_set()
            testing_text['text'] = f'Order #{order_no} testing by {qc_person}'


def init_qc():
    bot_send_msg(f'Order {order_no} testing by {qc_person}\nQC Started (installing drivers).\nTester notes: {note}')
    global allow_mode_switch, audio_button, usb_button, usbtest_frame, itemfailed_frame, item_failed_index, driver_progressbar
    begin_button.destroy()
    root.update()
    logger.info('Started driver installation')
    allow_mode_switch = True
    driver_phase_label = Label(frame, text='Installing drivers...')
    driver_phase_label.pack(expand=True, fill=BOTH, side=TOP)
    driver_progressbar = Progressbar(frame, orient=HORIZONTAL, mode='determinate', length=600)
    driver_progressbar.pack(expand=True, fill=BOTH, side=TOP)
    # Driver phase
    if network.physical_wifi_adaptors:  # no need check for multiple WiFi adaptors as driver can be installed for multiple
        driver.install_wifi_bt_driver()
    progressbar_step(driver_progressbar, total_steps=6)
    driver.install_mobo_driver()
    progressbar_step(driver_progressbar, total_steps=6)
    driver.install_lan_driver()
    progressbar_step(driver_progressbar, total_steps=6)
    driver.install_gpu_driver()
    progressbar_step(driver_progressbar, total_steps=6)
    driver.install_audio_driver()
    progressbar_step(driver_progressbar, total_steps=6)
    driver.checkDevices()
    progressbar_step(driver_progressbar, total_steps=6)
    driver_progressbar.destroy()
    driver_phase_label.destroy()

    test_phase_label = Label(frame, text='Setting up and testing hardware')
    test_phase_label.pack(expand=True, fill=BOTH, side=TOP)
    _3dmark_install_thread = threading.Thread(target=_3dmark.install)
    _3dmark_install_thread.start()
    if not time_synced: sync_time()
    logger.info('Setting power plan to High Performance')
    r = subprocess.call('powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c')
    if r:
        logger.error(f'Failed to set power plan to High Performance. Return code: {r}')
        item_failed.add('Setting power plan to High Performance')
    else:
        logger.info('Power plan set to High Performance')
    disk.init_disks()
    if network.physical_wifi_adaptors:
        if not network.multiple_wifi_adaptor:
            if driver.skip_wifi:
                logger.warning('WiFi test skipped as WiFi driver installation was skipped or failed. Detailed reason can be found in logs above.')
                item_failed.add('WiFi: skipped test as driver installation was skipped or failed')
            else:
                network.wifi_test()  # internally there is another skip_wifi_test check
        else:
            logger.warning('WiFi test skipped mas multiple WiFi adaptors detected')
        if driver.skip_bt:
            logger.warning('Bluetooth test skipped as bluetooth driver installation was skipped or failed. Detailed reason can be found in logs above.')
            item_failed.add('Bluetooth: skipped test as driver installation was skipped or failed')
        else:
            bt.test()

    windows_update()  # after networking tests to prevent possible interruptions
    test_phase_label.destroy()
    gatherInfo(inputs=False, specs=True, post_driver=True)  # Refresh specs for QC to confirm again
    iotest_frame = LabelFrame(frame, text='IO Tests', borderwidth=3, padding=(10, 10))
    iotest_frame.pack(expand=True, fill=BOTH, side=TOP)
    usbtest_frame = LabelFrame(iotest_frame, text='USB Test', borderwidth=3, padding=(20, 20))
    usbtest_frame.pack(side=LEFT, expand=True, fill=BOTH)
    usb_button = Button(usbtest_frame, text='Start USB Test', command=test_usb.start)  # threaded, non blocking, no freeze
    usb_button.pack(expand=True, fill=BOTH)
    usb_success_button = Button(usbtest_frame, text='PASS', command=lambda: iotest_callback('USB', True), width=10, style='green_text.TButton')  # threaded, non blocking, no freeze
    usb_success_button.pack(side=LEFT, expand=True, fill=BOTH)
    usb_failed_button = Button(usbtest_frame, text='FAIL', command=lambda: iotest_callback('USB', False), width=10, style='red_text.TButton')  # threaded, non blocking, no freeze
    usb_failed_button.pack(side=RIGHT, expand=True, fill=BOTH)
    audiotest_frame = LabelFrame(iotest_frame, text='Audio Test', borderwidth=3, padding=(20, 20))
    audiotest_frame.pack(side=RIGHT, expand=True, fill=BOTH)
    audio_button = Button(audiotest_frame, text='Play audio', command=audiojack_test)  # threaded, non blocking, no freeze
    audio_button.pack(expand=True, fill=BOTH)
    audio_success_button = Button(audiotest_frame, text='PASS', command=lambda: iotest_callback('Audio', True), width=10, style='green_text.TButton')  # threaded, non blocking, no freeze
    audio_success_button.pack(side=LEFT, expand=True, fill=BOTH)
    audio_failed_button = Button(audiotest_frame, text='FAIL', command=lambda: iotest_callback('Audio', False), width=10, style='red_text.TButton')  # threaded, non blocking, no freeze
    audio_failed_button.pack(side=RIGHT, expand=True, fill=BOTH)
    process_item_failed(check_devices=False)
    finish_button = Button(frame, text='Finish and Restart', command=mode_0_clean_up, width=40, state=DISABLED, style='red_text.TButton')
    finish_button.pack(expand=True, fill=BOTH, side=BOTTOM, pady=(15, 0))
    Button(frame, text='refresh failed item list', command=process_item_failed).pack(expand=True, fill=BOTH, side=BOTTOM, pady=(10, 0))  # below finish button as using pack bottom side
    if update_skipped: Button(frame, text='Retry Windows Update', command=windows_update).pack(expand=True, fill=BOTH, side=BOTTOM, pady=(10, 0))
    while _3dmark_install_thread.is_alive():
        status('Waiting for 3DMark to install', log=False)
        root.update()
    status('Ready')
    finish_button['state'] = NORMAL  # only allow clean up if 3DMark finished installing
    bot_send_msg(f'Order {order_no} testing by {qc_person}\nDriver installation finished, requires human review now.\nTester notes: {note}')  # send message at last after GUI fully loaded


class driversC:
    def __init__(self):
        status('Initializing drivers (not installing now)')
        self.skip_intel_chipset = False
        self.skip_amd_chipset = False
        self.skip_lan = False
        self.skip_realtek_lan = False
        self.skip_intel_lan = False
        self.skip_marvell_lan = False
        self.skip_wifi = False
        self.skip_bt = False
        self.skip_Ngpu = False
        self.skip_Agpu = False
        self.skip_Igpu = False
        self.skip_audio = False
        self.driver_dict = {}
        self.driver_install_switches = {'AMD_Chipset': '/S', 'intel_chipset': ['-s -norestart -downgrade', '-s -overwrite', '-s -overwrite'],  # intel list: chipset, ME, Serial IO
                                        'intel_wifi': '-q -s', 'intel_bt': '/qb',
                                        'intel_lan': '/s /nr', 'realtek_lan': '-s', 'marvell_lan': '/qb',
                                        'msi_audio': '/s', 'asus_audio': '/s',
                                        'nvidia': '/n /passive /noeula /nofinish /nosplash',
                                        'radeon': '-install',
                                        'intel_gpu': '-p'}

        logger.debug(f'Checking for compatibility and driver installation files in {driversFolder}/')
        if not os.path.isdir(driversFolder):
            logger.error(f'Driver folder is not found at {driversFolder}/. All driver installation will be skipped.')
            messagebox.showerror(title='Driver folder not found!', message=f'Driver folder is not found at {driversFolder}/. All driver installation will be skipped.')  # driversC class is not initialized in mode 1 so no need check mode here
            self.skip_intel_chipset = True  # only 1 var is needed as you can't have multiple gen of intel chipset on same system
            self.skip_amd_chipset = True
            self.skip_lan = True
            self.skip_realtek_lan = True
            self.skip_intel_lan = True
            self.skip_marvell_lan = True
            self.skip_wifi = True
            self.skip_bt = True
            self.skip_Ngpu = True
            self.skip_Agpu = True
            self.skip_Igpu = True
            self.skip_audio = True
        else:  # NOTE: all path uses \, thus need to use \ for ALL subprocess calls
            for driver in [path for path in os.listdir(driversFolder) if not (path.endswith('.txt') or path.endswith('.log'))]:
                self.driver_dict[' '.join(os.path.splitext(driver)[0].split()[:-1])] = {'path': os.path.join(driversFolder, driver), 'version': os.path.splitext(driver)[0].split()[-1] if os.path.isfile(os.path.join(driversFolder, driver)) else driver.split()[-1]}
            if 'msi_audio' not in self.driver_dict.keys():
                if mobo.brand == 'MSI':
                    messagebox.showwarning(title='MSI Realtek Audio Driver not found', message=f'Skipping MSI Realtek Audio Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "msi_audio [version]".')
                    logger.warning(f'Skipping MSI Realtek Audio Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "msi_audio [version]".')
                    self.skip_audio = True
                elif mobo.brand != 'Asus':  # if it is asus, then can install Asus audio driver, so MSI one is not needed
                    logger.info('Using MSI Audio driver as the generic audio driver for other mobo brands for now. Note that this may fail due to compatibility issues.')
                    messagebox.showwarning(title='Realtek Audio Driver not found', message=f'Skipping Realtek Audio Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "msi_audio [version]".')
                    logger.warning(f'Skipping Realtek Audio Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "msi_audio [version]".')
                    self.skip_audio = True
            if 'asus_audio' not in self.driver_dict.keys() and mobo.brand == 'Asus':
                messagebox.showwarning(title='Asus Realtek Audio Driver not found', message=f'Skipping Asus Realtek Audio Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "asus_audio [version]".')
                logger.warning(f'Skipping Asus Realtek Audio Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "asus_audio [version]".')
                self.skip_audio = True

            if 'A' in cpu.brandCode and 'AMD_Chipset' not in self.driver_dict.keys():
                messagebox.showwarning(title='AMD Chipset Driver not found', message=f'Skipping AMD Chipset Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "AMD_Chipset [version].exe".')
                logger.warning(f'Skipping AMD Chipset Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "AMD_Chipset [version].exe".')
                self.skip_amd_chipset = True
            elif 'I' in cpu.brandCode:
                if not mobo.intel_gen:
                    logger.warning(f'Skipping Intel Chipset Driver installation as chipset detected is not supported. Please install manually.')
                    self.skip_intel_chipset = True
                elif mobo.intel_gen == 11 and 'intel_chipset_11' not in self.driver_dict.keys():
                    messagebox.showwarning(title='Intel 11th Gen Chipset Driver not found', message=f'Skipping Intel 11th Gen Chipset Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "intel_chipset_11 0".')
                    logger.warning(f'Skipping Intel 11th Gen Chipset Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "intel_chipset_11 0".')
                    self.skip_intel_chipset = True
                elif mobo.intel_gen == 12 and 'intel_chipset_12' not in self.driver_dict.keys():
                    messagebox.showwarning(title='Intel 12th Gen Chipset Driver not found', message=f'Skipping Intel 12th Gen Chipset Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "intel_chipset_12 0".')
                    logger.warning(f'Skipping Intel 12th Gen Chipset Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "intel_chipset_12 0".')
                    self.skip_intel_chipset = True

            if network.physical_LAN_adaptors:
                if not all([venid in network.vendor_id_dict.keys() for venid in network.LAN_vendorids]):
                    messagebox.showwarning(title='Unsupported LAN adaptor', message=f'One or more LAN adaptor detected is not either Intel, Realtek or Marvell one, thus drivers for those adaptors will not be installed.')
                    logger.warning(f'One or more LAN adaptor detected is not either Intel, Realtek or Marvell one, thus drivers for those adaptors will not be installed.')
                if '8086' in network.LAN_vendorids and 'intel_lan' not in self.driver_dict.keys():
                    messagebox.showwarning(title='Intel LAN Driver not found', message=f'Skipping Intel LAN Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "intel_lan [version]".')
                    logger.warning(f'Skipping Intel LAN Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "intel_lan [version]".')
                    self.skip_intel_lan = True
                if '10EC' in network.LAN_vendorids and 'realtek_lan' not in self.driver_dict.keys():
                    messagebox.showwarning(title='Realtek LAN Driver not found', message=f'Skipping Realtek LAN Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "realtek_lan [version]".')
                    logger.warning(f'Skipping Realtek LAN Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "realtek_lan [version]".')
                    self.skip_realtek_lan = True
                if '1D6A' in network.LAN_vendorids and 'marvell_lan' not in self.driver_dict.keys():  # Marvell Aquantia, mostly found on 10G boards
                    messagebox.showwarning(title='Marvell LAN Driver not found', message=f'Skipping Marvell LAN Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "marvell_lan [version]".')
                    logger.warning(f'Skipping Marvell LAN Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "marvell_lan [version]".')
                    self.skip_marvell_lan = True
            else:
                logger.warning(f'Skipping LAN Driver installation as no LAN adaptor detected.')
                self.skip_lan = True

            if network.physical_wifi_adaptors:  # no warning for not found wifi as not every PC has to have WiFi
                if '8086' in network.wifi_vendorids or '14C3' in network.wifi_vendorids:  # only support Intel and MediaTek for now
                    if '8086' in network.wifi_vendorids:  # supports installing driver for multiple WiFi adaptors but only supports testing 1
                        if 'intel_wifi' not in self.driver_dict.keys():
                            messagebox.showwarning(title='Intel WiFi Driver not found', message=f'Skipping Intel WiFi Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "intel_wifi [version].exe".')
                            logger.warning(f'Skipping Intel WiFi Driver installation as it is not found at {driversFolder}/, check that it has been renamed to "intel_wifi [version].exe".')
                            self.skip_wifi = True
                        if 'intel_bt' not in self.driver_dict.keys():
                            messagebox.showwarning(title='Intel Bluetooth Driver not found', message=f'Skipping Intel Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "intel_bt [version]".')
                            logger.warning(f'Skipping Intel Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "intel_bt [version]".')
                            self.skip_bt = True
                    if '14C3' in network.wifi_vendorids:  # mediatek
                        if mobo.brand == 'Asus':
                            if 'asus_mediatek_wifi' not in self.driver_dict.keys():
                                messagebox.showwarning(title='Asus MediaTek WiFi Driver not found', message=f'Skipping Asus MediaTek WiFi Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "asus_mediatek_wifi [version]".')
                                logger.warning(f'Skipping Asus MediaTek WiFi Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "asus_mediatek_wifi [version]".')
                                self.skip_wifi = True
                            if 'asus_mediatek_bt' not in self.driver_dict.keys():
                                messagebox.showwarning(title='Asus MediaTek Bluetooth Driver not found', message=f'Skipping Asus MediaTek Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "asus_mediatek_bt [version]".')
                                logger.warning(f'Skipping Asus MediaTek Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "asus_mediatek_bt [version]".')
                                self.skip_bt = True
                        elif mobo.brand == 'Gigabyte':
                            if 'gigabyte_mediatek_wifi' not in self.driver_dict.keys():
                                messagebox.showwarning(title='Gigabyte MediaTek WiFi Driver not found', message=f'Skipping Asus MediaTek WiFi Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_wifi [version]".')
                                logger.warning(f'Skipping Gigabyte MediaTek WiFi Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_wifi [version]".')
                                self.skip_wifi = True
                            if 'gigabyte_mediatek_bt' not in self.driver_dict.keys():
                                messagebox.showwarning(title='Gigabyte MediaTek Bluetooth Driver not found', message=f'Skipping Asus MediaTek Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_bt [version]".')
                                logger.warning(f'Skipping Gigabyte MediaTek Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_bt [version]".')
                                self.skip_bt = True
                        else:
                            logger.info('Using Gigabyte MediaTek drivers as the generic WiFi and BT driver for other mobo brands for now. Note that this may fail due to compatibility issues.')
                            if 'gigabyte_mediatek_wifi' not in self.driver_dict.keys():
                                messagebox.showwarning(title='Gigabyte MediaTek WiFi Driver not found', message=f'Skipping Asus MediaTek WiFi Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_wifi [version]".')
                                logger.warning(f'Skipping Gigabyte MediaTek WiFi Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_wifi [version]".')
                                self.skip_wifi = True
                            if 'gigabyte_mediatek_bt' not in self.driver_dict.keys():
                                messagebox.showwarning(title='Gigabyte MediaTek Bluetooth Driver not found', message=f'Skipping Asus MediaTek Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_bt [version]".')
                                logger.warning(f'Skipping Gigabyte MediaTek Bluetooth Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "gigabyte_mediatek_bt [version]".')
                                self.skip_bt = True
                else:
                    messagebox.showwarning(title='Unsupported WiFi adaptor', message=f'Skipping WiFi and Bluetooth Driver installation as the WiFi adaptor detected is not an Intel or MediaTek one.')
                    logger.warning(f'Skipping WiFi and Bluetooth Driver installation as the WiFi adaptor detected is not an Intel or MediaTek one.')
                    self.skip_wifi = True
                    self.skip_bt = True
            if 'N' in gpu.brandCode:
                if gpu.quadro and 'nvidia_quadro' not in self.driver_dict.keys():
                    messagebox.showwarning(title='Nvidia GPU Driver not found', message=f'Skipping Nvidia Quadro GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "nvidia_quadro [version]".')
                    logger.warning(f'Skipping Nvidia Quadro GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "nvidia_quadro [version]".')
                    self.skip_Ngpu = True
                elif 'nvidia' not in self.driver_dict.keys():
                    if order_no: bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: Nvidia GeForce GPU Driver not found.\nTester notes: {note}')
                    messagebox.showwarning(title='Nvidia GPU Driver not found', message=f'Skipping Nvidia GeForce GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "nvidia [version]".')
                    logger.warning(f'Skipping Nvidia GeForce GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "nvidia [version]".')
                    self.skip_Ngpu = True
            if 'A' in gpu.brandCode and 'radeon' not in self.driver_dict.keys():
                messagebox.showwarning(title='AMD GPU Driver not found', message=f'Skipping AMD GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "radeon [version]".')
                logger.warning(f'Skipping AMD GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "radeon [version]".')
                self.skip_Agpu = True
            if 'I' in gpu.brandCode and 'intel_gpu' not in self.driver_dict.keys():
                messagebox.showwarning(title='Intel GPU Driver not found', message=f'Skipping Intel GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "intel_gpu [version]".')
                logger.warning(f'Skipping Intel GPU Driver installation as it is not found at {driversFolder}/, check that it has been extracted and renamed to "intel_gpu [version]".')
                self.skip_Igpu = True

    def install_mobo_driver(self):
        if 'A' in cpu.brandCode:
            if self.skip_amd_chipset:
                logger.warning('AMD Chipset Driver installation skipped. Detailed reason is logged above during initialization phase.')
                item_failed.add('AMD Chipset Driver installation skipped')
            else:
                status('Installing AMD Chipset Driver')
                part = 'AMD_Chipset'
                logger.info(f'Installing AMD Chipset Driver version {self.driver_dict[part]["version"]}')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}" {self.driver_install_switches[part]}')
                r = subprocess.call(fr'"{self.driver_dict[part]["path"]}" {self.driver_install_switches[part]}')  # removed return code check as AMD returns non 0 for successful installs but newer version exists
                logger.info(f'AMD Chipset Driver installed with return code {r}')
        elif 'I' in cpu.brandCode:
            if self.skip_intel_chipset:
                logger.warning('Intel Chipset Driver installation skipped. Detailed reason is logged above during initialization phase.')
                item_failed.add('Intel Chipset Driver installation skipped')
            else:
                status('Installing Intel Chipset Drivers')
                part = 'intel_chipset'
                logger.info(f'Installing Intel {mobo.intel_gen}th Gen Chipset Driver')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[f"{part}_{mobo.intel_gen}"]["path"]}\SetupChipset.exe" {self.driver_install_switches[part][0]}')
                r = subprocess.call(fr'"{self.driver_dict[f"{part}_{mobo.intel_gen}"]["path"]}\SetupChipset.exe" {self.driver_install_switches[part][0]}')
                if r and r != 3010:  # 3010 means success but need reboot
                    logger.warning(f'Intel Chipset Driver installation failed with return code {r}')
                    item_failed.add('Intel Chipset Driver installation failed')
                else:
                    logger.info(f'{Fore.GREEN}Intel Chipset Driver installed successfully.{Fore.RESET}')

                logger.info(f'Installing Intel {mobo.intel_gen}th Gen Management Engine Driver')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[f"{part}_{mobo.intel_gen}"]["path"]}\SetupME.exe" {self.driver_install_switches[part][1]}')
                r = subprocess.call(fr'"{self.driver_dict[f"{part}_{mobo.intel_gen}"]["path"]}\SetupME.exe" {self.driver_install_switches[part][1]}')
                if r and r != 3010:  # 3010 means success but need reboot:
                    logger.warning(f'Intel Management Engine Driver installation failed with return code {r}')
                    item_failed.add('Intel Management Engine Driver installation failed')
                else:
                    logger.info(f'{Fore.GREEN}Intel Management Engine Driver installed successfully.{Fore.RESET}')

                logger.info(f'Installing Intel {mobo.intel_gen}th Gen SerialIO Driver')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[f"{part}_{mobo.intel_gen}"]["path"]}\SetupSerialIO.exe" {self.driver_install_switches[part][2]}')
                r = subprocess.call(fr'"{self.driver_dict[f"{part}_{mobo.intel_gen}"]["path"]}\SetupSerialIO.exe" {self.driver_install_switches[part][2]}')
                if r and r != 3010:  # 3010 means success but need reboot
                    logger.warning(f'Intel SerialIO Driver installation failed with return code {r}')
                    item_failed.add('Intel SerialIO Driver installation failed')
                else:
                    logger.info(f'{Fore.GREEN}Intel SerialIO Driver installed successfully.{Fore.RESET}')
        status('Ready')

    def install_lan_driver(self):
        if self.skip_lan:
            logger.warning('LAN Driver installation skipped. Detailed reason is logged above during initialization phase.')
            item_failed.add('LAN Driver installation skipped')
        else:
            if '8086' in network.LAN_vendorids:
                if self.skip_intel_lan:
                    logger.warning('Intel LAN Driver installation skipped. Detailed reason is logged above during initialization phase.')
                    item_failed.add('Intel LAN Driver installation skipped')
                else:
                    status('Installing Intel LAN Driver')
                    part = 'intel_lan'
                    logger.info(f'Installing Intel LAN Driver version {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\APPS\SETUP\SETUPBD\Winx64\SetupBD.exe" {self.driver_install_switches[part]}')
                    r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\APPS\SETUP\SETUPBD\Winx64\SetupBD.exe" {self.driver_install_switches[part]}')
                    if r:
                        logger.error(f'Intel LAN Driver installer reported error. Please try again manually. Return code: {r}')
                        item_failed.add('Intel LAN Driver installation failed')
                    else:
                        logger.info(f'{Fore.GREEN}Intel LAN Driver installed successfully{Fore.RESET}')
            if '10EC' in network.LAN_vendorids:
                if self.skip_realtek_lan:
                    logger.warning('Realtek LAN Driver installation skipped. Detailed reason is logged above during initialization phase.')
                    item_failed.add('Realtek LAN Driver installation skipped')
                else:
                    status('Installing Realtek LAN Driver')
                    part = 'realtek_lan'
                    logger.info(f'Installing Realtek LAN Driver version {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\setup.exe" {self.driver_install_switches[part]}')
                    r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\setup.exe" {self.driver_install_switches[part]}')
                    if r:
                        logger.error(f'Realtek LAN Driver installer reported error. Please try again manually. Return code {r}')
                        item_failed.add('Realtek LAN Driver installation failed')
                    else:
                        logger.info(f'{Fore.GREEN}Realtek LAN Driver installed successfully{Fore.RESET}')
            if '1D6A' in network.LAN_vendorids:
                if self.skip_marvell_lan:
                    logger.warning('Marvell LAN Driver installation skipped. Detailed reason is logged above during initialization phase.')
                    item_failed.add('Marvell LAN Driver installation skipped')
                else:
                    status('Installing Marvell LAN Driver')
                    part = 'marvell_lan'
                    logger.info(f'Installing Marvell LAN Driver version {self.driver_dict[part]["version"]}')
                    eli_file = [file for file in os.listdir(f'drivers/{self.driver_dict[part]["path"]}') if file.endswith('.msi') and 'x64' in file]
                    if len(eli_file) > 1: logger.warning(f'Multiple 64-bit MSI installers found for Marvell LAN Driver. Using the first one: {eli_file[0]}')
                    if verbose_logging: logger.debug(fr'Executing start /wait msiexec /i "{self.driver_dict[part]["path"]}\{eli_file[0]}" {self.driver_install_switches[part]}')
                    r = subprocess.call(fr'start /wait msiexec /i "{self.driver_dict[part]["path"]}\{eli_file[0]}" {self.driver_install_switches[part]}', shell=True)
                    if r:
                        logger.error(f'Marvell LAN Driver installer reported error. Please try again manually. Return code {r}')
                        item_failed.add('Marvell LAN Driver installation failed')
                    else:
                        logger.info(f'{Fore.GREEN}Marvell LAN Driver installed successfully{Fore.RESET}')
            status('Ready')

    def install_wifi_bt_driver(self):
        if self.skip_wifi:
            logger.warning('WiFi Driver installation skipped. Detailed reason is logged above during initialization phase.')
            item_failed.add('WiFi Driver installation skipped')
        else:
            if '8086' in network.wifi_vendorids:
                status('Installing Intel WiFi Driver')
                part = 'intel_wifi'
                logger.info(f'Installing Intel WiFi Driver {self.driver_dict[part]["version"]}')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}" {self.driver_install_switches[part]}')
                r = subprocess.call(fr'"{self.driver_dict[part]["path"]}" {self.driver_install_switches[part]}')
                if r:
                    logger.error(f'Intel WiFi Driver installer reported error. Please try again manually. Return code: {r}')
                    item_failed.add('Intel WiFi Driver installation failed')
                    self.skip_wifi = True
                else:
                    logger.info(f'{Fore.GREEN}Intel WiFi Driver installed successfully.{Fore.RESET}')
            if '14C3' in network.wifi_vendorids:
                status('Installing MediaTek WiFi Driver')
                if mobo.brand == 'Asus':
                    part = 'asus_mediatek_wifi'
                    logger.info(f'Installing Asus MediaTek WiFi Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "pnputil -a "{self.driver_dict[part]["path"]}\mtkwl6ex.inf" /install"')
                    r = subprocess.call(fr'pnputil -a "{self.driver_dict[part]["path"]}\mtkwl6ex.inf" /install')
                    if r and r != 3010:  # 3010 means success but need reboot
                        if r == 259 and '14C3' not in network.wifi_devices_no_driver.keys():  # driver already exists OR no supported device found. Do another check if there is any no driver mediatek wifi device
                            # there is mediatek device but its not in no driver list, means that it already has a driver, so we can ignore r = 259 here
                            logger.error(f'Asus MediaTek WiFi Driver seems to be already installed. Return code: {r}')
                            item_failed.add('Asus MediaTek WiFi Driver is already installed - please verify')
                        else:
                            logger.error(f'Asus MediaTek WiFi Driver installer reported error. Please try again manually. Return code: {r}')
                            item_failed.add('Asus MediaTek WiFi Driver installation failed')
                            self.skip_wifi = True
                    else:
                        logger.info(f'{Fore.GREEN}Asus MediaTek WiFi Driver installed successfully.{Fore.RESET}')
                elif mobo.brand == 'Gigabyte':
                    part = 'gigabyte_mediatek_wifi'
                    logger.info(f'Installing Gigabyte MediaTek WiFi Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "pnputil -a "{self.driver_dict[part]["path"]}\mtkwl6ex.inf" /install"')
                    r = subprocess.call(fr'pnputil -a "{self.driver_dict[part]["path"]}\mtkwl6ex.inf" /install')
                    if r and r != 3010:
                        if r == 259 and '14C3' not in network.wifi_devices_no_driver.keys():
                            logger.error(f'Gigabyte MediaTek WiFi Driver seems to be already installed. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek WiFi Driver is already installed - please verify')
                        else:
                            logger.error(f'Gigabyte MediaTek WiFi Driver installer reported error. Please try again manually. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek WiFi Driver installation failed')
                            self.skip_wifi = True
                    else:
                        logger.info(f'{Fore.GREEN}Gigabyte MediaTek WiFi Driver installed successfully.{Fore.RESET}')
                else:
                    logger.info('Using Gigabyte MediaTek drivers as the generic WiFi driver for other mobo brands for now. Note that this may fail due to compatibility issues.')
                    part = 'gigabyte_mediatek_wifi'
                    logger.info(f'Installing Gigabyte MediaTek WiFi Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "pnputil -a "{self.driver_dict[part]["path"]}\mtkwl6ex.inf" /install"')
                    r = subprocess.call(fr'pnputil -a "{self.driver_dict[part]["path"]}\mtkwl6ex.inf" /install')
                    if r and r != 3010:
                        if r == 259 and '14C3' not in network.wifi_devices_no_driver.keys():
                            logger.error(f'Gigabyte MediaTek WiFi Driver seems to be already installed. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek WiFi Driver is already installed - please verify')
                        else:
                            logger.error(f'Gigabyte MediaTek WiFi Driver installer reported error. Please try again manually. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek WiFi Driver installation failed')
                            self.skip_wifi = True
                    else:
                        logger.info(f'{Fore.GREEN}Gigabyte MediaTek WiFi Driver installed successfully.{Fore.RESET}')

        if self.skip_bt:
            logger.warning('Bluetooth Driver installation skipped. Detailed reason is logged above during initialization phase.')
            item_failed.add('Bluetooth Driver installation skipped')
        else:
            if '8086' in network.wifi_vendorids:
                status('Installing Intel Bluetooth Driver')
                part = 'intel_bt'
                logger.info(f'Installing Intel Bluetooth Driver {self.driver_dict[part]["version"]}')
                if verbose_logging: logger.debug(fr'Executing start /wait msiexec /i "{self.driver_dict[part]["path"]}\Intel Bluetooth.msi" {self.driver_install_switches[part]}')
                r = subprocess.call(fr'start /wait msiexec /i "{self.driver_dict[part]["path"]}\Intel Bluetooth.msi" {self.driver_install_switches[part]}', shell=True)
                if r and r != 3010:  # 3010 means success but need reboot
                    logger.error(f'Intel Bluetooth Driver installer reported error. Please try again manually. Return code: {r}')
                    item_failed.add('Intel Bluetooth Driver installation failed')
                    self.skip_bt = True
                else:
                    logger.info(f'{Fore.GREEN}Intel Bluetooth Driver installed successfully.{Fore.RESET}')
            if '14C3' in network.wifi_vendorids:
                status('Installing MediaTek Bluetooth Driver')
                if mobo.brand == 'Asus':
                    part = 'asus_mediatek_bt'
                    logger.info(f'Installing Asus MediaTek Bluetooth Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "pnputil -a "{self.driver_dict[part]["path"]}\mtkbtfilter.inf" /install"')
                    r = subprocess.call(fr'pnputil -a "{self.driver_dict[part]["path"]}\mtkbtfilter.inf" /install')
                    if r and r != 3010:
                        if r == 259 and '14C3' not in network.wifi_devices_no_driver.keys():
                            logger.error(f'Asus MediaTek Bluetooth Driver seems to be already installed. Return code: {r}')
                            item_failed.add('Asus MediaTek Bluetooth Driver is already installed - please verify')
                        else:
                            logger.error(f'Asus MediaTek Bluetooth Driver installer reported error. Please try again manually. Return code: {r}')
                            item_failed.add('Asus MediaTek Bluetooth Driver installation failed')
                            self.skip_bt = True
                    else:
                        logger.info(f'{Fore.GREEN}Asus MediaTek Bluetooth Driver installed successfully.{Fore.RESET}')
                elif mobo.brand == 'Gigabyte':
                    part = 'gigabyte_mediatek_bt'
                    logger.info(f'Installing Gigabyte MediaTek Bluetooth Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "pnputil -a "{self.driver_dict[part]["path"]}\mtkbtfilter.inf" /install"')
                    r = subprocess.call(fr'pnputil -a "{self.driver_dict[part]["path"]}\mtkbtfilter.inf" /install')
                    if r and r != 3010:
                        if r == 259 and '14C3' not in network.wifi_devices_no_driver.keys():
                            logger.error(f'Gigabyte MediaTek Bluetooth Driver seems to be already installed. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek Bluetooth Driver is already installed - please verify')
                        else:
                            logger.error(f'Gigabyte MediaTek Bluetooth Driver installer reported error. Please try again manually. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek Bluetooth Driver installation failed')
                            self.skip_bt = True
                    else:
                        logger.info(f'{Fore.GREEN}Gigabyte MediaTek Bluetooth Driver installed successfully.{Fore.RESET}')
                else:
                    logger.info('Using Gigabyte MediaTek drivers as the generic Bluetooth driver for other mobo brands for now. Note that this may fail due to compatibility issues.')
                    part = 'gigabyte_mediatek_bt'
                    logger.info(f'Installing Gigabyte MediaTek Bluetooth Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "pnputil -a "{self.driver_dict[part]["path"]}\mtkbtfilter.inf" /install"')
                    r = subprocess.call(fr'pnputil -a "{self.driver_dict[part]["path"]}\mtkbtfilter.inf" /install')
                    if r and r != 3010:
                        if r == 259 and '14C3' not in network.wifi_devices_no_driver.keys():
                            logger.error(f'Gigabyte MediaTek Bluetooth Driver seems to be already installed. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek Bluetooth Driver is already installed - please verify')
                        else:
                            logger.error(f'Gigabyte MediaTek Bluetooth Driver installer reported error. Please try again manually. Return code: {r}')
                            item_failed.add('Gigabyte MediaTek Bluetooth Driver installation failed')
                            self.skip_bt = True
                    else:
                        logger.info(f'{Fore.GREEN}Gigabyte MediaTek Bluetooth Driver installed successfully.{Fore.RESET}')
        status('Ready')

    def install_gpu_driver(self):
        status('Installing GPU Driver')
        if 'N' in gpu.brandCode:
            if self.skip_Ngpu:
                logger.warning('Nvidia GPU Driver installation skipped. Detailed reason is logged above during initialization phase.')
                item_failed.add('Nvidia GPU Driver installation skipped')
            else:
                if gpu.quadro:
                    part = 'nvidia_quadro'
                    logger.info(f'Installing Nvidia Quadro GPU Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\setup.exe" {self.driver_install_switches["nvidia"]}')
                    r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\setup.exe" {self.driver_install_switches["nvidia"]}')  # here can reuse GeForce switches as they are the same
                    if r:
                        logger.error(f'Nvidia Quadro GPU Driver installer reported error. Please try again manually. Return code: {r}')
                        item_failed.add('Nvidia Quadro GPU Driver installation failed')
                    else:
                        logger.info(f'{Fore.GREEN}Nvidia Quadro GPU Driver installed successfully.{Fore.RESET}')
                else:
                    part = 'nvidia'
                    logger.info(f'Installing Nvidia GeForce GPU Driver {self.driver_dict[part]["version"]}')
                    if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\setup.exe" {self.driver_install_switches[part]}')
                    r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\setup.exe" {self.driver_install_switches[part]}')
                    if r:
                        logger.error(f'Nvidia GeForce GPU Driver installer reported error. Please try again manually. Return code: {r}')
                        item_failed.add('Nvidia GeForce GPU Driver installation failed')
                    else:
                        logger.info(f'{Fore.GREEN}Nvidia GeForce GPU Driver installed successfully.{Fore.RESET}')

        if 'A' in gpu.brandCode:
            if self.skip_Agpu:
                logger.warning('AMD GPU Driver installation skipped. Detailed reason is logged above during initialization phase.')
                item_failed.add('AMD GPU Driver installation skipped')
            else:
                part = 'radeon'
                logger.info(f'Installing AMD GPU Driver {self.driver_dict[part]["version"]}')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\Setup.exe" {self.driver_install_switches[part]}')
                r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\Setup.exe" {self.driver_install_switches[part]}')
                if r:
                    logger.error(f'AMD GPU Driver installer reported error. Please try again manually. Return code: {r}')
                    item_failed.add('AMD GPU Driver installation failed')
                else:
                    logger.info(f'{Fore.GREEN}AMD GPU Driver installed successfully.{Fore.RESET}')

        if 'I' in gpu.brandCode:
            if self.skip_Igpu:
                logger.warning('Intel GPU Driver installation skipped. Detailed reason is logged above during initialization phase.')
                item_failed.add('Intel GPU Driver installation skipped')
            else:
                part = 'intel_gpu'
                logger.info(f'Installing Intel GPU Driver {self.driver_dict[part]["version"]}')
                if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\Installer.exe" {self.driver_install_switches[part]}')
                r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\Installer.exe" {self.driver_install_switches[part]}')
                if r and r != 14 and r != 15:  # 14 and 15 both means installed successfully but need restart
                    logger.error(f'Intel GPU Driver installer reported error. Please try again manually. Return code: {r}')
                    item_failed.add('Intel GPU Driver installation failed')
                else:
                    logger.info(f'{Fore.GREEN}Intel GPU Driver installed successfully.{Fore.RESET}')
        status('Ready')

    def install_audio_driver(self):
        if self.skip_audio:
            logger.warning('Realtek Audio Driver installation skipped. Detailed reason is logged above during initialization phase.')
            item_failed.add('Realtek Audio Driver installation skipped')
        else:
            status('Installing Audio Driver')
            if mobo.brand == 'MSI':
                part = 'msi_audio'
                logger.info(f'Installing MSI Realtek Audio Driver {self.driver_dict[part]["version"]}')
            elif mobo.brand == 'Asus':
                part = 'asus_audio'
                logger.info(f'Installing Asus Realtek Audio Driver {self.driver_dict[part]["version"]}')
            else:
                logger.info('Using MSI Audio driver as the generic audio driver for other mobo brands for now. Note that this may fail due to compatibility issues.')
                part = 'msi_audio'
                logger.info(f'Installing MSI Realtek Audio Driver {self.driver_dict[part]["version"]}')
            if verbose_logging: logger.debug(fr'Executing "{self.driver_dict[part]["path"]}\Setup.exe" {self.driver_install_switches[part]}')
            r = subprocess.call(fr'"{self.driver_dict[part]["path"]}\Setup.exe" {self.driver_install_switches[part]}')
            if r:
                if r == 2147753984:  # this specific return code means driver is not compatible with the motherboard.
                    logger.warning(f'Realtek Audio Driver installer failed as driver is incompatible. Please try again using driver from official website. Return code: {r}')
                    item_failed.add('Realtek Audio Driver installation failed (incompatible)')
                else:
                    logger.error(f'Realtek Audio Driver installer reported error. Please try again manually. Return code: {r}')
                    item_failed.add('Realtek Audio Driver installation failed (unknown reason)')
            else:
                logger.info(f'{Fore.GREEN}Realtek Audio Driver installed successfully.{Fore.RESET}')
            status('Ready')

    def checkDevices(self):
        status('Checking all devices')
        global old_error_device, new_error_device
        problem_device = subprocess.check_output('pnputil /enum-devices /problem').decode('GB2312').split('\r\n')[2:-1]
        problem_device_grouped, temp = [], []
        for i in problem_device:  # Output from PnP Util is different in lines per device, so we need to split them into sublist by the space in middle
            if i:
                temp += [i]
            else:
                problem_device_grouped += [temp]
                temp = []  # reset sublist to empty
        new_error_device = [' '.join(i[1].split(':')[1:]).strip() for i in problem_device_grouped]  # captions only
        if old_error_device:
            fixed_item = [item for item in old_error_device if item not in new_error_device]
            for device in fixed_item:
                try:
                    item_failed.remove(f'{device}: Error in Device Manager')
                except KeyError:
                    pass  # if user manually tick the checkbox before error device scan, the item will already be removed from item failed so will have KeyError here
        if new_error_device:
            if not old_error_device:
                subprocess.Popen('devmgmt.msc', shell=True)  # only open device manager on first run but if a device is fixed then broke again, it will trigger this again
                if online:
                    import webbrowser
                    webbrowser.open(f'https://www.google.com/search?q={mobo.brand} {mobo.name} drivers')  # open google and search for the mobo for drivers
            logger.warning(f'The following devices reported a non-OK status in Device Manager. Please check on them manually. These devices will be included in a checklist at the end of driver phase.')
            for device in new_error_device:
                print(device)  # not using logger here as don't need timing and level details
                item_failed.add(f'{device}: Error in Device Manager')
        else:
            logger.info(f'{Fore.GREEN}No problematic device found.{Fore.RESET}')
        old_error_device = new_error_device.copy()
        status('Ready')


def windows_update():
    status('Running Windows Update')
    global online, update_skipped, win_update_thread
    if not online:
        check_online()
    if not online:
        logger.warning('You are not connected to Internet. Skipping windows update.')
        item_failed.add('Windows update - not connected to Internet')
        update_skipped = True
    else:
        def run():
            subprocess.call(['powershell.exe', 'Set-ExecutionPolicy -ExecutionPolicy Bypass -force | Out-Null'])
            p = subprocess.Popen(['powershell.exe', get_res_path('windows_update.ps1')], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            exec_live_output('Windows Update', p)

        update_skipped = False
        win_update_thread = threading.Thread(target=run)
        win_update_thread.start()


class usb_test:
    def __init__(self):
        status('Initializing USB test')
        self.stop = False
        self.devices = 0
        self.test_usb_thread = threading.Thread(target=self.mon_usb)

    def start(self):
        usb_button['text'] = 'Starting USB Test'
        usb_button['command'] = self.stop_test
        usb_button['state'] = DISABLED
        self.devices = 0
        status(f'Testing USB - Devices found: {self.devices}', log=False)
        try:
            self.test_usb_thread.start()
        except RuntimeError as e:
            if str(e) == 'threads can only be started once':
                self.test_usb_thread = threading.Thread(target=self.mon_usb)
                self.start()
            else:
                logger.error(f'Error when starting USB testing thread: {e}.')

    def stop_test(self):
        usb_button['text'] = 'Stopping USB test'
        usb_button['command'] = self.start
        usb_button['state'] = DISABLED
        self.stop = True
        status('Insert any USB device to exit usb testing', log=False)
        while self.test_usb_thread.is_alive():  # wait for it to end
            root.update()  # prevent freezing
        logger.info('USB insertion monitoring stopped')
        usb_button['state'] = NORMAL
        usb_button['text'] = 'Start USB Test'
        status('Ready')

    def mon_usb(self):
        status('Testing USB')
        self.stop = False
        import pythoncom
        pythoncom.CoInitialize()
        c = wmi.WMI()
        watcher = c.watch_for(raw_wql="SELECT * FROM __InstanceCreationEvent WITHIN 0.01 WHERE TargetInstance ISA \'Win32_USBHub\'", delay_secs=0.01)
        logger.info('USB insertion monitoring started')
        usb_button['state'] = NORMAL
        usb_button['text'] = 'Stop USB Test'
        while not self.stop:
            usb = watcher()
            self.devices += 1
            logger.info(f'{self.devices} USB insertions detected so far')
            logger.debug(usb)
            status(f'Testing USB - Devices found: {self.devices}', log=False)
        pythoncom.CoUninitialize()


def audiojack_test():
    status('Playing test audio')
    audio_button['text'] = 'Playing'
    audio_button['state'] = DISABLED
    import simpleaudio as sa
    logger.debug(f'Playing test audio from {get_res_path("test.wav")}')
    subprocess.call(f'{get_res_path("SoundVolumeView.exe")} /RunAsAdmin /Unmute DefaultRenderDevice /SetVolume DefaultRenderDevice 69')
    try:
        wave_obj = sa.WaveObject.from_wave_file(get_res_path("test.wav"))
        play_obj = wave_obj.play()
        while play_obj.is_playing():
            root.update()
    except Exception as e:
        logger.error(f'Error occurred when trying to play audio:\n{e}\nYou can try again')
    audio_button['state'] = NORMAL
    audio_button['text'] = 'Play audio'
    status('Ready')


def iotest_callback(part, r):
    if r:
        logger.info(f'{Fore.GREEN}{part} Test passed{Fore.RESET}')
        iotest_results[part] = True
        try:
            item_failed.remove(f'{part} Test')
        except KeyError:
            pass
    else:
        logger.warning(f'{part} Test failed.')
        item_failed.add(f'{part} Test')
        iotest_results[part] = False
    try:
        item_failed.remove(f'{part} Test pending')  # no matter result, can remove pending
    except KeyError:
        pass


def process_item_failed(check_devices=True):
    global item_failed_index, itemfailed_frame, item_failed_first_run, item_failed_reminder_label
    if check_devices: driver.checkDevices()  # to prevent duplicate run in init_qc
    if item_failed_first_run:
        item_failed_reminder_label = Label(frame, text='')
        item_failed_reminder_label.pack(pady=(10, 0))  # only pack once
    else:
        try:
            itemfailed_frame.destroy()
        except NameError:
            pass
        for k, v in iotest_results.items():  # only add the pending ones after one refresh
            if v is None: item_failed.add(f'{k} Test pending')
    if item_failed:
        item_failed_reminder_label['text'] = 'Please check the following failed items'
        item_failed_reminder_label['style'] = 'TLabel'  # reset to default style
        item_failed_index = {i: v for i, v in enumerate(item_failed.copy())}
        itemfailed_frame = LabelFrame(frame, text='Item Failed', borderwidth=3, padding=(15, 15))
        itemfailed_frame.pack(expand=True, fill=BOTH, side=TOP)
        if item_failed_first_run:
            bot_send_msg(f'Order {order_no} failed items:\n' + '\n'.join(item_failed), tag='failed_items')  # no need get return here as internally the sent msg obj is saved to another var
        elif failed_items_msg_sent:  # if failed items msg was sent successfully then can edit, else no need do anything because edit unsent message internally also no use, can just edit once sent out
            try:
                failed_items_telegram_msg.edit_text(f'Order {order_no} failed items:\n' + '\n'.join(item_failed))
            except telegram.error.BadRequest:  # if edited message is same as previous one, ignore this error
                pass
            except Exception as e:
                logger.error(f'Failed to edit message to "Order {order_no} everything OK!" due to error:\n{e}')
        else:
            bot_send_msg()
        for i, v in item_failed_index.items():
            v = v.replace("\'", "\\'").replace('\"', '\\"')  # prevent code mixing and syntax error in exec below
            exec(f'checkbox_var_{i} = IntVar(value=-1)', globals())
            exec(f"Checkbutton(itemfailed_frame, text='{v}', onvalue=1, offvalue=-1, var=checkbox_var_{i}, command=lambda: checkbox_itemfailed_update('{v}', checkbox_var_{i}.get())).pack(expand=True, fill=BOTH)", globals())
    else:
        if item_failed_first_run:
            item_failed_reminder_label['text'] = 'All tests passed except for manual IO tests above'
            item_failed_reminder_label['style'] = 'green_text.TLabel'
            bot_send_msg(f'Order {order_no} everything OK except for IO tests (to be done manually)', tag='failed_items')
        else:
            item_failed_reminder_label['text'] = 'All tests passed'
            item_failed_reminder_label['style'] = 'green_text.TLabel'
            bot_send_msg()
            if failed_items_msg_sent:
                try:
                    failed_items_telegram_msg.edit_text(f'Order {order_no} everything OK!')
                except telegram.error.BadRequest:  # if edited message is same as previous one, ignore this error
                    pass
                except Exception as e:
                    logger.error(f'Failed to edit message to "Order {order_no} everything OK!" due to error:\n{e}')
    item_failed_first_run = False


def checkbox_itemfailed_update(item, r):
    if r == 1:
        try:
            item_failed.remove(item)
        except KeyError:
            pass  # under rare cases the item is already removed somewhere else, so can ignore KeyError here
    else:
        item_failed.add(item)


def create_restart_shortcut():
    status('Setting auto-startup')
    with open(restart_shortcut_path, 'w', encoding='utf-8') as f:
        f.write(f'@echo off\nstart "" /d "{workingDir}" "{exec_path}"\nexit /b')
        f.flush()
    status('Ready')


# def wait_update_finish():
#     status('Waiting for Windows Update to finish')
#     while True:  # repeatedly check if the key exist as it only exists if Win Update is done installing
#         root.update()
#         try:
#             reg.OpenKey(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
#         except OSError:
#             time.sleep(5)
#         else:
#             break


def mode_0_clean_up():
    if all(list(iotest_results.values())) or messagebox.askyesno(title='Confirmation', message='Not all IO tests are marked as passed yet. Are you sure to continue?'):  # if all() is True, the pop up will not appear
        if messagebox.askyesno(title='Confirmation', message='The software will wait for windows update to finish then reboot the PC and launch into benchmark mode. Please make sure you have carried out all checks. Windows update will install during the reboot. Continue?'):
            status('Cleaning up...')
            process_item_failed()
            if item_failed:
                logger.warning('Item failed:\n' + '\n'.join(item_failed))
                with open(os.path.join(desktop, 'item failed.txt'), 'w+') as f:
                    f.write('\n'.join(item_failed))
                logger.info('Failed items written to item failed.txt on Desktop')
            else:
                logger.info(f'{Fore.GREEN}All tests passed.{Fore.RESET}')
            with open(os.path.join(os.environ['ProgramData'], 'mode.txt'), 'w') as f:
                f.write('1')
            create_restart_shortcut()
            logger.info('Mode switched. Next time this software is launched, it will enter benchmark mode.\n')
            temps.stop(kill=True)  # kill hwinfo in case anything
            if not update_skipped:
                status('Waiting for Windows Update')
                while win_update_thread.is_alive():
                    root.update()
                logger.info('Windows Update Finished. Cleaning up')
                subprocess.run(f"taskkill /im powershell.exe /f /t", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                subprocess.run(f"taskkill /im pwsh.exe /f /t", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                subprocess.call('powershell -NoProfile -NonInteractive -Command "Uninstall-Module -Name PSWindowsUpdate -force | Out-Null"', shell=True)
            else:
                logger.debug('Not doing Windows Update clean up as it is skipped.')
            logger.info('Program exit (If this console window does not close itself, you can manually close it now)')
            logFile.close()
            subprocess.call(f'{get_res_path("wu_reboot.exe")} /f /r')
            sys.exit(0)


def countdown(time, func, msg='Counting down'):
    time -= 1
    status(f'{msg} ({time}sec)', log=False)
    if time:
        root.after(1000, countdown, time, func, msg)  # execute countdown() every 1 second (recursive function)
    else:
        func()


def benchmark_init():
    status('Initializing benchmarks')
    global frame, burnin_avail, burnin_path, cinebench_avail, cb_path, bench_progressbar
    burnin_avail = True
    cinebench_avail = True
    burnin_path = 'BurnInTest'
    cb_path = 'CinebenchR20/Cinebench.exe'

    if not os.path.isfile(f'{burnin_path}/bit.exe'):
        burnin_avail = False
        logger.error(f'BurnIn Test executable bit.exe not found at {burnin_path}. Check if it exists besides the main program.')
        bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: BurnIn Test not found!\nTester notes: {note}')
        messagebox.showerror(title='BurnIn Test not found!', message=f'BurnIn Test executable bit.exe not found at {burnin_path}. Check if it exists besides the main program. If continue, BurnIn Test will be skipped.')

    if not os.path.isfile(f'{cb_path}'):
        cinebench_avail = False
        logger.error(f'Cinebench executable Cinebench.exe not found at {cb_path}. Check if it exists besides the main program.')
        bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: Cinebench not found!\nTester notes: {note}')
        messagebox.showerror(title='Cinebench not found!', message=f'Cinebench executable Cinebench.exe not found at {cb_path}. Check if it exists besides the main program. If continue, Cinebench will be skipped.')

    bench_progressbar = Progressbar(frame, orient=HORIZONTAL, mode='determinate')
    bench_progressbar.pack(expand=True, fill=BOTH, side=TOP)
    status('Paused for system to load (120sec)')
    temps.start()  # TODO: always takes a few tries to succeed. Very LOW priority to fix. Tested not start() function issue, should be keyboard presses not registering right after HWINFO launch.
    countdown(120, benchmark, msg='Paused for system to load')


def benchmark():
    global temp_plot_button, power_plot_button, cinebench_score
    idle_cpu_temp = temps.minTemp('CPU Temp')
    if gpu.dgpu:
        idle_gpu_temps = [temps.minTemp(i) for i in temps.data.columns[2:]]
        idle_gpu_temps_logs = [f'GPU{i + 1}: {idle_gpu_temps[i]}' for i in range(len(idle_gpu_temps))]
        logger.info(f'Recorded idle temps: CPU: {idle_cpu_temp}, {" ".join(idle_gpu_temps_logs)}')
    else:
        logger.info(f'Recorded idle temps: CPU: {idle_cpu_temp}')

    if burnin_avail:
        burnintest(burnin_path)
    else:
        logger.warning('Skipping BurnIn Test. Please see log for reason.')
    progressbar_step(bench_progressbar, total_steps=3)

    if cinebench_avail:
        cinebench_score = cinebench(cb_path)
    else:
        logger.warning('Skipping Cinebench. Please see log for reason.')
        cinebench_score = 'Cinebench skipped'
    progressbar_step(bench_progressbar, total_steps=3)

    if _3dmark.installed and _3dmark.ready_to_run:
        if gpu.quadro:
            _3dmark.run(bench='firestrike', loop=2)
        elif gpu.dgpu:  # has one or more dGPU that is not Quadro
            _3dmark.run(bench='timespy', loop=2)
        else:  # only iGPU
            _3dmark.run(bench='nightraid', loop=2)
    else:
        logger.warning('Skipping 3DMark as 3DMark is not installed or ready to run. Please see log for details.')
    progressbar_step(bench_progressbar, total_steps=3)
    bench_progressbar.destroy()
    # sleep(10)  # TODO: remove this and below sample scores in production
    # cinebench_score = 6969
    # _3dmark.scores = [(40, 69, 4069), (11, 22, 333)]
    # _3dmark.bench = 'Time Spy'
    logger.info('Benchmarks completed.')
    status('Ready')
    temps.stop()

    ramtest(size=ram.capacity * 1000, thread=cpu.thread_no[0])  # using thread count of first CPU only
    logger.debug('Loading post-benchmark GUI')
    avg_cpu_temp = temps.avgTemp('CPU Temp')
    if gpu.dgpu:
        avg_gpu_temps = [temps.avgTemp(i) for i in temps.data.columns[2:]]
        max_gpu_temps = [temps.maxTemp(i) for i in temps.data.columns[2:]]
    max_cpu_temp = temps.maxTemp('CPU Temp')
    min_cpu_pwr = temps.minTemp('CPU Power')
    avg_cpu_pwr = temps.avgTemp('CPU Power')
    max_cpu_pwr = temps.maxTemp('CPU Power')
    if gpu.dgpu:
        logger.info(f'Idle temps: CPU: {idle_cpu_temp}, {" ".join(idle_gpu_temps_logs)}')
        avg_gpu_temps_logs = [f'GPU{i + 1}: {avg_gpu_temps[i]}' for i in range(len(avg_gpu_temps))]
        logger.info(f'Avg temps: CPU: {avg_cpu_temp}, {" ".join(avg_gpu_temps_logs)}')
        max_gpu_temps_logs = [f'GPU{i + 1}: {max_gpu_temps[i]}' for i in range(len(max_gpu_temps))]
        logger.info(f'Max temps: CPU: {max_cpu_temp}, {" ".join(max_gpu_temps_logs)}')
    else:
        logger.info(f'Idle temps: CPU: {idle_cpu_temp}')
        logger.info(f'Avg temps: CPU: {avg_cpu_temp}')
        logger.info(f'Max temps: CPU: {max_cpu_temp}')
    logger.info(f'\nMin CPU power: {min_cpu_pwr}W\nAvg CPU power: {avg_cpu_pwr}W\nMax CPU power: {max_cpu_pwr}W')

    bench_scores_frame = LabelFrame(frame, text='Benchmark Scores', borderwidth=3, padding=(10, 10))
    bench_scores_frame.pack(expand=True, fill=BOTH, side=TOP)
    Label(bench_scores_frame, text=f'Cinebench: {cinebench_score}\n').pack(expand=True, fill=BOTH, side=TOP)
    i = 1
    for score in _3dmark.scores:
        _ = LabelFrame(bench_scores_frame, text=f'3DMark Run {i} ({_3dmark.bench})', borderwidth=3, padding=(10, 10))
        _.pack(expand=True, fill=BOTH, side=LEFT)
        Label(_, text=f'CPU: {score[0]}').pack(expand=True, fill=BOTH, side=TOP)
        Label(_, text=f'GPU: {score[1]}').pack(expand=True, fill=BOTH, side=TOP)
        Label(_, text=f'Total: {score[2]}').pack(expand=True, fill=BOTH, side=TOP)
        i += 1
    temps_frame = LabelFrame(frame, text='Hardware Temperatures (C)', borderwidth=3, padding=(5, 5))
    temps_frame.pack(expand=True, fill=BOTH, side=TOP)
    idle_temp_frame = LabelFrame(temps_frame, text='Idle', borderwidth=3, padding=(10, 10))
    idle_temp_frame.pack(expand=True, fill=BOTH, side=LEFT)
    Label(idle_temp_frame, text=f'CPU: {idle_cpu_temp}').pack(expand=True, fill=BOTH, side=TOP)
    if gpu.dgpu:
        for t in idle_gpu_temps_logs:
            Label(idle_temp_frame, text=t).pack(expand=True, fill=BOTH, side=TOP)
    max_temp_frame = LabelFrame(temps_frame, text='Maximum', borderwidth=3, padding=(10, 10))
    max_temp_frame.pack(expand=True, fill=BOTH, side=LEFT)
    Label(max_temp_frame, text=f'CPU: {max_cpu_temp}').pack(expand=True, fill=BOTH, side=TOP)
    if gpu.dgpu:
        for t in max_gpu_temps_logs:
            Label(max_temp_frame, text=t).pack(expand=True, fill=BOTH, side=TOP)
    avg_temp_frame = LabelFrame(temps_frame, text='Average', borderwidth=3, padding=(10, 10))
    avg_temp_frame.pack(expand=True, fill=BOTH, side=LEFT)
    Label(avg_temp_frame, text=f'CPU: {avg_cpu_temp}').pack(expand=True, fill=BOTH, side=TOP)
    if gpu.dgpu:
        for t in avg_gpu_temps_logs:
            Label(avg_temp_frame, text=t).pack(expand=True, fill=BOTH, side=TOP)
    power_frame = LabelFrame(temps_frame, text='CPU Power (W)', borderwidth=3, padding=(10, 10))
    power_frame.pack(expand=True, fill=BOTH, side=LEFT)
    Label(power_frame, text=f'Min: {min_cpu_pwr}').pack(expand=True, fill=BOTH, side=TOP)
    Label(power_frame, text=f'Avg: {avg_cpu_pwr}').pack(expand=True, fill=BOTH, side=TOP)
    Label(power_frame, text=f'Max: {max_cpu_pwr}').pack(expand=True, fill=BOTH, side=TOP)
    if temps.data is None:
        logger.warning('Plotting buttons have been hidden as there is no data available.')
    else:
        graph_buttons_frame = LabelFrame(frame, text='Graphing', borderwidth=3, padding=(10, 10))
        graph_buttons_frame.pack(expand=True, fill=BOTH, side=TOP)
        temp_plot_button = Button(graph_buttons_frame, text='Temperatures', command=temps.plot_temps)
        temp_plot_button.pack(expand=True, fill=BOTH, side=LEFT)
        power_plot_button = Button(graph_buttons_frame, text='Power', command=temps.plot_cpu_pwr)
        power_plot_button.pack(expand=True, fill=BOTH, side=RIGHT)
    cleanup_button = Button(frame, text='Clean up', command=bench_cleanup, width=40, style='red_text.TButton')
    cleanup_button.pack(expand=True, fill=BOTH, side=TOP)


def burnintest(path):
    status('BurnIn Test')
    disk_to_burn = disk.get_burnin_disks()
    with open(f'{path}/burnin_script.txt', 'w+') as f:
        f.write('SETTEST DISK\nSETDUTYCYCLE DISK 100\n')
        if disk_to_burn: f.write('\n'.join(['SETDISK DISK ' + drive for drive in disk_to_burn]))  # disk to burn can be empty
        f.write('\nSETDISK ALL no\nRUN CONFIG')
    subprocess.Popen(fr'"{path}/bit.exe"  -S "{path}\burnin_script.txt" -C "{get_res_path("dreamcore.bitcfg")}" -R -P -V -W')
    for i in range(1, 900):
        status(f'BurnIn Test running ({900 - i}sec left)\n(this window may appear frozen)', log=False)
        sleep(0.92)  # adjusted for lag in burnin test， equal to 855sec, 45sec offset
        root.update()
    logger.info('BurnIn Test finished. Test window is kept open.')
    status('Ready')


def cinebench(path):
    status('Cinebench')
    cinebench_score = None

    def run():
        nonlocal cinebench_score
        try:
            result = subprocess.check_output(f'{path} g_CinebenchCpuXTest=true g_acceptDisclaimer=true').decode('GB2312')
        except subprocess.CalledProcessError as e:
            logger.error(f'An error occurred during Cinebench. Score will be recorded as zero. Please try again manually. Detailed error: {e}')
            cinebench_score = 0
        else:
            with open(os.path.join(desktop, 'cinebench_log.txt'), 'a+') as f:
                f.write(result)
            lines = result.split()
            cinebench_score = round(float(lines[-2]))

    cinebench_thread = threading.Thread(target=run)
    cinebench_thread.start()
    while cinebench_thread.is_alive():
        root.update()
    logger.info(f'Cinebench completed. Score: {cinebench_score}. Detailed log saved to cinebench_log.txt on Desktop')
    status('Ready')
    return cinebench_score  # can be None


class _3dmarkC:
    def __init__(self):
        status('Initializing 3DMark (not installing now)')
        self.path = r'C:\Program Files\UL'
        self.installer_path = r'3DMark\3dmark-setup.exe'
        self.install_files = ['3DMark/3dmark-setup.exe', '3DMark/redist', '3DMark/data_x64.msi', '3DMark/data_x86.msi']
        self.install_files_ready = [os.path.exists(file) for file in self.install_files]
        self.dlcpath = r'C:\3DMarkDlc'
        self.dlczip = r'3DMark\3DMarkDlc.zip'
        self._7z_e_path = '\\'.join(self.dlcpath.split('\\')[:-1]) + '/'
        self.exepath = fr'{self.path}\3DMark\3DMarkCmd.exe'
        self.installed = os.path.isfile(self.exepath)
        self.ready_to_run = self.installed  # set it same as whether it is installed for now, but later on during installation it may be changed.
        self.bench = None

        if not all(self.install_files_ready):
            self.installer_exist = False
            for i in range(len(self.install_files_ready)):
                if not self.install_files_ready[i]:
                    logger.error(f'3DMark installation file {self.install_files[i]} not found. Check if it exists in the directory or named correctly.')
            if order_no: bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: 3DMark installation failed.\nTester notes: {note}')
            messagebox.showerror(title='3DMark installation files not found!', message=f'3DMark installation files not found at 3DMark/. Check if it exists in the directory or named correctly. If continue, 3DMark will be SKIPPED. More details can be found in console log.')
        else:
            self.installer_exist = True
        self.scores = []
        if not self.installed and mode == 1:
            logger.warning('3DMark is not installed even though it should be during previous run of QC Software. It will be installed now. Ignore this warning if you manually entered benchmark mode via mode.txt, else, please report to developer.')
            self.install()
        status('Ready')

    def install(self):
        if self.installer_exist:
            status('Installing 3DMark')
            logger.info(f'Installing 3DMark to {self.path}')
            subprocess.call(fr'"{self.installer_path}" /install /quiet /silent /installpath="{self.path}"')
            self.installed = os.path.isfile(self.exepath)
            self.ready_to_run = self.installed  # refresh it here but still may be overridden later
            self.sent_telegram = False  # need use class attribute here else this var will not be preserved
            if self.installed:
                logger.info(f'3DMark installed. CMD exe is located at {self.exepath}')
            else:
                logger.error(f'3DMark installation failed. CMD exe is not found at {self.exepath} after running installer.')
                if not self.sent_telegram and not mode:
                    bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: 3DMark installation failed.\nTester notes: {note}')
                    self.sent_telegram = True
                messagebox.showerror(title='3DMark installation failed!', message=f'3DMark installation failed. CMD exe is not found at {self.exepath} after running installer.')

            if not os.path.isfile(self.dlczip):
                self.ready_to_run = False
                logger.error(f'3DMark DLC zip file not found at {self.dlczip}. Check if it exists besides the main program.')
                if not self.sent_telegram and not mode:
                    bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: 3DMark installation failed.\nTester notes: {note}')
                    self.sent_telegram = True
                messagebox.showerror(title='3DMark DLC files not found!', message=f'3DMark DLC zip file not found at {self.dlczip}. Check if it exists besides the main program. If continue, 3DMark may not run properly.')
            else:
                logger.info(fr'Extracting 3DMark DLC from {self.dlczip} to {self.dlcpath}')
                r = subprocess.run(fr'{get_res_path("7za.exe")} x "{self.dlczip}" -aoa -o"{self._7z_e_path}" -y', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode
                if not r:
                    logger.info(f'3DMark DLC extraction success. Setting 3DMark DLC path to {self.dlcpath}')
                    subprocess.run(fr'"{self.exepath}" --path="{self.dlcpath}"', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                else:
                    logger.error(f'3DMark DLC extraction failed due to code {r}: {_7z_return_code.get(r, f"Unknown Error: {r}")}.')
                    if not self.sent_telegram and not mode:
                        bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: 3DMark installation failed.\nTester notes: {note}')
                        self.sent_telegram = True
                    messagebox.showerror(title='3DMark extraction failed', message=f'3DMark DLC extraction failed due to code {r}: {_7z_return_code.get(r, f"Unknown Error: {r}")}.')
                    self.ready_to_run = False

            logger.info(fr'Extracting 3DMark definition files into {self.path}\3DMark')
            r = subprocess.run(fr'{get_res_path("7za.exe")} e {get_res_path("3dmdef.zip")} -aoa -o"{self.path}\3DMark" -y', stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode
            if not r:
                logger.info(f'3DMark definition files extraction success.')
            else:
                logger.error(f'3DMark definition files extraction failed due to code {r}: {_7z_return_code.get(r, f"Unknown Error: {r}")}.')
                if not self.sent_telegram and not mode:
                    bot_send_msg(f'Order {order_no} testing by {qc_person}\nError: 3DMark installation failed.\nTester notes: {note}')
                    self.sent_telegram = True
                messagebox.showerror(title='3DMark extraction failed', message=f'3DMark definition files extraction failed due to code {r}: {_7z_return_code.get(r, f"Unknown Error: {r}")}.')
                self.ready_to_run = False

            self.activate()
            status('Ready')
        else:
            logger.error('3DMark installation files not found. See top for detailed error logs. 3DMark run will be skipped.')

        if self.installed:
            if self.ready_to_run:
                logger.info(f'{Fore.GREEN}3DMark installed successfully and ready to run.{Fore.RESET}')
            else:
                logger.warning('3DMark installed successfully but not ready to run. See above for detailed error logs.')
                if not mode: item_failed.add('3DMark installed but not ready to run (see log for details)')  # no point adding if not in driver mode
                if mode == 1 and messagebox.askyesno(title='3DMark installed but not ready to run', message='3DMark installed successfully but not ready to run (see logs for detailed reason). Do you still want to run it? It may crash or give unexpected results.'):  # only show in benchmark mode as in driver mode it is shown in the checklist
                    logger.debug('User confirmed to run 3DMark even though it is not ready to run.')
                    self.ready_to_run = True
        else:
            logger.error('3DMark installation failed (see log for details)')
            if not mode: item_failed.add('3DMark installation failed (see log for details)')
            self.installed = False

    def activate(self):
        status('Configuring 3DMark')
        with reg.CreateKeyEx(reg.HKEY_USERS, fr'{SID}\Software\UL\3DMark', 0, reg.KEY_ALL_ACCESS) as key:
            reg.SetValueEx(key, 'KeyCode', 0, reg.REG_SZ, 'REDACTED')
            reg.FlushKey(key)

    def run(self, bench, loop=1):
        root.attributes("-topmost", False)
        for run in range(1, loop + 1):  # result.xml will be overwritten for multiple runs, have to separately run
            status(f'3DMark (run no {run})')
            if bench == 'firestrike':  # Quadro
                logger.info(f'3DMark run no {run} (Fire Strike) starting')
                self.bench = 'Fire Strike'
                subprocess.call(fr'start "" /wait /max "{self.path}\3DMark\3DMarkCmd.exe" --definition="{self.path}\3DMark\firestrike.3dmdef" --loop=1 --export={desktop}\3dmresult{run}.xml --audio=off --systeminfo=off --systeminfomonitor=off --online=off --log={desktop}\3dmark.log', shell=True)
                logger.info(fr'3DMark run no {run} (Fire Strike) finished. Result saved to {desktop}\3dmresult{run}.xml')
            elif bench == 'timespy':  # all dGPU
                logger.info(f'3DMark run no {run} (Time Spy) starting')
                self.bench = 'Time Spy'
                subprocess.call(fr'start "" /wait /max "{self.path}\3DMark\3DMarkCmd.exe" --definition="{self.path}\3DMark\timespy.3dmdef" --loop=1 --export={desktop}\3dmresult{run}.xml --audio=off --systeminfo=off --systeminfomonitor=off --online=off --log={desktop}\3dmark.log', shell=True)
                logger.info(fr'3DMark run no {run} (Time Spy) finished. Result saved to {desktop}\3dmresult{run}.xml')
            else:  # iGPUs
                logger.info(f'3DMark run no {run} (Night Raid) starting')
                self.bench = 'Night Raid'
                subprocess.call(fr'start "" /wait /max "{self.path}\3DMark\3DMarkCmd.exe" --definition="{self.path}\3DMark\nightraid.3dmdef" --loop=1 --export={desktop}\3dmresult{run}.xml --audio=off --systeminfo=off --systeminfomonitor=off --online=off --log={desktop}\3dmark.log', shell=True)
                logger.info(fr'3DMark run no {run} (Night Raid) finished. Result saved to {desktop}\3dmresult{run}.xml')
        root.attributes("-topmost", True)
        logger.info(f'3DMark run completed. Ran {bench} for {loop} {"time" if loop == 1 else "times"}.')
        status('Ready')
        self.parse_score(firestrike=True if bench == 'firestrike' else False)

    def parse_score(self, firestrike=False):
        status('Parsing 3DMark results')
        import xml.etree.ElementTree as ET
        resultFiles = [file for file in os.listdir(desktop) if file.split('.')[0].startswith('3dmresult') and file.split('.')[1] == 'xml']
        logger.info(f'Found 3DMark result files: {resultFiles}')
        for file in resultFiles:
            logger.info(f'Parsing 3DMark result from {os.path.join(desktop, file)}')
            tree = ET.parse(os.path.join(desktop, file)).getroot()
            if firestrike:  # combined score is ignored but calculated into overall score. add this into documentation
                logger.debug('Entered Fire Strike score parsing mode')
                cpu = tree[0][0][5].text  # physics score
                gpu = tree[0][0][7].text  # graphics score
                total = tree[0][0][4].text  # includes cpu, gpu and another combined test
            else:
                logger.debug('Entered Time Spy/Night Raid score parsing mode')
                scores = {score.tag: score.text for score in tree[0][0] if 'Score' in score.tag}
                cpu = [value for key, value in scores.items() if 'CPUScore' in key][0]
                gpu = [value for key, value in scores.items() if 'GraphicsScore' in key][0]
                total = [value for key, value in scores.items() if '3DMarkScore' in key][0]
            self.scores += [(cpu, gpu, total)]
            logger.info(f'Parsed CPU score: {cpu}, GPU score: {gpu}, overall score: {total}')
        logger.info(f'Score parsing completed.')
        status('Ready')
        return self.scores

    def uninstall(self):
        status('Uninstalling 3DMark')
        logger.info('Uninstalling Futuremark SystemInfo')
        subprocess.call('start /wait /min wmic product where name="Futuremark SystemInfo" call uninstall /nointeractive', shell=True)
        logger.info('Futuremark SystemInfo uninstalled. Uninstalling 3DMark')
        subprocess.call(fr'3DMark\3dmark-setup.exe /uninstall /quiet /silent /installpath="{self.path}"')
        logger.info('3DMark uninstalled. Cleaning up registry')
        delete_sub_key(reg.HKEY_USERS, fr'{SID}\Software\UL\3DMark', task='3DMark')
        logger.info(f'Registry cleaned up. Removing DLC files from {self.dlcpath}')
        try:
            shutil.rmtree(self.dlcpath)
        except FileNotFoundError:  # sometimes 3DMark uninstaller removes the dlc folder so will have error here
            pass
        logger.info(f'DLC files removed from {self.dlcpath}')
        status('Ready')


def ramtest(size, thread):
    benchmark_scores_msg_str = ['Benchmark Scores:', f'Cinebench: {cinebench_score}\n', '3DMark:'] + [f'Run {i + 1} {_3dmark.bench} - CPU: {_3dmark.scores[i][0]}, GPU: {_3dmark.scores[i][1]}, Total: {_3dmark.scores[i][2]}' for i in range(len(_3dmark.scores))]
    _3dmark_total_scores = [float(score[2]) for score in _3dmark.scores]
    if max(_3dmark_total_scores) / min(_3dmark_total_scores) - 1 >= 0.05:  # if largest deviation of 3DMark score is larger than 5%
        logger.warning('3DMark scores differ by more than 5% between runs. Please verify thermal performance.')
        benchmark_scores_msg_str.append('WARNING: 3DMark scores differ by more than 5% between runs. Please verify thermal performance.')
    if os.path.isfile('RAM Test/ramtest.exe'):
        bot_send_msg(f'Order {order_no} testing by {qc_person}\nBenchmarks completed and RAM Test started.\nTester notes: {note}\n\n' + "\n".join(benchmark_scores_msg_str))
        status('Running RAM Test')
        logger.info(f'Launching RAM Test with {thread} threads and {size}MB size')
        import pywinauto
        app = pywinauto.application.Application(backend='uia').start('RAM Test/ramtest.exe', work_dir='RAM Test')
        while not if_running('ramtest.exe'):
            root.update()
        window = app.window(title='RAM Test')
        window.set_focus()
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('ctrl+a')
        for i in str(size):
            keyboard.press_and_release(i)
        keyboard.press_and_release('tab')
        keyboard.press_and_release('ctrl+a')
        for i in str(thread):
            keyboard.press_and_release(i)
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('tab')
        keyboard.press_and_release('enter')
        logger.info(f'RAM Test started')
    else:
        logger.error(f'RAM Test exe not found at RAM Test/ramtest.exe')
        bot_send_msg(f'Order {order_no} testing by {qc_person}\nBenchmarks completed but RAM Test is not found!\nTester notes: {note}\n\n' + "\n".join(benchmark_scores_msg_str))
        messagebox.showerror(title='Unable to launch RAM Test', message=f'RAM Test exe not found at RAM Test/ramtest.exe')


def bench_cleanup():
    if messagebox.askyesno(title='Confirm?', message=f'Continuing will uninstall 3DMark. Continue?'):
        status('Cleaning up')
        _3dmark_uninstall_thread = threading.Thread(target=_3dmark.uninstall)
        _3dmark_uninstall_thread.start()
        temps.stop(kill=True)

        logger.debug(fr'Removing Cinebench log files (scores can still be checked in the log file on desktop.)')
        try:
            os.remove(os.path.join(desktop, 'cinebench_log.txt'))
        except OSError:
            pass

        logger.debug(fr'Removing 3DMark log files (scores can still be checked in the log file on desktop.)')
        for run in range(1, 2 + 1):  # change the 2 here to new no of loop if needed
            try:
                os.remove(f'{desktop}/3dmresult{run}.xml')
            except OSError:
                pass
        try:
            os.remove(os.path.join(desktop, '3dmark.log'))
        except OSError:
            pass

        logger.debug('Removing HWINFO log files')
        delete_sub_key(reg.HKEY_CURRENT_USER, r'Software\HWiNFO64', task='HWINFO')
        for file in [file for file in os.listdir(temps.hwinfopath) if os.path.splitext(file)[1].lower() == '.csv' and file.startswith('HWiNFO_LOG_')]:  # not using temps.tempfiles as clean up may not remove manually generated csvs which interfere with determination of latest log file
            try:
                os.remove(os.path.join(temps.hwinfopath, file))
            except OSError:
                pass

        logger.debug('Removing RAM Test log files')
        try:
            os.remove('RAM Test/ramtest.log')
        except OSError:
            pass
        try:
            os.remove('RAM Test/settings.txt')
        except OSError:
            pass

        while _3dmark_uninstall_thread.is_alive():
            status('Waiting for 3DMark to uninstall', log=False)
            root.update()

        with open(os.path.join(os.environ['ProgramData'], 'mode.txt'), 'w') as f:  # set to OOB mode
            f.write('2')
        logger.info('Mode switched. Next time this software is launched, it will enter post-OOBE mode.\n')
        logger.info('Clean up done! Please OOBE now.')
        status('Ready')
        messagebox.showinfo(title='Clean up done!', message='Clean up done! Please OOBE now.')


def oob():
    if network.physical_wifi_adaptors:
        for i in range(len(network.profile_path)):
            network.load_profile(network.ssid[i], network.profile_path[i])
        for ssid in network.ssid:
            if network.connect(ssid): break  # only try to connect to WiFi no need connection tests as it can be assumed to be working
        else:
            logger.warning(f'WiFi: Failed to connect to all of these WiFi: {network.ssid}')
    windows_update()
    status('Loading OOB')
    global activate_button, view_activation_info_button, oem_submit_button
    try:
        with reg.OpenKeyEx(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion') as key:
            os_edition = reg.QueryValueEx(key, 'EditionId')[0].replace('Core', 'Home')  # Windows calls Home 'Core'
        logger.info(f'Detected Windows OS Edition: {os_edition}')
    except OSError as e:
        os_edition = 'Unknown'
        logger.error(f'Unable to detect Windows edition due to error: {e}')

    oem_model_frame = LabelFrame(frame, text='OEM Info', borderwidth=3, padding=(10, 10))
    oem_model_frame.pack(expand=True, fill=BOTH, side=TOP, pady=(5, 0))
    oem_model_input_frame = Frame(oem_model_frame, borderwidth=0, padding=(5, 0))
    oem_model_input_frame.pack(expand=True, fill=BOTH, side=TOP, pady=(0, 8))
    Label(oem_model_input_frame, text='Build Name: ').pack(side=LEFT)
    model_input = Entry(oem_model_input_frame, font=tk_default_font)
    model_input.pack(expand=True, fill=BOTH, side=RIGHT)

    def update_listbox(data):
        models_listbox.delete(0, END)
        for item in data:
            models_listbox.insert(END, item)

    def fill_model(e):
        model_input.delete(0, END)
        model_input.insert(0, models_listbox.get(ANCHOR))

    def search(e):
        typed = model_input.get()
        if not typed:
            data = models
        else:
            data = [item for item in models if typed.lower() in item.lower()]
        update_listbox(data)

    models = ['Alpha', 'Alpha Pro', 'Apollo', 'Dream Machine', 'Dream Machine Pro', 'Fuel', 'Ghost', 'Ghost Pro', 'Office', 'Phantom', 'Reverie']
    models_listbox = Listbox(oem_model_frame)
    models_listbox.pack(expand=True, fill=BOTH, side=TOP)
    models_listbox.config(width=0, height=0)  # reset the height and width to fit to content
    update_listbox(models)
    models_listbox.bind('<<ListboxSelect>>', fill_model)
    models_listbox.bind('<Return>', lambda event: set_oem_info(model_input.get().strip()))
    model_input.bind('<KeyRelease>', search)
    model_input.bind('<Return>', lambda event: set_oem_info(model_input.get().strip()))
    model_input.focus_set()
    oem_submit_button = Button(oem_model_frame, text='Submit', command=lambda: set_oem_info(model_input.get().strip()))
    oem_submit_button.pack(expand=True, fill=BOTH, side=TOP, pady=(10, 0))
    win_activation_frame = LabelFrame(frame, text=f'Windows Activation (OS Edition: {os_edition})', borderwidth=3, padding=(10, 10))
    win_activation_frame.pack(expand=True, fill=BOTH, side=TOP, pady=(5, 0))
    key_input_frame = Frame(win_activation_frame, borderwidth=0)
    key_input_frame.pack(expand=True, fill=BOTH, side=TOP)
    Label(key_input_frame, text='Activation Key: ').pack(side=LEFT)
    key_input = Entry(key_input_frame, font=tk_default_font)
    key_input.pack(expand=True, fill=BOTH, side=RIGHT)
    key_input.bind('<Return>', lambda event: activate(key=key_input.get().strip().upper()))
    activate_button_frame = Frame(win_activation_frame, borderwidth=0)
    activate_button_frame.pack(expand=True, fill=BOTH, side=TOP)
    activate_button = Button(activate_button_frame, text='Activate', command=lambda: activate(key=key_input.get().strip().upper()))
    activate_button.pack(expand=True, fill=BOTH, side=LEFT, pady=(12, 0))
    view_activation_info_button = Button(activate_button_frame, text='View Activation Info', command=lambda: activate(viewInfo=True))
    view_activation_info_button.pack(expand=True, fill=BOTH, side=RIGHT, pady=(12, 0))
    Button(frame, text='Finish and clean up', command=oob_cleanup, style='red_text.TButton').pack(expand=True, fill=BOTH, side=TOP, pady=(15, 0))
    status('Ready')
    # TODO: install rgb here, to research on MSI center and Asus mobo software and other common stuff


def set_oem_info(model):
    oem_submit_button['state'] = DISABLED
    if model:
        status('Setting OEM Info')
        logger.info('Installing OEM Logo')
        shutil.copy(get_res_path('oemlogo.bmp'), r'C:\Windows\System32\oemlogo.bmp')
        logger.info(f'Setting OEM Registry (model name will be "Dreamcore {model}")')
        with reg.CreateKeyEx(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation', 0, reg.KEY_ALL_ACCESS) as key:
            reg.SetValueEx(key, 'Manufacturer', 0, reg.REG_SZ, 'Dreamcore')
            reg.SetValueEx(key, 'Model', 0, reg.REG_SZ, 'Dreamcore ' + model)
            reg.SetValueEx(key, 'Logo', 0, reg.REG_SZ, r'C:\Windows\System32\oemlogo.bmp')
            reg.SetValueEx(key, 'SupportURL', 0, reg.REG_SZ, 'https://dreamcore.com.sg')
            reg.FlushKey(key)
        model = model.replace(" ", "").replace(".", "")  # remove spaces and dots to fit PC name requirement below
        logger.info(f'Renaming PC to {model}')
        subprocess.call(['PowerShell.exe', f'Rename-Computer -NewName "{model}" -Force'])
    else:
        messagebox.showerror(title='Empty Build Name!', message='Build Name cannot be blank!')
    logger.info('Setting wallpaper')
    shutil.copy(get_res_path('wallpaper.jpg'), r'C:\Windows\System32\dreamcore_wallpaper.jpg')
    ctypes.windll.user32.SystemParametersInfoW(20, 0, r'C:\Windows\System32\dreamcore_wallpaper.jpg', 0)  # set wallpaper
    with reg.OpenKey(reg.HKEY_CURRENT_USER, r'Control Panel\Desktop', 0, reg.KEY_ALL_ACCESS) as key:
        reg.SetValueEx(key, 'WallpaperStyle', 0, reg.REG_SZ, '10')
        reg.SetValueEx(key, 'JPEGImportQuality', 0, reg.REG_DWORD, 0x00000064)  # sets wallpaper quality to 100%
        reg.SetValueEx(key, 'WallPaper', 0, reg.REG_SZ, r'C:\Windows\System32\dreamcore_wallpaper.jpg'.lower())  # windll function don't preserve after reboot, so need modify registry too
    subprocess.run('taskkill /im explorer.exe /f & start explorer.exe', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  # restart file explorer for change to take effect
    status('Ready')
    oem_submit_button['state'] = NORMAL


def activate(viewInfo=False, key=None):
    if viewInfo:
        root.attributes("-topmost", False)
        view_activation_info_button['state'] = DISABLED
        subprocess.call('slmgr /dli', shell=True)  # xpr -> dli -> dlv from simple to complicated
        view_activation_info_button['state'] = NORMAL
        root.attributes("-topmost", True)
    else:
        activate_button['text'] = 'Activating'
        activate_button['state'] = DISABLED
        if not key:
            messagebox.showerror(title='Empty Activation Key!', message='Windows Activation Key cannot be blank.')
        elif len(key.split('-')) != 5 or len(key) != 29 or not all(len(i) == 5 for i in key.split('-')):  # key format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
            messagebox.showerror(title='Invalid Activation Key!', message='Invalid activation key format. Correct format is XXXXX-XXXXX-XXXXX-XXXXX-XXXXX')
        elif check_online():
            status('Activating Windows')
            logger.info(f'Activating Windows with key {key}')
            subprocess.call(f'slmgr //b /ipk {key}', shell=True)
            subprocess.call(f'slmgr //b /ato', shell=True)
            # subprocess.call(f'slmgr //b /cpky', shell=True)  # removes key from registry, disabled for now as Dreamcore may still need it in future.
            activation_status = subprocess.check_output(r'cscript /Nologo C:\Windows\System32\slmgr.vbs /xpr').decode('GB2312').split('\r\n')[1].strip()
            logger.info(f'Windows activation status: {activation_status}, activation successful: {activation_status == "Windows is in Notification mode"}')
            if activation_status == 'Windows is in Notification mode':
                messagebox.showerror(title='Windows Activation Failed', message='Windows Activation Failed. Please try again manually in Settings.')
            else:
                messagebox.showinfo(title='Windows Activation Successful!', message='Windows Activation Successful!')
            status('Ready')
        else:
            logger.error('Windows activation failed as there is no Internet connection.')
            messagebox.showerror(title='No Internet connection!', message='Windows activation failed as there is no Internet connection. Please connect to the Internet and try again.')
        activate_button['state'] = NORMAL
        activate_button['text'] = 'Activate'


def oob_cleanup():
    if messagebox.askyesno(title='Confirm?', message=f'Continuing will delete all temporary files created by the software and wait for Windows Update to finish before rebooting to create restore point. Continue?'):
        status('Cleaning up')
        create_restart_shortcut()
        with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system', 0, reg.KEY_ALL_ACCESS) as key:  # disable UAC for auto reboot for restore point
            reg.SetValueEx(key, 'EnableLUA', 0, reg.REG_DWORD, 0x00000000)
            reg.FlushKey(key)
        delete_sub_key(reg.HKEY_CURRENT_USER, r'Software\HWiNFO64', task='HWINFO')  # clean up again in case tester manually open HWINFO
        try:
            os.remove(os.path.join(desktop, 'item failed.txt'))
        except OSError:
            pass
        try:
            os.remove(os.path.join(desktop, 'disks_to_burnin_test.txt'))
        except OSError:
            pass

        with open(os.path.join(os.environ['ProgramData'], 'mode.txt'), 'w+') as f:
            f.write('3')

        if not update_skipped:
            status('Waiting for Windows Update')
            while win_update_thread.is_alive():
                root.update()
            logger.info('Windows Update Finished. Cleaning up')
            subprocess.run(f"taskkill /im powershell.exe /f /t", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.run(f"taskkill /im pwsh.exe /f /t", stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.call('powershell -NoProfile -NonInteractive -Command "Uninstall-Module -Name PSWindowsUpdate -force | Out-Null"', shell=True)
        else:
            logger.debug('Not doing Windows Update clean up as it is skipped.')
        logger.info('Program exit (If this console window does not close itself, you can manually close it now)')
        logFile.close()
        subprocess.call(f'{get_res_path("wu_reboot.exe")} /f /r')
        sys.exit(0)


def restore_point_mode():
    global restore_point_label
    restore_point_label = Label(frame, text='Creating restore point')
    restore_point_label.pack()
    if not restorePoint(): Button(frame, text='Retry create restore point', command=restorePoint).pack(expand=True, fill=BOTH, side=BOTTOM, pady=(10, 0))


def restorePoint():
    status('Creating restore point')
    global restore_point_label
    restore_point_label['text'] = 'Creating restore point'
    restore_point_label['style'] = 'TLabel'
    root.update()
    success = False
    from datetime import date
    logger.info(f'Creating restore point with name "DC_{str(date.today())}"')
    with reg.CreateKeyEx(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore', 0, reg.KEY_ALL_ACCESS) as key:
        reg.SetValueEx(key, 'SystemRestorePointCreationFrequency', 0, reg.REG_DWORD, 0x00000000)
        reg.FlushKey(key)
    with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system', 0, reg.KEY_ALL_ACCESS) as key:  # enable UAC back
        reg.SetValueEx(key, 'EnableLUA', 0, reg.REG_DWORD, 0x00000001)
        reg.FlushKey(key)
    if subprocess.call(f'powershell.exe -ExecutionPolicy Bypass -Command "Checkpoint-Computer -Description "DC_{str(date.today())}" -RestorePointType "MODIFY_SETTINGS""'):
        logger.error('Failed to create restore point. Please try again manually.')
        restore_point_label['text'] = 'Restore point creation failed'
        restore_point_label['style'] = 'red_text.TLabel'
        bot_send_msg(f'Order {order_no} testing by {qc_person}\nRestore point creation failed!\nTester notes: {note}')
    else:
        success = True
        logger.info(f'{Fore.GREEN}Restore point created successfully{Fore.RESET}')
        restore_point_label['text'] = 'Restore point created successfully'
        restore_point_label['style'] = 'green_text.TLabel'
        bot_send_msg(f'Order {order_no} testing by {qc_person}\nRestore point created. Ready to ship.\nTester notes: {note}')
        logger.debug('Cleaning up temp files')
        try:
            os.remove(os.path.join(os.environ["ProgramData"], 'tester.txt'))
        except OSError:
            pass
        try:
            os.remove(restart_shortcut_path)
        except OSError:
            logger.warning(f'Failed to remove restart shortcut from {restart_shortcut_path}')
        try:
            os.remove(os.path.join(os.environ['ProgramData'], 'mode.txt'))
        except OSError:
            logger.warning(f'Failed to remove mode.txt from {os.path.join(os.environ["ProgramData"], "mode.txt")}')
        logger.debug('Removing log file (this should be the last log entry you see)')
        logFile.close()
        os.remove(os.path.join(desktop, 'log.log'))
    status('Ready', log=False)  # log file alr gone so should not log
    return success


if frozen:
    pyi_splash.close()
status('Ready')
if mode == 1:
    gatherInfo(inputs=False, specs=True)  # silent spec info gathering
    from matplotlib import pyplot as plt  # plotting is only needed in benchmark mode

    benchmark_init()
elif mode == 2:
    gatherInfo(inputs=False, specs=True)  # silent spec info gathering
    oob()
elif mode == 3:
    restore_point_mode()
else:
    logger.info('Entered Driver Mode')
    start_button = Button(frame, text='Start', command=gatherInfo)
    start_button.pack(expand=True, fill=BOTH, side=TOP)
    start_button.bind('<Return>', lambda event: gatherInfo())
    start_button.focus_set()
root.mainloop()  # blocking until GUI window is closed
try:
    temps.stop(kill=True, log=False)  # kill hwinfo in case anything but cannot log because it might be called after log file is closed
except NameError:
    pass
if mode and mode != 3:  # mode 0 clean up alr closed log file so printing this will throw error, mode 3 deletes log file so should not create again.
    logger.info('Program exit (If this console window does not close itself, you can manually close it now)')
    logFile.close()
