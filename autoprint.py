import os
import time
import re
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
from subprocess import call
import pygetwindow as gw
from tkinter import Tk, Label, StringVar, Frame, BOTH
from tkinter import Canvas
from tkinter import messagebox
import signal
import sys
import pystray
from PIL import Image, ImageDraw
import winshell
from win32com.client import Dispatch


# Automatically find the Downloads folder
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

def find_acrobat():
    possible_paths = [
        "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe",
        "C:\\Program Files (x86)\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe",
        "C:\\Program Files\\Adobe\\Acrobat Reader DC\\Reader\\Acrobat.exe",
        "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\Acrobat.exe"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

def create_image():
    # Generate an image with a green circle indicating active monitoring
    image = Image.new('RGB', (64, 64), (0, 128, 0))  # Green background
    dc = ImageDraw.Draw(image)
    dc.ellipse((16, 16, 48, 48), fill=(0, 255, 0))  # Lighter green circle

    return image


def quit_app(icon, item):
    icon.stop()
    watch.stop()
    root.quit()
    root.destroy()
    sys.exit(0)


def show_window(icon, item):
    icon.stop()
    root.after(0, root.deiconify)


def setup_tray():
    icon_image = create_image()
    menu = pystray.Menu(
        pystray.MenuItem("Show", show_window),
        pystray.MenuItem("Exit", quit_app)
    )
    icon = pystray.Icon("Auto Print", icon_image, "Auto Print", menu)
    icon.run()


def signal_handler(sig, frame):
    print("Exiting gracefully...")
    app.update_status("Exiting gracefully...")
    watch.stop()
    root.destroy()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def create_startup_shortcut():
    startup_folder = winshell.startup()
    script_path = os.path.realpath(sys.argv[0])
    shortcut_path = os.path.join(startup_folder, "PDFAutoPrint.lnk")
    
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = script_path
    shortcut.WorkingDirectory = os.path.dirname(script_path)
    shortcut.save()

# Tkinter UI
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Print")

        # Frame for the monitoring status and green orb
        top_frame = Frame(root)
        top_frame.pack(fill=BOTH, expand=True)

        self.monitoring_status = StringVar()
        self.monitoring_status.set("Monitoring folder: " + DOWNLOADS_FOLDER)
        
        self.canvas = Canvas(top_frame, width=20, height=20)
        self.canvas.create_oval(5, 5, 15, 15, fill="green")
        self.canvas.pack(side="left", padx=10, pady=10)
        
        self.monitoring_label = Label(top_frame, textvariable=self.monitoring_status, relief="sunken", anchor="w")
        self.monitoring_label.pack(side="left", fill=BOTH, expand=True)

        # Frame for the log status
        bottom_frame = Frame(root)
        bottom_frame.pack(fill=BOTH, expand=True)
        
        self.log_status = StringVar()
        self.log_status.set("")

        self.log_label = Label(bottom_frame, textvariable=self.log_status, relief="sunken", anchor="w")
        self.log_label.pack(side="bottom", fill=BOTH)
        
        root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
    
    def minimize_to_tray(self):
        self.root.withdraw()
        threading.Thread(target=setup_tray).start()
    
    def update_monitoring_status(self, message):
        self.root.after(0, self._update_monitoring_status, message)

    def _update_monitoring_status(self, message):
        self.monitoring_status.set(message)
        self.root.update_idletasks()

    def update_log_status(self, message):
        self.root.after(0, self._update_log_status, message)

    def _update_log_status(self, message):
        self.log_status.set(message)
        self.root.update_idletasks()

    def on_closing(self):
        watch.stop()
        self.root.quit()


# Function to rename file
def rename_file(file_path):
    pattern_with_parenthesis = r"(\(.*?\))\.[a-zA-Z0-9]+\.pdf\.part$"
    pattern_without_parenthesis = r"\.[a-zA-Z0-9]+\.pdf\.part$"
    dir_name, file_name = os.path.split(file_path)
    match_with_parenthesis = re.search(pattern_with_parenthesis, file_name)
    match_without_parenthesis = re.search(pattern_without_parenthesis, file_name)
    if match_with_parenthesis:
        new_file_name = file_name[:match_with_parenthesis.start()] + match_with_parenthesis.group(1) + '.pdf'
    elif match_without_parenthesis:
        new_file_name = file_name[:match_without_parenthesis.start()] + '.pdf'
    else:
        return ""
    return os.path.join(dir_name, new_file_name)

# Watchdog event handler
class Handler(FileSystemEventHandler):
    @staticmethod
    def on_moved(event):
        try:
            # firefox
            file = rename_file(event.src_path)
            
            if '.pdf' in file and os.path.exists(file):
                print(f"Received file: {file}. Starting print thread...")
                #app.update_log_status(f"Received file: {file}. Starting print thread...")
                threading.Thread(target=process_file, args=(file,)).start()
                return

            # chrome 
            other_file = event.src_path.removesuffix('.crdownload')

            if '.pdf' in other_file and os.path.exists(other_file):
                print(f"Received file: {other_file}. Starting print thread...")
                #app.update_log_status(f"Received file: {other_file}. Starting print thread...")
                threading.Thread(target=process_file, args=(other_file,)).start()
        
        except Exception as exception:
            print("Error: ", exception)
            #app.update_log_status(f"Error: {exception}")

class OnMyWatch:
    WATCHED_FOLDER = DOWNLOADS_FOLDER

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.WATCHED_FOLDER, recursive=False)
        self.observer.start()

    
    def stop(self):
        self.observer.stop()
        self.observer.join()

# Process file function
def process_file(file_path):
    try:
        time.sleep(1)  # Wait for file to be completely written

        active_window = gw.getActiveWindow()

        print_file(file_path)

        time.sleep(0.5)

        if active_window:
            try:
                active_window.activate()
            except Exception as e:
                error_message = str(e)
                if "Error code from Windows: 0" not in error_message:
                    print("Error while refocusing the window:", e)
        else:
            print("No active window found.")

        time.sleep(4.5)  # Give some time to finish printing
        os.remove(file_path)
        print(f"Deleted {file_path}. ")
        #app.update_log_status(f"Deleted {file_path}.")

    except Exception as e:
        print("Error while printing:", e)
        #app.update_log_status(f"Error while printing: {e}")

# Print file function
def print_file(file_path):
    print(f"Now printing {file_path}...")
    call([ACRO_DIR, "/s", "/h", "/p", file_path])

if __name__ == '__main__':
    ACRO_DIR = find_acrobat()
    if not ACRO_DIR:
        messagebox.showerror("Error", "Adobe Acrobat not found on the system. Please ensure it is installed.")
        sys.exit(1)

    # Create startup shortcut
    create_startup_shortcut()

    root = Tk()
    app = App(root)
    watch = OnMyWatch()
    threading.Thread(target=watch.run).start()
    root.withdraw()  # Start minimized
    setup_tray()
    root.mainloop()
