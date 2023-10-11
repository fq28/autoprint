import os
import time
import win32print
import win32api
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Folder to monitor (can be adjusted as per need)
WATCHED_FOLDER = 'E:\\Picture'

class OnMyWatch:
    # Set the directory
    WATCHED_FOLDER = WATCHED_FOLDER

    # Define event handler
    class Handler(FileSystemEventHandler):
        @staticmethod
        def on_created(event):
            if '.pdf' in event.src_path and os.path.exists(event.src_path):  # Check if a pdf file and if it still exists
                print(f"Received file: {event.src_path}. Waiting for file to settle...")
                time.sleep(1)  # wait for 5 seconds or adjust as necessary
                print(f"Now printing {event.src_path}...")
                print_file(event.src_path)
                print(f"Finished printing {event.src_path}. Deleting now...")
                time.sleep(5)  # Give some time to finish printing (can be adjusted as per need)
                os.remove(event.src_path)

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = OnMyWatch.Handler()
        self.observer.schedule(event_handler, self.WATCHED_FOLDER, recursive=False)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            self.observer.stop()
            print("Observer stopped.")
        self.observer.join()

def print_file(file_path):
    """
    Print the file using default printer.
    """
    printer_name = win32print.GetDefaultPrinter()
    win32api.ShellExecute(
        0,
        "print",
        file_path,
        '/d:"%s"' % printer_name,
        ".",
        0
    )

if __name__ == '__main__':
    watch = OnMyWatch()
    watch.run()
