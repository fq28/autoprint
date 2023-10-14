import os
import time
import win32print
import win32api
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from collections import deque

# Folder to monitor (can be adjusted as per need)
WATCHED_FOLDER = 'E:\\Picture'
PRINT_QUEUE = deque()
PRINTING = False

class OnMyWatch:
    # Set the directory
    WATCHED_FOLDER = WATCHED_FOLDER

    # Define event handler
    class Handler(FileSystemEventHandler):
        @staticmethod
        def on_moved(event):
            file = event.src_path.removesuffix(".crdownload")

            # Windows1 10:if '.pdf' in event.src_path and os.path.exists(event.src_path):  # Check if a pdf file and if it still exists
            if '.pdf' in event.src_path:
                print(f"Received file: {file}. Adding to print queue...")
                PRINT_QUEUE.append(file)
                process_queue()


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

def process_queue():
    global PRINTING, PRINT_QUEUE
    if PRINTING:
        return  # Return if a print job is already in progress

    if PRINT_QUEUE:
        PRINTING = True
        file_path = PRINT_QUEUE.popleft()  # Fetch the next file from the queue
        
        print(f"Waiting for file {file_path} to settle...")
        time.sleep(1)
        
        print(f"Now printing {file_path}...")
        print_file(file_path)
        
        print(f"Finished printing {file_path}. Deleting now...")
        time.sleep(5)
        os.remove(file_path)
        
        PRINTING = False
        process_queue()  # Check the queue again


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
    process_queue()


if __name__ == '__main__':
    watch = OnMyWatch()
    watch.run()
