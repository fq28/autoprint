import os
import time
import win32print
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
from subprocess import call

# Folder to monitor
WATCHED_FOLDER = 'E:\\Picture'

# program that will be used for printing
SUMATRA_DIR = "C:\\Users\\fqdej\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe" 

# Mapping of filename prefixes to printer names
PRINTER_MAPPING = {
    'PACKINGSLIP': 'ET-7700 Series(Network)',  # Replace with actual printer name
    'DEFAULT': 'ZDesigner GK420d'  # Default printer for labels
}

class OnMyWatch:
    WATCHED_FOLDER = WATCHED_FOLDER

    class Handler(FileSystemEventHandler):
        @staticmethod
        def on_created(event):
            file = event.src_path.removesuffix(".crdownload")

            # check if the file is a pdf
            if '.pdf' in file and os.path.exists(file):
                print(f"Received file: {file}. Starting print thread...")
                threading.Thread(target=process_file, args=(file,)).start()

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

def process_file(file_path):
    try:
        time.sleep(1)  # Wait for file to be completely written

        filename = os.path.basename(file_path)
        printer_name = PRINTER_MAPPING['DEFAULT']
        for prefix, printer in PRINTER_MAPPING.items():
            if filename.startswith(prefix):
                printer_name = printer
                break
        
        print(f"Now printing {file_path} on printer {printer_name}...")
        print_file(file_path, printer_name)

        time.sleep(5)  # Give some time to finish printing
        os.remove(file_path)
        print(f"Deleted {file_path}. ")

    except Exception as e:
        print("Eror while printing:", e)


def printer_does_exists(printer_name):
    # Get the list of all installed printers
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]

    # Check if the specified printer is in the list of installed printers
    if printer_name not in printers:
        print(f"Error: Printer '{printer_name}' not found. Available printers are: {printers}")
        return False
    return True

def print_file(file_path, printer_name):
    """
    Print the file using the specified printer.
    """
    if printer_does_exists(printer_name):
        call([SUMATRA_DIR, "-print-to", printer_name, "-silent", file_path])

if __name__ == '__main__':
    watch = OnMyWatch()
    watch.run()
