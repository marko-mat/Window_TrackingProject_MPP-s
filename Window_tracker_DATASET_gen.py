import time
import ctypes
import psutil
import csv
from collections import defaultdict
from ctypes import wintypes
from datetime import datetime


# Getting window name
def get_active_window_process():
    user32 = ctypes.windll.user32
    pid = wintypes.DWORD()  # Correct way to use wintypes
    hwnd = user32.GetForegroundWindow()  # Get handle to the foreground window
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))  # Get process ID of the active window
    process = psutil.Process(pid.value)
    return process.name()


# appending data to csv
def write_to_csv(window_times, session_id, productivefactor, filename='window_time_tracking.csv'):
    file_exists = False
    try:
        with open(filename, mode='r'):
            file_exists = True
    except FileNotFoundError:
        file_exists = False

    with open(filename, mode='a', newline='') as file:
        writer = csv.writer(file)
                                          #write header for newfiles
        if not file_exists:
            writer.writerow(['Session ID', 'Application', 'Total Time (seconds)', 'Productive Factor'])

                                          # write with session id
        for window, total_time in window_times.items():
            writer.writerow([session_id, window, f'{total_time:.2f}', productivefactor])

    print(f"\nTime tracking data saved to {filename} for session {session_id}" + productivefactor)


                                          # Time tracking
def monitor_active_window():
    active_window = None
    start_time = time.time()
    window_times = defaultdict(float)

                                     # unique session ID using timestamp
    session_id = datetime.now().strftime('%Y%m%d_%H%M%S')

    try:
        while True:
            new_window = get_active_window_process()

            if new_window != active_window:
                end_time = time.time()
                if active_window:
                    elapsed_time = end_time - start_time
                    window_times[active_window] += elapsed_time
                    print(f"App '{active_window}' was active for {elapsed_time:.2f} seconds.")

                active_window = new_window
                start_time = time.time()

            time.sleep(1)

    except KeyboardInterrupt:
        productivity = input("\n")
        for window, total_time in window_times.items():
            print(f"App '{window}' total time: {total_time:.2f} seconds")

        # Write data to CSV file with session ID
        write_to_csv(window_times, session_id,productivity)


if __name__ == "__main__":
    print("Please remember, the FIRST number you type in the cosole, is your how productive your session is... \nPlease "
          "input a number between 0 and 100 before you close the program!")
    monitor_active_window()
