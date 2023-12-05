import tkinter as tk
from tkinter import ttk, messagebox
from bacnet import bacnetInitialize, deviceScan, readAv, readBv
import subprocess

# Classes
class CustomButton(tk.Button):
    def __init__(self, master=None, **kwargs):
        tk.Button.__init__(self, master, **kwargs)
        self.configure(bg="#3C8FDD", fg="white", font=("Arial", 10, "bold"), width=15)


# Functions
def device_scan():
    try:
        deviceScan(bacnet)  # Call the readAv function from bacnet.py
        messagebox.showinfo("Device Scan", "Devices scanned successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error scanning devices: {e}")

def read_bv_values():
    try:
        readBv(bacnet)  # Call the readAv function from bacnet.py
        messagebox.showinfo("Read BV", "BV values read successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading AV: {e}")

def read_av_values():
    try:
        readAv(bacnet)  # Call the readAv function from bacnet.py
        messagebox.showinfo("Read AV", "AV values read successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading AV: {e}")

def adjust_settings():
    try:
        subprocess.Popen(['notepad.exe', 'settings.ini'])  # Opens settings.ini with Notepad on Windows
    except Exception as e:
        messagebox.showerror("Error", f"Error opening settings.ini: {e}")


# Main
bacnet = bacnetInitialize()

# Main Window
root = tk.Tk()
root.title("BACnet Scan Utility")
root.geometry("300x300")
root.configure(bg="white")

# Create a Notebook (tabbed interface)
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Frame for Device Scan tab
device_scan_frame = ttk.Frame(notebook)
notebook.add(device_scan_frame, text='Device Scan')
# device_scan_frame.configure(bg="white")

# Buttons for Device Scan tab
button_device_scan = CustomButton(device_scan_frame, text="Device Scan", command=device_scan)
button_device_scan.pack(pady=10)

button_read_av = CustomButton(device_scan_frame, text="Read AV", command=read_av_values)
button_read_av.pack(pady=10)

button_read_bv = CustomButton(device_scan_frame, text="Read BV", command=read_bv_values)
button_read_bv.pack(pady=10)

button_adjust_settings = CustomButton(device_scan_frame, text="Adjust Settings", command=adjust_settings)
button_adjust_settings.pack(pady=10)


# Frame for Settings Adjustment tab
settings_frame = ttk.Frame(notebook)
notebook.add(settings_frame, text='Adjust Settings')

# Button to adjust settings.ini
button_adjust_settings = CustomButton(settings_frame, text="Adjust Settings", command=adjust_settings)
button_adjust_settings.pack(pady=10)

root.mainloop()