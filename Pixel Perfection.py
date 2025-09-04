import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import ctypes as ct

# Hide the console window
ct.windll.user32.ShowWindow(ct.windll.kernel32.GetConsoleWindow(), 0)

# Function to filter and get active windows with meaningful titles
def get_active_windows():
    windows = {}
    
    def enum_windows_callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd).strip():
            title = win32gui.GetWindowText(hwnd)
            windows[title] = hwnd
    
    win32gui.EnumWindows(enum_windows_callback, None)
    return windows

# Function to set window size
def set_window_size(window_title, width, height):
    try:
        hwnd = active_windows.get(window_title)
        if not hwnd:
            messagebox.showerror("Error", "Window not found.")
            return
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOP,
            0, 0,  # x and y positions remain unchanged
            width, height,
            win32con.SWP_NOMOVE | win32con.SWP_NOZORDER
        )
        messagebox.showinfo("Success", f"Window resized to {width}x{height}.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to resize window: {e}")

# Function to refresh the window list
def refresh_windows():
    global active_windows
    active_windows = get_active_windows()
    window_selector['values'] = list(active_windows.keys())

# Function to clear all selections
def clear_selections():
    window_selector.set("")
    width_entry.delete(0, tk.END)
    height_entry.delete(0, tk.END)
    predefined_sizes.set("")

# Predefined sizes grouped by aspect ratio
PREDEFINED_SIZES = {
    "16:9": [
        "256x144 (144p)", "426x240 (240p)", "480x360 (360p)", 
        "640x360 (nHD)", "854x480 (FWVGA)", "960x540 (qHD)", 
        "1024x576 (WSVGA)", "1280x720 (HD)", "1366x768 (FWXGA)", 
        "1600x900 (HD+)", "1920x1080 (Full HD | 1080p)", 
        "2560x1440 (Quad HD | QHD)", "3200x1800 (QHD+)", 
        "3840x2160 (4K UHD)", "5120x2880 (5K)", 
        "7680x4320 (8K UHD)"
    ],
    "21:9": [
        "2560x1080 (UltraWide HD)", "3440x1440 (UltraWide QHD)", 
        "5120x2160 (UltraWide 4K)", "7680x3240 (UltraWide 8K)", 
        "10240x4320 (UltraWide 10K)", 
        "1920x800 (Cinema 21:9)", "2880x1200 (UltraWide 2.4:1)", 
        "3840x1600 (UltraWide WQHD)", "4320x1800 (UltraWide UHD)", 
        "5160x2160 (UltraWide 21+1/2:9)", "5760x2400 (UltraWide+)", 
        "6880x2880 (UltraWide Max)", "7680x3200 (UltraWide UHD+)", 
        "8640x3600 (UltraWide 21+3/5:9)"
    ],
    "16:10": [
        "1280x800 (WXGA)", "1920x1200 (WUXGA)"
    ],
    "4:3": [
        "640x480 (480p)", "800x600 (SVGA)", "1024x768 (XGA)", 
        "1600x1200 (UXGA)"
    ],
    "5:4": [
        "1280x1024 (SXGA)"
    ],
    "17:9": [
        "2048x1080 (2K)"
    ]
}

# Initialize Tkinter app
app = tk.Tk()
app.title("Pixel Perfection")
app.geometry("380x450")
app.resizable(True, True)

# Active windows dictionary
active_windows = get_active_windows()

# Window selection dropdown
tk.Label(app, text="Select Window").pack(pady=5)
window_selector = ttk.Combobox(app, values=list(active_windows.keys()), state="readonly", width=50)
window_selector.pack(pady=5)

# Predefined sizes dropdown
tk.Label(app, text="Predefined Sizes").pack(pady=5)
predefined_sizes = ttk.Combobox(app, state="readonly", width=50)
size_options = []
for ratio, sizes in PREDEFINED_SIZES.items():
    size_options.append(f"-- {ratio} --")
    size_options.extend(sizes)
predefined_sizes['values'] = size_options
predefined_sizes.pack(pady=5)

# Input fields for custom width and height
tk.Label(app, text="Custom Width").pack(pady=5)
width_entry = tk.Entry(app)
width_entry.pack(pady=5)

tk.Label(app, text="Custom Height").pack(pady=5)
height_entry = tk.Entry(app)
height_entry.pack(pady=5)

# Button to apply resizing
def apply_resizing():
    selected_window = window_selector.get()
    selected_size = predefined_sizes.get()
    width, height = None, None

    if selected_size and "--" not in selected_size:
        size_split = selected_size.split()[0].split("x")
        width, height = int(size_split[0]), int(size_split[1])
    elif width_entry.get() and height_entry.get():
        try:
            width, height = int(width_entry.get()), int(height_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid width or height input.")
            return
    else:
        messagebox.showerror("Error", "Please select a predefined size or input custom dimensions.")
        return

    if selected_window:
        height+= 35 # Adjustment for Windows 35 pixel fixed top border
        set_window_size(selected_window, width, height)
    else:
        messagebox.showerror("Error", "No window selected.")

tk.Button(app, text="Apply", command=apply_resizing, bg="green", fg="white").pack(pady=10)

# Button to refresh window list
tk.Button(app, text="Refresh Windows", command=refresh_windows).pack(pady=5)

# Button to clear selections
tk.Button(app, text="Clear Selections", command=clear_selections).pack(pady=5)

# Button to close the app
tk.Button(app, text="Exit", command=app.quit, bg="red", fg="white").pack(pady=10)

# Run the app
refresh_windows()  # Populate the initial window list
app.mainloop()
