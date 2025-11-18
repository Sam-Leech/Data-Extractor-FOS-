#!/usr/bin/env python3
import sys
import ctypes
import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# Hide the console window (Windows only)
if sys.platform == "win32":
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# ---------------- DARK MODE COLORS ---------------- #
BG = "#1e1e1e"         # Dark background
FG = "#ffffff"         # White text
BTN_BG = "#3a3a3a"     # Button background
BTN_FG = "#ffffff"     # Button text
ACCENT = "#4CAF50"     # Green action button


def parse_filename(filename):
    """
    Parse filenames like:
    AFPC02852862-Tony Duffy-Advantage Finance.pdf
    """
    name = filename
    if name.lower().endswith('.pdf'):
        name = name[:-4]

    # Normalize spacing around hyphens
    name = re.sub(r'\s*-\s*', '-', name)

    parts = [p.strip() for p in name.split('-') if p.strip()]

    if len(parts) < 3:
        return None, None, None

    ref = parts[0]
    lender = parts[-1]
    client = " ".join(parts[1:-1])

    return ref, client, lender


def generate_excel(folder_path, log_box):
    if not os.path.isdir(folder_path):
        messagebox.showerror("Error", f"Folder not found:\n{folder_path}")
        return

    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    if not files:
        messagebox.showinfo("No PDFs", "No PDF files found in the selected folder.")
        return

    rows = []

    log_box.insert(tk.END, f"Found {len(files)} PDF files...\n")

    for f in files:
        ref, client, lender = parse_filename(f)

        if any(v is None for v in [ref, client, lender]):
            log_box.insert(tk.END, f"WARNING: Could not parse: {f}\n")

        rows.append({
            'Reference Number': ref if ref else '',
            'Client Name': client if client else '',
            'Lender': lender if lender else '',
            'Original Filename': f
        })

    df = pd.DataFrame(rows, columns=['Reference Number', 'Client Name', 'Lender', 'Original Filename'])

    out_xlsx = os.path.join(folder_path, "client_data.xlsx")
    df.to_excel(out_xlsx, index=False)

    log_box.insert(tk.END, f"\nExcel created:\n{out_xlsx}\n")
    log_box.insert(tk.END, "Done!\n")
    messagebox.showinfo("Success", f"Excel created:\n{out_xlsx}")


def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_selected)


# ------------ Splash Screen ------------ #

# ------------ Splash Screen ------------ #

def show_splash(root):
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    splash.configure(bg="#1e1e1e")

    # Splash window size
    w, h = 400, 200

    # Center splash window
    ws = splash.winfo_screenwidth()
    hs = splash.winfo_screenheight()
    x = int((ws / 2) - (w / 2))
    y = int((hs / 2) - (h / 2))
    splash.geometry(f"{w}x{h}+{x}+{y}")

    label = tk.Label(
        splash,
        text="FOS Client Extractor\nLoading...",
        font=("Arial", 16, "bold"),
        bg="#1e1e1e",
        fg="white"
    )
    label.pack(expand=True)

    # After 2 seconds → hide splash, show main window in SAME spot
    root.after(2000, lambda: finish_splash(root, splash, x, y))


def finish_splash(root, splash, x, y):
    splash.destroy()

    # Set main window to appear where the splash was
    root.geometry(f"+{x}+{y}")

    root.deiconify()


# ---------------- GUI ---------------- #

root = tk.Tk()
root.withdraw()  # Hide main window at startup

show_splash(root)  # Show splash screen

root.title("FOS Client Extractor")
root.geometry("650x480")
root.configure(bg=BG)


# Folder selection
folder_label = tk.Label(root, text="Select PDF Folder:", font=("Arial", 11), bg=BG, fg=FG)
folder_label.pack(pady=5)

folder_entry = tk.Entry(root, width=60, font=("Arial", 10), bg="#2a2a2a", fg=FG, insertbackground=FG)
folder_entry.pack(pady=5)

browse_btn = tk.Button(root, text="Browse...", command=browse_folder, width=12,
                       bg=BTN_BG, fg=BTN_FG, activebackground="#555")
browse_btn.pack(pady=5)

# Run button
run_btn = tk.Button(
    root,
    text="Generate Excel",
    command=lambda: generate_excel(folder_entry.get(), log_box),
    width=20,
    height=2,
    bg=ACCENT,
    fg="white",
    activebackground="#45a049"
)
run_btn.pack(pady=10)

# Log box
log_label = tk.Label(root, text="Log Output:", font=("Arial", 11), bg=BG, fg=FG)
log_label.pack()

log_box = scrolledtext.ScrolledText(
    root, width=75, height=12, font=("Arial", 9),
    bg="#2a2a2a", fg=FG, insertbackground=FG
)
log_box.pack(pady=5)

root.mainloop()
