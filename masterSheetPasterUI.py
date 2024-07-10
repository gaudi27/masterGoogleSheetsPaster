import tkinter as tk
from tkinter import messagebox
from masterSheet import paster
import webbrowser


def on_paster():
    master_url = master_url_entry.get()
    sheet_urls = sheet_urls_text.get("1.0", tk.END).strip().split("\n")
    if master_url and sheet_urls:
        try:
            paster(master_url, sheet_urls)
            messagebox.showinfo("Success", "Data has been successfully consolidated into the master sheet.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Input Error", "Please provide both the master sheet URL and at least one other sheet URL.")

def open_master_sheet():
    master_url = master_url_entry.get()
    if master_url:
        webbrowser.open(master_url)
    else:
        messagebox.showwarning("Input Error", "Please provide the master sheet URL.")

# Set up the main application window
root = tk.Tk()
root.title("Google Sheets Consolidator")

# Create and place the master sheet URL input
tk.Label(root, text="Master Google Sheet URL:").pack(pady=5)
master_url_entry = tk.Entry(root, width=50)
master_url_entry.pack(pady=5)

# Create and place the other sheets URL input
tk.Label(root, text="Other Google Sheets URLs (one per line):").pack(pady=5)
sheet_urls_text = tk.Text(root, width=100, height=30)
sheet_urls_text.pack(pady=5)

# Create and place the consolidate button
consolidate_button = tk.Button(root, text="Consolidate Sheets", command=on_paster)
consolidate_button.pack(pady=20)

open_master_button = tk.Button(root, text="Open Master Sheet", command=open_master_sheet)
open_master_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()


