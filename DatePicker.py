from datetime import date
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import messagebox

def cal_done():
    top.withdraw()
    root.quit()

root = tk.Tk()
root.withdraw() # keep the root window from appearing

top = tk.Toplevel(root, background='#856ff8')
top.title("GmailtoDrive")
top.geometry("300x200")
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()

def get_date():
    #Create a Label
    ttk.Label(top, width=18, background='red', text='Choose From Date').pack(padx=10, pady=10)
    cal1 = DateEntry(top, width=12, background='yellow',
                    foreground='black', borderwidth=2)
    cal1.pack(padx=10, pady=10)

    ttk.Label(top, width=18, background='red', text='Choose To Date').pack(padx=10, pady=10)
    cal = DateEntry(top, width=12, background='darkblue',
                    foreground='black', borderwidth=2, maxdate=date.today())
    cal.pack(padx=10, pady=10)

    ttk.Button(top, text="Proceed", command=cal_done).pack(padx=10, pady=10, side='right')

    top.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    return (cal1.get_date(), cal.get_date())