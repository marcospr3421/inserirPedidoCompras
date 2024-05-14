import tkinter as tk
from tkinter import filedialog


def open_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not filename:
        print("No file selected.")
    else:
        setOrder(filename)

def setOrder(filename):
    # Your existing code here
    pass

root = tk.Tk()
root.title("Order Processing")

open_file_button = tk.Button(root, text="Open Excel File", command=open_file)
open_file_button.pack()

root.mainloop()