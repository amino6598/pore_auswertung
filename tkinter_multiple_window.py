# -*- coding: utf-8 -*-
"""
Created on Fri Nov 23 09:15:11 2018

@author: q463532
"""
import cv2
import numpy as np
import glob
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk


#root = tk.Tk()
#root.withdraw()
#root.focus_set()
#directory_name = filedialog.askdirectory(title='Ordner Auswählen')
#print(directory_name)


def open_directory():
    root.withdraw()
    directory_name = filedialog.askdirectory(title='Ordner Auswählen',parent=root)
    print(directory_name)
    root.deiconify()

def input_name():
    def callback():
        print(e.get())
        root.quit()
    e = ttk.Entry(root)
    NORM_FONT = ("Helvetica", 10)
    label = ttk.Label(root,text='Enter the name of the ROI', font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    e.pack(side = 'top',padx = 10, pady = 10)
    e.focus_set()
    b = ttk.Button(root, text = "OK", width = 10, command = callback)
    b.pack()

def close_window():
    root.quit()
    root.destroy()
        
root = tk.Tk()
#root.withdraw()
open_directory()
input_name()
root.mainloop()
close_window()