import tkinter as tk
from tkinter import ttk

def clear(event):
    statusBar_value.set('')

def set_statusBar(event):
    statusBar_value.set(widget_name[event.widget])

root = tk.Tk()

statusBar_value = tk.StringVar()
statusBar_value.set('Status Bar...')

entry1 = ttk.Entry(root)
dummy =  ttk.Entry(root)
entry2 = ttk.Entry(root)

widget_name = {entry1:'Entry 1 has focus', entry2:'Entry 2 has focus'}

statusBar = ttk.Label(root, textvariable = statusBar_value)

entry1.grid()
dummy.grid()
entry2.grid()

statusBar.grid()

entry1.bind('<FocusIn>', set_statusBar)
entry1.bind('<FocusOut>', clear)

entry2.bind('<FocusIn>', set_statusBar)
entry2.bind('<FocusOut>', clear)

root.mainloop()
