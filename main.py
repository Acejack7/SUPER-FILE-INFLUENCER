#!python3

import os
import tkinter as tk
from tkinter import ttk

# GUI
window = tk.Tk()
window.title("SUPER FILE INFLUENCER")


# FUNCTION - add new sentence to the file
def change_file():
    filepath = user_filepath.get()

    if os.path.isfile(filepath) is False:
        filepath_info.set("This is not the file. Please correct the filepath.")
        return

    with open(filepath, 'a', encoding='utf-8') as edit_file:
        edit_file.write('\n\n*** The new sentence. ***')
        edit_file.close()
        filepath_info.set("New sentence added! Click again to add another.")
        return


# Label for filepath
filepath_info = tk.StringVar()
filepath_info.set("Provide filepath:")
ttk.Label(window, textvariable=filepath_info).grid(column=0, row=0)

# Textbox for crop name
user_filepath = tk.StringVar()
user_filepath_entered = ttk.Entry(window, width=40, textvariable=user_filepath)
user_filepath_entered.grid(column=0, row=1)

# Action Button
action = ttk.Button(window, text="Add new sentence",
                    command=change_file)
action.grid(column=1, row=5)

if __name__ == "__main__":
    window.mainloop()
