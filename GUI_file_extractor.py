import tkinter as tk
import os
from tkinter import filedialog

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)
    list_files(folder_path)

def list_files(folder_path):
    try:
        files = os.listdir(folder_path)
    except Exception as e:
        print(e)
    file_list.delete(0, tk.END)
    for i, file in enumerate(files):
        if i % 2 != 0:
            file_list.insert(i, file)
        else:
            file_list.insert(i, file)

def on_right_click(event):
    context_menu.post(event.x_root, event.y_root)

def copy_file_name():
    selected_index = file_list.curselection()
    if selected_index:
        file_name = file_list.get(selected_index)
        root.clipboard_clear()
        root.clipboard_append(file_name)

def open_file_context():
    selected_index = file_list.curselection()
    if selected_index:
        file_name = file_list.get(selected_index)
        file_path = os.path.join(folder_entry.get(), file_name)
        os.startfile(file_path)

def browse_file():
    file_path = filedialog.askopenfilename()
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)
    print(file_path)

root = tk.Tk()
root.title("Tkinter Example")

frame = tk.Frame(root)
frame.pack(fill="both", expand=True)

folder_browse_button = tk.Button(frame, text="Browse Folder", command=browse_folder)
folder_browse_button.grid(row=0, column=0, padx=5, pady=5)

folder_entry = tk.Entry(frame, width=200)
folder_entry.grid(row=0, column=1, padx=5, pady=5)

file_browse_button = tk.Button(frame, text="Browse Keywords", command=browse_file)
file_browse_button.grid(row=1, column=0, padx=5, pady=5)

file_entry = tk.Entry(frame, width=200)
file_entry.grid(row=1, column=1, padx=5, pady=5)

run_button = tk.Button(frame, text="Run")
run_button.grid(row=2, column=0, pady=5)

file_list = tk.Listbox(frame, selectmode="single", width=200, height=200)
file_list.grid(row=3, columnspan=2, padx=5, pady=5, sticky="nsew")
file_list.bind("<Button-3>", on_right_click)

context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy file name", command=copy_file_name)
context_menu.add_command(label="Open file", command=open_file_context)

root.mainloop()
