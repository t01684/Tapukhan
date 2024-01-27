import tkinter as tk
from tkinter import ttk
import openpyxl

def load_data():
    path = "E:/codefirst.io/Tkinter Excel App/user_data.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

    print(name, age, subscription_status, employment_status)

    # insert row into excel sheet
    path = "E:/codefirst.io/Tkinter Excel App/user_data.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    # insert row into treeview
    treeview.insert('', tk.END, values=row_values)

    # clear the values
    name_entry.delete(0, "end")
    age_spinbox.delete(0, "end")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

root = tk.Tk()
root.title("Excel App")

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

combo_list = ["Subscribed", "Not Subscribed", "Other"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert Row")

widgets_frame.grid(row=0, column=0, padx=20, pady=10)

ttk.Label(widgets_frame, text="Name:").grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")
name_entry = ttk.Entry(widgets_frame)
name_entry.grid(column=1, row=1)
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.insert(0, "Age")
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Employed", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame, orient="horizontal")
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

tree_frame = ttk.Frame(frame)
tree_frame.grid(row=0, column=1, pady=10)
tree_scroll = ttk.Scrollbar(tree_frame)
tree_scroll.pack(side="right", fill="y")

columns = ("Name", "Age", "Subscription", "Employment")
treeview = ttk.Treeview(tree_frame, show="headings",
                        yscrollcommand=tree_scroll.set, columns=columns, height=13)
for col in columns:
    treeview.heading(col, text=col)
    treeview.column(col, width=100)

treeview.pack()
tree_scroll.config(command=treeview.yview)

# Load data initially
load_data()

# Run the Tkinter main loop
root.mainloop()

