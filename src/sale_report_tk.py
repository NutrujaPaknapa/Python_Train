# %% [markdown]
# # Example

# %%
import pandas as pd
import xlwings as xw
import datetime

def sale_report(path,save_file):

    import datetime
    import pandas as pd
    import os

    today = datetime.date.today()
    year = today.year-1

    xlxs_file_lists = []

    for root,dirs,files in os.walk(path):
        for name in files:
            file_path = os.path.join(root,name)
            if file_path.split("\\")[-2] == str(year): #change filter
                xlxs_file_lists.append(file_path)
    xlxs_file_lists

    # %%
    df_lists = []

    for f in xlxs_file_lists:
        df = pd.read_excel(f)
        df_lists.append(df)
    df_lists

    # %%
    df_summary = pd.concat(df_lists)
    df_summary

    # %%
    pivot = pd.pivot_table(df_summary,index="transaction_date",columns="store",values="amount",aggfunc="sum")
    pivot

    # %%
    summary_monthly = pivot.resample("M").sum()
    summary_monthly

    # %%
    import matplotlib
    fig = summary_monthly.plot(kind="bar",figsize=(20,12),fontsize=26,title="monthly sale summary").get_figure()

    # %%
    import xlwings as xw

    import datetime
    now = datetime.datetime.now()
    date_file_name = f'{str(now.date())}_{str(now.time()).split(".")[0].replace(":","_")}'


    #template = xw.Book(r"D:\My Documents\Desktop\python_office_11_OCT_2023\src\data\sale_template.xlsx")
    template = xw.Book(r"data\sale_template.xlsx")

    app = xw.apps.active
    sheet = template.sheets["summary"]
    sheet["A1"].value = summary_monthly

    pivot_page = template.sheets["pivot"]
    pivot_page["A1"].value = pivot

    #add picture
    sheet_report = template.sheets["report"]
    sheet_report["A1"].value = "Summary by month"
    sheet_report['A1'].font.size = 24
    sheet_report["A1"].api.Font.Bold = True
    plot= sheet_report.pictures.add(fig,top=sheet["A3"].top,left=sheet["A3"].left)
    plot.width = plot.width*0.8
    plot.height = plot.height*0.8

    template.save(save_file+".xlsx")
    template.close()
    app.kill()

def browse_files():
    global path

    file_name = filedialog.askdirectory()
    path = file_name     
    file_explorer_label.configure(text=path)

def save_as():
    global save_as_dir

    save_file = filedialog.asksaveasfilename(initialdir = "/",
                                          title = "save as a File",
                                          filetypes = (("excel files",
                                                        "*.xlsx*"),
                                                       ("all files",
                                                        "*.*")))

    save_as_dir = save_file
    save_as_label.configure(text=save_as_dir)

def create_report_button_clicked():
    """create prodution report"""
    try:
        global path,save_as_dir
        sale_report(path,save_as_dir)
        file_explorer_label.config(text="browse files")
        save_as_label.config(text="Save as")
        messagebox.showinfo("info","Done")
    except Exception as e:
        messagebox.showerror("Error!!",e)
    

import tkinter as tk
from tkinter import ttk 
from tkinter import filedialog
from tkinter import messagebox

#global variable
path = None
save_as_dir = None

# root window
window = tk.Tk()
window.title("sale report program")
window.geometry("640x120")
window.resizable(False,False)

# frame
frame = ttk.Frame(window)

# field options
options = {'padx':5,'pady':5}

# Browse label
file_explorer_label = ttk.Label(frame,text="Browse File",width = 85)
file_explorer_label.grid(column=0,row=0,**options)

# browse button
browse_button = ttk.Button(frame,text="browse files")
browse_button.grid(column=1,row=0,**options)
browse_button.configure(command=browse_files)

# save_as label
save_as_label = ttk.Label(frame,text="Save as",width = 85)
save_as_label.grid(column=0,row=1,**options)

# save as path button
save_as_label_button = ttk.Button(frame,text="file location")
save_as_label_button.grid(column=1,row=1,**options)
save_as_label_button.configure(command=save_as)
 
# create report button
create_report_button = ttk.Button(frame,text="CREATE!")
create_report_button.grid(columnspan=2,row=3,**options)
create_report_button.configure(command=create_report_button_clicked)

# add frame
frame.grid(pady=10,padx=10)

window.mainloop() #run

# %% [markdown]
# # Build

# %%
#!pyinstaller.exe --noconsole sale_report_tk.py


