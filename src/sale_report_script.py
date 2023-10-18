# %% [markdown]
# # Automate Read Excel

# %%
import datetime
import pandas as pd
import os
import xlwings as xw
# %%

path = r'D:\Python_Project\src\data\sales_data'
today = datetime.date.today()
year = today.year-1

xlsx_file_lists = []
for root,dirs,files in os.walk(path):
    for name in files:
        file_path =os.path.join(root,name)
        print("file_path",file_path)
        xlsx_file_lists.append(file_path)
print("xlsx_file_lists",xlsx_file_lists)

# %%
df = pd.read_excel(xlsx_file_lists[0])
df
#อ่านไฟล์ที่วนออกมาได้ดังนี้

# %%
msg = r'D:\Python_Project\src\data\sales_data\2021\April.xlsx'
#split แตกชื่อไฟล์ตามที่ต้องการ
print(msg.split("\\"))
print(msg.split("\\")[-2])


# %%
for f in xlsx_file_lists:
    if f.split("\\")[-2] == "2021":
        print(f)

# %%
msg = r'D:\Python_Project\src\data\sales_data\2021\April.xlsx'
msg2 = r'D:\Python_Project\src\data\sales_data\2022\April.xlsx'
#split แตกชื่อไฟล์ตามที่ต้องการ หาเดือน เมษาของทุกปี
print(msg.split("\\"))
print(msg.split("\\")[-2])

print(msg2.split("\\"))
print(msg2.split("\\")[-2])

# %%
for f in xlsx_file_lists:
    if f.split('\\')[-1] == 'April.xlsx':
        print(f)

# %%
import datetime
today = datetime.date.today()
year = today.year-1
#print(year)

for f in xlsx_file_lists:
    if f.split("\\")[-2] == str(year):
        if f.split("\\")[-1].split(".")[0] == "April":
            print(f)
            df = pd.read_excel(f)
df

# %% [markdown]
# #### Example Pivot Table

# %%
df_lists = []
for f in xlsx_file_lists:
    df = pd.read_excel(f)
    df_lists.append(df)
df_lists

# %%
df_summary = pd.concat(df_lists)
df_summary

# %%
pivot = pd.pivot_table(df_summary,index="transaction_date",columns="store",values="amount",aggfunc="sum")

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


template = xw.Book(r"D:\Python_Project\src\data\sale_template.xlsx")

app = xw.apps.active
sheet = template.sheets["summary"]
sheet["A1"].value = summary_monthly

pivote = template.sheets["pivot"]
pivote["A1"].value = pivot

#add picture
sheet_report = template.sheets["report"]
sheet_report['A1'].value = "Summary by month"
#sheet_report['A1'].font.size = 24
sheet_report['A1'].api.Font.Bold = True
plot= sheet_report.pictures.add(fig,top=sheet["A3"].top,left=sheet["A3"].left)
plot.width = plot.width*0.8
plot.height = plot.height*0.8

template.save(f"export\summary_sale_report_{date_file_name}.xlsx")
template.close()
app.kill()


