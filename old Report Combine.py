import pandas as pd
from openpyxl import load_workbook
import os
from datetime import date
from tkinter import *
from tkinter import filedialog
from datetime import date, timedelta
from openpyxl.utils.dataframe import dataframe_to_rows


def browse_files():
    directory = filedialog.askdirectory(initialdir="/", title="Select a File")


def index_match(df, col1, search_col):
    return df[df[col1].isin(search_col)]


def value_match(df, col1, value):
    return df[df[col1] == value]


def generate_sales_orders(start, end):
    start = pd.to_datetime(start, format='%m/%d/%Y').date()
    end = pd.to_datetime(end, format='%m/%d/%Y').date()
    events_df = None
    directory = r'G:\Finance\Work\Accts Receivable\CUSTOMER INVOICING\~Daily Line Mtc Activity\CVG\\'
    for file in os.listdir(directory):
        try:
            d = pd.to_datetime(file[34:42], format='%Y%m%d').date()
        except:
            continue
        if start <= d <= end:
            print(file[34:42])
            try:
                df = pd.read_excel(directory + file, sheet_name='Events', skiprows=1, usecols=['Customer', 'Event', 'Count'])
                old_labor = pd.read_excel(directory + file, sheet_name='Labor')
                df = df.fillna(0)
                print(file)
            except:
                continue
            sheet_date = old_labor.loc[0, 'WORK_DATE'].date()
            if events_df is None:
                events_df = df
                old_labor_df = old_labor
            else:
                events_df['Count'] += df['Count']
                # events_df['{}'.format(sheet_date)] = df['Count']
                old_labor_df = old_labor_df.append(old_labor)
    writer = pd.ExcelWriter(r'G:\Finance\Work\Aiden Files\\' + 'events_sheet.xlsx')
    events_df.to_excel(writer, sheet_name='Events')
    old_labor_df.to_excel(writer, sheet_name='Labor')
    writer.save()


root = Tk()
root.title('Sales Order and SAP Upload Generator')
root.config(bg='white')

start_date_label = Label(root, text='Start Date (In format MM/DD/YYYY): ', bg='White').grid(row=0)
start_date = Entry(root)
start_date.insert(0, (date.today() - timedelta(1)).strftime('%m/%d/%Y'))
start_date.grid(row=0, column=1)

end_date_label = Label(root, text='End Date (In format MM/DD/YYYY): ', bg='White').grid(row=1)
end_date = Entry(root)
end_date.insert(0, date.today().strftime('%m/%d/%Y'))
end_date.grid(row=1, column=1)

file_select_button = Button(root, text="Select Folder", command=browse_files)
file_select_button.grid(row=2)

generate_sales_orders_button = Button(root, text='Generate Combined Report', command=lambda: generate_sales_orders(start_date.get(), end_date.get()))
generate_sales_orders_button.grid(row=2, column=1)

root.mainloop()
