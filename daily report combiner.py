import os
import numpy as np
import pandas as pd
from tkinter import *
from datetime import date, timedelta


def generate_report(loc, start, end):
    directory = r'G:\Line Reports\Reports\{} Daily Events'.format(loc) + '\\'
    print(directory)
    d = r'G:\Line Reports\Reports\Wings Daily Data'
    for file in os.listdir(d):
        if file.endswith('.xlsx'):
            wings_file = d + '\\' + file
            break
    raw_labor_data = pd.read_excel(wings_file, sheet_name='Labor')
    raw_labor_data['WORK_DATE'] = pd.to_datetime(raw_labor_data['WORK_DATE'], format='%Y-%m-%d %H:%M:%S')
    raw_labor_data['DATE'] = [x.date() for x in pd.to_datetime(raw_labor_data['WORK_DATE'], format='%Y-%m-%d')]
    mask = (start <= raw_labor_data['DATE']) & (raw_labor_data['DATE'] <= end)
    raw_labor_data = raw_labor_data[mask]
    events_df = pd.read_excel('Line Maintenance Report.xlsx', sheet_name='Events', index_col=0)
    for file in os.listdir(directory):
        if file[:3] == loc and file[14:] == ' Line Maintenance Report.xlsx':
            try:
                d = pd.to_datetime(file[4:14], format='%Y-%m-%d').date()
            except:
                continue
        else:
            continue
        if start <= d <= end:
            print(file)
            df = pd.read_excel(directory + file, sheet_name='Events', index_col=0)
            df = df.fillna(0)
            events_df += df
    events_df = events_df[(events_df.T != 0).any()]
    events_df = events_df.loc[:, (events_df != 0).any(axis=0)]
    writer = pd.ExcelWriter(r'G:\Line Reports\Reports\Combined Reports\{} {} - {} combined sheet.xlsx'.format(loc, start, end))
    events_df.to_excel(writer, sheet_name='Events')
    raw_labor_data.to_excel(writer, sheet_name='Labor', index=False)
    writer.save()


root = Tk()
root.title('Daily Report Combiner')
root.config(bg='white')


start_date_label = Label(root, text='Start Date (In format MM/DD/YYYY): ', bg='White').grid(row=0)
start_date = Entry(root)
start_date.insert(0, (date.today() - timedelta(1)).strftime('%m/%d/%Y'))
start_date.grid(row=0, column=1)

end_date_label = Label(root, text='End Date (In format MM/DD/YYYY): ', bg='White').grid(row=1)
end_date = Entry(root)
end_date.insert(0, date.today().strftime('%m/%d/%Y'))
end_date.grid(row=1, column=1)


location = StringVar()
location.set('CVG')
location_list = OptionMenu(root, location, 'CVG', 'ILN', 'MIA')
location_list.grid(row=2)
generate_sales_orders_button = Button(root, text='Generate Report', command=lambda: generate_report(location.get(), pd.to_datetime(start_date.get(), format='%m/%d/%Y').date(), pd.to_datetime(end_date.get(), format='%m/%d/%Y').date()))
generate_sales_orders_button.grid(row=2, column=1)

root.mainloop()
