import os
import pandas as pd
from tkinter import *
from datetime import date, timedelta, datetime
import openpyxl
from os import mkdir
from os.path import isdir
from math import isnan


def remove_excess(d):
    # d = d[(d.T != 0).any()]
    # print(d)
    d = d.loc[:, (d != 0).any(axis=0)]
    # print(d)
    return d


def sum_events(loc, start, end):
    delta = (end - start).days + 1
    print(delta)
    directory = r'G:\Line Reports\Reports\{} Daily Events'.format(loc) + '\\'
    print(directory)
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
            if events_df is None:
                events_df = df
            else:
                events_df += df
    events_df = remove_excess(events_df)
    return events_df, delta


def calculate_margin(df, start, end, loc, days):
    tech_rate_dict = {'CVG': 37.89, 'ILN': 34.60, 'MIA': 41.50}
    customers_dict = {'CVG': ['ABX', 'ATI AMZ', 'ATI DHL', 'DHL', '21 Air', 'Frontier', 'Cargojet', 'Aerologic', 'Commutair', 'Air Georgian', 'MESA', 'Southwest', 'Republic', 'United Airlines', 'National', 'Northern Air Cargo', 'Swift'], 'ILN': ['ABX', 'ATI', 'Atlas'], 'MIA': ['ABX', 'ATI', 'DHL', 'Cargojet', 'Northern Air Cargo', 'Amerijet', 'Sunwing']}
    rate = tech_rate_dict[loc]
    hours_sheet = '{} Hours Sheet.xlsx'.format(loc)
    prices = pd.read_excel(hours_sheet, sheet_name='Event Prices', index_col=0)
    prices = prices.fillna(0)
    event_descriptions = pd.read_excel(hours_sheet, sheet_name='Event Descriptions', index_col=[0, 1])
    fixed_descriptions = pd.read_excel(hours_sheet, sheet_name='Fixed Rate and Overhead Desc')
    d = r'G:\Line Reports\Reports\Wings Daily Data'
    for file in os.listdir(d):
        if file.endswith('.xlsx'):
            hours_df = pd.read_excel(r'G:\Line Reports\Reports\Wings Daily Data\\' + file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'ACTUAL_TIME', 'DESCRIPTION', 'WORK_DATE'])
            break
    hours_df['WORK_DATE'] = pd.to_datetime(hours_df['WORK_DATE'], format='%Y-%m-%d')
    hours_df['WORK_DATE'] = hours_df['WORK_DATE'].dt.date
    mask = (hours_df['WORK_DATE'] >= start) & (hours_df['WORK_DATE'] <= end)
    hours_df = hours_df[mask]
    D = {'ACTUAL_TIME': 'sum', 'WORK_ORDER_NUMBER': 'first'}
    hours_df = hours_df.groupby(['DESCRIPTION']).agg(D)
    events = list(df)
    companies = df.index
    loc_companies = customers_dict[loc]
    margin_wb = openpyxl.Workbook()
    margin_ws = margin_wb.active
    margin_ws.title = 'Margin Report'
    margin_ws.append(['Company', 'Margin', 'Revenue', 'Expense'])
    for company in companies:
        # empty = True
        ws = margin_wb.create_sheet(company)
        ws.title = company
        ws.append(['Company', 'Description', 'Margin', 'Count', 'Revenue', 'Expense'])
        total_revenue = 0
        total_expense = 0
        for event in events:
            count = df.loc[company, event]
            price = prices.loc[company, event]
            if price > 0 and count > 0:
                print(company)
                if days < count:
                    days = count
                # empty = False
                if event == 'Borescopes':
                    revenue = price * count + 150*(company == 'ATI AMZ' and event == 'Turns')*days
                    expense = 750 * count
                    margin = (revenue - expense) / revenue
                    line = [company, event, margin, count, revenue, expense]
                    ws.append(line)
                    total_revenue += revenue
                    total_expense += expense
                    continue
                print(company, event, price, count)
                try:
                    desc = event_descriptions.loc[(company, event), 'Description']
                    hours = hours_df.loc[desc, 'ACTUAL_TIME']
                except KeyError:
                    hours = 0
                # print(desc, ': ', hours)
                revenue = price * count + 150*(company == 'ATI AMZ' and event == 'Turns')*days
                expense = hours * rate
                try:
                    margin = (revenue - expense) / revenue
                except ZeroDivisionError:
                    margin = 0
                line = [company, event, margin, count, revenue, expense]
                ws.append(line)
                total_revenue += revenue
                total_expense += expense
        # if 'ABX' in company:
        #     company = 'ABX'
        try:
            company_col = fixed_descriptions[company]
        except KeyError:
            margin_wb.remove(ws)
            continue
        for desc in company_col:
            if type(desc) == int or type(desc) == float:
                if isnan(desc):
                    break
                rate = desc
            elif type(desc) == str:
                try:
                    hours = hours_df.loc[desc, 'ACTUAL_TIME']
                except KeyError:
                    continue
                revenue = hours * rate
                expense = hours * 37.89
                margin = (revenue - expense) / revenue
                line = [company, desc, margin, hours, revenue, expense]
                ws.append(line)
                # empty = False
                total_revenue += revenue
                total_expense += expense
        # if empty:
        #     margin_wb.remove(ws)
        #     continue
        try:
            margin = (total_revenue - total_expense) / total_revenue
        except ZeroDivisionError:
            margin = 0
        line = [company, margin, total_revenue, total_expense]
        margin_ws.append(line)

    overhead_col = fixed_descriptions['Overhead']
    overhead_hours = 0
    for desc in overhead_col:
        if type(desc) == str:
            try:
                hours = hours_df.loc[desc, 'ACTUAL_TIME']
            except KeyError:
                hours = 0
            overhead_hours += hours
    margin_ws.append(['Overhead', 0, 0, overhead_hours*37.89])
    margin_ws['F1'] = 'Report Date: {} - {}'.format(start, end)
    margin_ws['F2'] = 'Station: {}'.format(loc)
    time_str = datetime.now().strftime("%m-%d-%Y %H%M%S")
    margin_wb.save(r'G:\Line Reports\Reports\Margin Reports\{}\{} {} - {} margin report {}.xlsx'.format(loc, loc, start, end, time_str))


def generate_report(loc, start, end):
    events, days = sum_events(loc, start, end)
    calculate_margin(events, start, end, loc, days)


root = Tk()
root.title('Margin Report Generator')
root.config(bg='white')

start_date_label = Label(root, text='Start Date (In format MM/DD/YYYY): ', bg='White').grid(row=0)
start_date = Entry(root)
start_date.insert(0, (date.today() - timedelta(2)).strftime('%m/%d/%Y'))
start_date.grid(row=0, column=1)

end_date_label = Label(root, text='End Date (In format MM/DD/YYYY): ', bg='White').grid(row=1)
end_date = Entry(root)
end_date.insert(0, (date.today() - timedelta(1)).strftime('%m/%d/%Y'))
end_date.grid(row=1, column=1)

location = StringVar()
location.set('CVG')
location_list = OptionMenu(root, location, 'CVG', 'ILN', 'MIA')
location_list.grid(row=2)
# file_select_button = Button(root, text="Select Directory", command=browse_files)
# file_select_button.grid(row=2)
# generate_sales_orders_button = Button(root, text='Generate Report', command=lambda: generate_report(path, pd.to_datetime(start_date.get(), format=r'%m/%d/%Y').date(), pd.to_datetime(end_date.get(), format=r'%m%/d%/Y').date()))
generate_sales_orders_button = Button(root, text='Generate Report', command=lambda: generate_report(location.get(), pd.to_datetime(start_date.get(), format='%m/%d/%Y').date(), pd.to_datetime(end_date.get(), format='%m/%d/%Y').date()))
generate_sales_orders_button.grid(row=2, column=1)

root.mainloop()
