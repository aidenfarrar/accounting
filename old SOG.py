import pandas as pd
from openpyxl import load_workbook
import os
from datetime import date
from tkinter import *
from tkinter import filedialog
from openpyxl.utils.dataframe import dataframe_to_rows

global directory


def browse_files():
    global directory
    directory = filedialog.askdirectory(initialdir="/", title="Select a File")


def index_match(df, col1, search_col):
    return df[df[col1].isin(search_col)]


def value_match(df, col1, value):
    return df[df[col1] == value]


def generate_sales_orders():
    global directory
    directory += '\\'
    events_df = None
    old_labor_df = None
    # updated_labor_df = None
    for file in os.listdir(directory):
        customer_info_dict = pd.read_excel('Sales Order Info.xlsx', sheet_name=None)
        if file.startswith('Wings'):
            print(events_df)
            for i, name in enumerate(events_df['Customer']):
                if name[-1] == ' ':
                    events_df.loc[i, 'Customer'] = name[:-1]

            maintenance_tracking_sheet = pd.read_excel(directory + file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'BILLABLE_TIME', 'WORK_DATE', 'ACTUAL_TIME', 'DESCRIPTION'])
            maintenance_tracking_sheet['WORK_DATE'] = pd.to_datetime(maintenance_tracking_sheet['WORK_DATE'], format='%Y-%m-%d %H:%M:%S')
            D = {'BILLABLE_TIME': 'sum', 'ACTUAL_TIME': 'sum', 'WORK_ORDER_NUMBER': 'first'}
            maintenance_tracking_sheet.groupby(['DESCRIPTION']).agg(D)
            raw_labor_data = pd.read_excel(directory + file, sheet_name='Labor')

            start_date = min(maintenance_tracking_sheet['WORK_DATE'])
            end_date = max(maintenance_tracking_sheet['WORK_DATE'])
            prepared_by = prepared_by_entry.get()
            employee_number = employee_number_entry.get()
            sales_order_num = int(sales_order_num_entry.get())
            date_prepared = date.today()
            posting_date = date_prepared  # pd.to_datetime(input('Posting Date (Input as MM/DD/YYYY): '), format='%m/%d/%Y')
            for customerName, df in customer_info_dict.items():
                try:
                    customer_rows = index_match(maintenance_tracking_sheet, 'WORK_ORDER_NUMBER', df['Work Order Numbers'])
                    # customer_rows = maintenance_tracking_sheet[
                    #     maintenance_tracking_sheet['WORK_ORDER_NUMBER'].isin(df['Work Order Numbers'])]
                    #     maintenance_tracking_sheet['WORK_ORDER_NUMBER'].isin(df['Work Order Numbers'])]
                    raw_customer_rows = index_match(raw_labor_data, 'WORK_ORDER_NUMBER', df['Work Order Numbers'])
                    # raw_customer_rows = raw_labor_data[raw_labor_data['WORK_ORDER_NUMBER'].isin(df['Work Order Numbers'])]
                    raw_old_customer_rows = index_match(old_labor_df, 'WORK_ORDER_NUMBER', df['Work Order Numbers'])
                except KeyError:
                    continue

                event_rows = None
                for name in events_df['Customer']:
                    if name in customerName:
                        event_rows = events_df[events_df['Customer'] == name]
                        break

                if event_rows is None:
                    print('Something wrong with event rows')
                    continue
                sales_order_form_workbook = load_workbook('Blank Sales Order.xlsx')
                sales_order_form = sales_order_form_workbook['Sales Order Form']
                SAP_sheet = sales_order_form_workbook['SAP Upload']
                Labor_data_sheet = sales_order_form_workbook['Labor Data']
                Events_data_sheet = sales_order_form_workbook['Events Data']

                sales_order_num += 1
                sales_order_form['A7'] = customerName
                sales_order_form['D7'] = prepared_by
                sales_order_form['E7'] = date_prepared
                sales_order_form['G7'] = int(employee_number)
                sales_order_form['G5'] = sales_order_num
                sales_order_form['E13'] = df['Payment Terms'][0]
                customer_number = df['Customer number'][0]
                desc = ''

                for r in dataframe_to_rows(raw_customer_rows, index=False, header=False):
                    Labor_data_sheet.append(r)

                for r in dataframe_to_rows(event_rows, index=False):
                    Events_data_sheet.append(r)

                for i, line in enumerate(df['Billing Address']):
                    if type(line) == float:
                        break
                    sales_order_form['A{}'.format(i + 9)] = line
                    sales_order_form['D{}'.format(i + 9)] = line

                for i, line in enumerate(df['Project Description']):
                    if type(line) == float:
                        break
                    if line == 'Date Range':
                        line = str(start_date)[:-9] + ' - ' + str(end_date)[:-9]
                        sales_order_form['G{}'.format(i + 9)] = line
                    else:
                        sales_order_form['G{}'.format(i + 9)] = line
                    desc = desc + ' ' + line

                line_changes_df = pd.DataFrame(columns=['Description', 'Old Hours', 'New Hours', 'Old Revenue', 'New Revenue', 'Old Cost', 'New Cost'])
                for i, description in enumerate(df['Description of Service']):
                    hours = 0
                    old_hours = 0
                    if type(description) == float:
                        break
                    if df[description][0] == 'Events':
                        try:
                            hours = value_match(event_rows, 'Event', description)['Count'].iloc[0]
                            # hours = event_rows[event_rows['Event'] == description]['Count'].iloc[0]
                            old_hours = hours
                        except IndexError:
                            pass
                        work_orders = index_match(customer_rows, 'DESCRIPTION', df.loc[1:, description])
                        old_work_orders = index_match(raw_old_customer_rows, 'DESCRIPTION', df.loc[1:, description])
                        # work_orders = customer_rows[customer_rows['DESCRIPTION'].isin(df.loc[1:, description])]
                        hours += sum(work_orders['BILLABLE_TIME'])
                        old_hours += sum(old_work_orders['BILLABLE_TIME'])
                    else:
                        work_orders = index_match(customer_rows, 'DESCRIPTION', df[description])
                        old_work_orders = index_match(raw_old_customer_rows, 'DESCRIPTION', df[description])
                        # work_orders = customer_rows[customer_rows['DESCRIPTION'].isin(df[description])]
                        hours += sum(work_orders['BILLABLE_TIME'])
                        old_hours += sum(old_work_orders['BILLABLE_TIME'])
                    rate = df['Rate Per Hour/ Events'][i]
                    cost_center = df['Cost Center'][i]
                    account = df['Account'][i]

                    sales_order_form['A{}'.format(i + 16)] = description
                    sales_order_form['E{}'.format(i + 16)] = hours
                    sales_order_form['F{}'.format(i + 16)] = rate
                    sales_order_form['H{}'.format(i + 16)] = cost_center
                    sales_order_form['I{}'.format(i + 16)] = account

                    try:
                        sales_order_form['G{}'.format(i + 16)] = rate * hours
                        SAP_sheet.append(
                            [1, 3500, date_prepared, posting_date, customer_number, sales_order_num, None,
                             description, cost_center, None, account, 'Credit', hours, hours * rate, None, None, None, 'O0',
                             'OH0140000', desc])
                    except TypeError:
                        continue
                    if old_hours != hours:
                        line_changes = [description, old_hours, hours, old_hours * rate, hours * rate, old_hours * 37.89, hours * 37.89]
                        line_changes_df = line_changes_df.append(pd.Series(line_changes, index=line_changes_df.columns), ignore_index=True)

                for i, approver in enumerate(df['Approver(s) Printed Name']):
                    if type(approver) == float:
                        break
                    sales_order_form['A{}'.format(i + 32)] = approver
                    try:
                        sales_order_form['F{}'.format(i + 32)] = df['Employee Number'][i]
                    except TypeError:
                        continue
                writer = pd.ExcelWriter(directory + 'Line Changes for {}.xlsx'.format(customerName))
                line_changes_df.to_excel(writer, index=False)
                writer.save()
                sales_order_form_workbook.save(directory + 'Sales Order for {}.xlsx'.format(customerName))
            updated_labor_df = pd.read_excel(directory + file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'BILLABLE_TIME', 'WORK_DATE', 'ACTUAL_TIME'])
            updated_labor_df['WORK_DATE'] = pd.to_datetime(updated_labor_df['WORK_DATE'], format='%Y-%m-%d').dt.date

        else:
            try:
                df = pd.read_excel(directory + file, sheet_name='Events', skiprows=1, usecols=['Customer', 'Event', 'Count'])
                old_labor = pd.read_excel(directory + file, sheet_name='Labor')
                df = df.fillna(0)
                print(file)
            except:
                continue
            try:
                sheet_date = old_labor.loc[0, 'WORK_DATE'].date()
            except AttributeError:
                sheet_date = pd.to_datetime(old_labor.loc[0, 'WORK_DATE'], format='%Y-%m-%d').date()
            if events_df is None:
                events_df = df
            else:
                events_df['Count'] += df['Count']
                events_df['{}'.format(sheet_date)] = df['Count']
            if old_labor_df is None:
                old_labor_df = old_labor
            else:
                old_labor_df = pd.concat([old_labor_df, old_labor], ignore_index=True)

    writer = pd.ExcelWriter(directory + 'events_sheet.xlsx')
    events_df.to_excel(writer, sheet_name='Events')
    writer.save()


root = Tk()
root.title('Sales Order and SAP Upload Generator')
root.config(bg='white')

name_label = Label(root, text='Name: ', bg='White').grid(row=0)
employee_number_label = Label(root, text='Employee Number: ', bg='White').grid(row=1)
prepared_by_entry = Entry(root)  # input('Prepared by: ')
employee_number_entry = Entry(root)  # input('Employee Number: ')
prepared_by_entry.grid(row=0, column=1)
employee_number_entry.grid(row=1, column=1)
sales_order_num_entry_label = Label(root, text='Last Sales Order Number: ', bg='White').grid(row=2)
sales_order_num_entry = Entry(root)
sales_order_num_entry.grid(row=2, column=1)

file_select_button = Button(root, text="Select Folder", command=browse_files)
file_select_button.grid(row=3)

generate_sales_orders_button = Button(root, text='Generate Sales Orders', command=generate_sales_orders)
generate_sales_orders_button.grid(row=3, column=1)

root.mainloop()
