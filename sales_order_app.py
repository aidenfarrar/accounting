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


def generate_sales_orders():
    global directory
    directory += '\\'
    events_df = None
    for file in os.listdir(directory):
        customer_info_dict = pd.read_excel('Sales Order Info.xlsx', sheet_name=None)
        if file.startswith('Wings data'):
            events_df['Count'] = events_df['Count'].fillna(0)
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

            window = Toplevel(root)
            name_label = Label(window, text='Name: ').grid(row=0)
            employee_number_label = Label(window, text='Employee Number: ').grid(row=1)
            prepared_by_entry = Entry(window)  # input('Prepared by: ')
            employee_number_entry = Entry(window)  # input('Employee Number: ')
            prepared_by.grid(row=0, column=1)
            employee_number.grid(row=1, column=1)
            submit_button = Button(window, text='Submit', command=window.quit, )
            submit_button.grid(row=3)
            window.mainloop()
            prepared_by = prepared_by_entry.get()
            employee_number = employee_number_entry.get()
            print(prepared_by, employee_number)
            date_prepared = date.today()
            posting_date = date_prepared  # pd.to_datetime(input('Posting Date (Input as MM/DD/YYYY): '), format='%m/%d/%Y')

            sales_order_num = 123456  # int(input('Last Sales Order Number: '))
            for customerName, df in customer_info_dict.items():
                try:
                    # indexes = [x for x in maintenance_tracking_sheet['WORK_ORDER_NUMBER'] if x in sales_order_info['ABX Orders']]
                    customer_rows = maintenance_tracking_sheet[
                        maintenance_tracking_sheet['WORK_ORDER_NUMBER'].isin(df['Work Order Numbers'])]
                    raw_customer_rows = raw_labor_data[raw_labor_data['WORK_ORDER_NUMBER'].isin(df['Work Order Numbers'])]
                except:
                    continue

                for name in events_df['Customer']:
                    if name in customerName:
                        event_rows = events_df[events_df['Customer'] == name]
                        break

                sales_order_form_workbook = load_workbook('Blank Sales Order.xlsx')
                sales_order_form = sales_order_form_workbook['Sales Order Form']
                SAP_sheet = sales_order_form_workbook['SAP Upload']
                Labor_data_sheet = sales_order_form_workbook['Labor Data']
                Events_data_sheet = sales_order_form_workbook['Events Data']

                for r in dataframe_to_rows(raw_customer_rows, index=False, header=False):
                    Labor_data_sheet.append(r)

                for r in dataframe_to_rows(event_rows, index=False, header=False):
                    Events_data_sheet.append(r)

                daily_billable_hours = sum(customer_rows['BILLABLE_TIME'])
                daily_total_hours = sum(customer_rows['ACTUAL_TIME'])

                sales_order_form['A7'] = customerName
                sales_order_form['D7'] = prepared_by
                sales_order_form['E7'] = date_prepared
                sales_order_form['G7'] = int(employee_number)
                sales_order_form['G5'] = sales_order_num

                for i, line in enumerate(df['Billing Address']):
                    if type(line) == float:
                        break
                    sales_order_form['A{}'.format(i + 9)] = line
                    sales_order_form['D{}'.format(i + 9)] = line

                # note = ''  # input('Billing note: ')
                customer_reference = ''  # input('Project Number: ')
                customer_number = df['Customer number'][0]

                desc = ''
                for i, line in enumerate(df['Project Description']):
                    if type(line) == float:
                        break
                    if line == 'Date Range':
                        line = str(start_date)[:-9] + ' - ' + str(end_date)[:-9]
                        sales_order_form['G{}'.format(i + 9)] = line
                    else:
                        sales_order_form['G{}'.format(i + 9)] = line
                    desc = desc + ' ' + line

                for i, description in enumerate(df['Description of Service']):
                    hours = 0
                    if type(description) == float:
                        break
                    if df[description][0] == 'Events':
                        try:
                            hours = event_rows[event_rows['Event'] == description]['Count'].iloc[0]
                        except IndexError:
                            pass
                        work_orders = customer_rows[customer_rows['DESCRIPTION'].isin(df.loc[1:, description])]
                        hours += sum(work_orders['BILLABLE_TIME'])
                    else:
                        work_orders = customer_rows[customer_rows['DESCRIPTION'].isin(df[description])]
                        hours += sum(work_orders['BILLABLE_TIME'])
                    rate = df['Rate Per Hour/ Events'][i]
                    cost_center = df['Cost Center'][i]
                    account = df['Account'][i]

                    sales_order_form['A{}'.format(i + 16)] = description
                    sales_order_form['E{}'.format(i + 16)] = hours
                    sales_order_form['F{}'.format(i + 16)] = rate
                    sales_order_form['H{}'.format(i + 16)] = cost_center
                    sales_order_form['I{}'.format(i + 16)] = account

                    try:
                        # hours = float(sales_order_form['E{}'.format(i + 16)].value)
                        sales_order_form['G{}'.format(i + 16)] = rate * hours
                        SAP_sheet.append(
                            [1, 3500, date_prepared, posting_date, customer_number, sales_order_num, customer_reference,
                             description, cost_center, '', account, 'Credit', hours, hours * rate, '', '', '', 'O0',
                             'OH0140000', desc])
                    except TypeError:
                        continue

                for i, approver in enumerate(df['Approver(s) Printed Name']):
                    if type(approver) == float:
                        break
                    sales_order_form['A{}'.format(i + 32)] = approver
                    # sales_order_form['G{}'.format(i + 32)] = date_prepared
                    try:
                        sales_order_form['F{}'.format(i + 32)] = df['Employee Number'][i]
                    except TypeError:
                        continue

                sales_order_form['E13'] = df['Payment Terms'][0]

                sales_order_form_workbook.save(directory + 'Sales Order for {}.xlsx'.format(customerName))
        else:
            try:
                print(file)
                df = pd.read_excel(directory + file, sheet_name='Events', skiprows=1)
            except:
                continue
            if events_df is None:
                events_df = df
            else:
                events_df['Count'] += df['Count']
    # events_df.to_excel(directory + 'events_sheet.xlsx')


root = Tk()
root.title('Sales Order and SAP Upload Generator')
root.config(bg='white')

file_select_button = Button(root, text="Select Folder", command=browse_files)
file_select_button.pack()

generate_sales_orders_button = Button(root, text='Generate Sales Orders', command=generate_sales_orders)
generate_sales_orders_button.pack()

root.mainloop()
