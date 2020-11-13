import pandas as pd
from openpyxl import load_workbook
import os
from tkinter import *
from tkinter import filedialog
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import date, timedelta
import MySQLdb

hostname = 'localhost'
username = 'root'
password = 'fF3&CKw@8Xop6ysxRY^2Td!@yTVKSabg'
database = 'billing'
myConnection = MySQLdb.connect(host=hostname, user=username, passwd=password, db=database)


def doQuery(conn, query):
    cur = conn.cursor()
    cur.execute(query)
    return cur.fetchall()


def browse_files():
    return filedialog.askopenfilename(initialdir="/", title="Select Wings File")


def browse_directory():
    return filedialog.askdirectory(initialdir="/", title="Select Save Location")


def index_match(df, col1, search_col):
    return df[df[col1].isin(search_col)]


def value_match(df, col1, value):
    return df[df[col1] == value]


def generate_sales_orders(loc, start, end):
    start = pd.to_datetime(start, format='%m/%d/%Y').date()
    end = pd.to_datetime(end, format='%m/%d/%Y').date()
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
            df = pd.read_excel(directory + file, sheet_name='Events', index_col=0)  # , usecols=['Customer', 'Turns', 'Borescopes', 'Meet and Greet'])
            # labor = pd.read_excel(directory + file, sheet_name='Labor')
            df = df.fillna(0)
            events_df = events_df.add(df, fill_value=0)
    customer_info_dict = pd.read_excel('{} Sales Order Info.xlsx'.format(loc), sheet_name=None)
    d = r'G:\Line Reports\Reports\Wings Daily Data'
    for file in os.listdir(d):
        if file.endswith('.xlsx'):
            wings_file = d + '\\' + file
            break
    save_path = browse_directory()

    maintenance_tracking_sheet = pd.read_excel(wings_file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'BILLABLE_TIME', 'WORK_DATE', 'ACTUAL_TIME', 'DESCRIPTION'])
    maintenance_tracking_sheet['WORK_DATE'] = pd.to_datetime(maintenance_tracking_sheet['WORK_DATE'], format='%Y-%m-%d %H:%M:%S')
    D = {'BILLABLE_TIME': 'sum', 'ACTUAL_TIME': 'sum', 'WORK_ORDER_NUMBER': 'first'}
    maintenance_tracking_sheet.groupby(['DESCRIPTION']).agg(D)
    raw_labor_data = pd.read_excel(wings_file, sheet_name='Labor')

    employee_number = employee_number_entry.get()
    [prepared_by] = doQuery(myConnection, f"SELECT * FROM my_users WHERE employee_num = {employee_number};")
    sales_order_num = 1  # int(sales_order_num_entry.get())
    date_prepared = date.today()
    posting_date = date_prepared  # pd.to_datetime(input('Posting Date (Input as MM/DD/YYYY): '), format='%m/%d/%Y')
    loc = location.get()
    for [customer, payment_terms, customer_number, billing_street, billing_city, billing_state, billing_zip] in doQuery(myConnection, f"SELECT cust_name, payment_terms, customer_number, billing_street, billing_city, billing_state, billing_zip FROM customers;"):
        wos = doQuery(myConnection, f"SELECT wo_number FROM work_orders WHERE company = '{customer}';")
        if wos:
            wos = [x for x in wos[0]]
        customer_rows = index_match(maintenance_tracking_sheet, 'WORK_ORDER_NUMBER', wos)
        raw_customer_rows = index_match(raw_labor_data, 'WORK_ORDER_NUMBER', wos)

        event_rows = None
        for name in events_df.index:
            if name in customer:
                event_rows = events_df.loc[name]
                break

        if event_rows is None:
            print('Something wrong with {} event rows'.format(customer))
            continue

        sales_order_form_workbook = load_workbook('Blank Sales Order.xlsx')
        sales_order_form = sales_order_form_workbook['Sales Order Form']
        SAP_sheet = sales_order_form_workbook['SAP Upload']
        Labor_data_sheet = sales_order_form_workbook['Labor Data']
        Events_data_sheet = sales_order_form_workbook['Events Data']

        sales_order_num += 1
        sales_order_form['A7'] = customer
        sales_order_form['D7'] = prepared_by[0]
        sales_order_form['E7'] = date_prepared
        sales_order_form['G7'] = int(employee_number)
        sales_order_form['G5'] = sales_order_num
        sales_order_form['E13'] = payment_terms
        desc = ''

        for r in dataframe_to_rows(raw_customer_rows, index=False, header=False):
            Labor_data_sheet.append(r)

        for r in dataframe_to_rows(event_rows.to_frame().T, index=False):
            Events_data_sheet.append(r)

        for i, line in enumerate([billing_street, billing_city + ", " + billing_state + ", " + billing_state]):
            sales_order_form['A{}'.format(i + 9)] = line
            sales_order_form['D{}'.format(i + 9)] = line

        sales_order_form['G9'] = start.strftime(format='%m/%d/%Y') + ' - ' + end.strftime(format='%m/%d/%Y')
        row_delta = 0
        for [descriptions, service, fixed_rate, event_price, cost_center, account_num, use_events] in doQuery(myConnection, f"SELECT GROUP_CONCAT(work_description), extended_descriptions.service, fixed_rate, event_price, cost_center, account_num, use_events FROM extended_descriptions LEFT JOIN extended_services ON extended_services.service = extended_descriptions.service AND extended_services.customer = extended_descriptions.customer AND extended_services.location = extended_descriptions.location WHERE extended_descriptions.location = '{loc}' AND extended_descriptions.customer = '{customer}' GROUP BY extended_descriptions.service;"):
            print(customer, loc, service, fixed_rate, event_price, cost_center, account_num, use_events)
            fixed_rate, event_price = float(fixed_rate), float(event_price)
            work_orders = index_match(customer_rows, 'DESCRIPTION', descriptions.split(','))
            hours = sum(work_orders['BILLABLE_TIME'])
            print(hours, fixed_rate, event_price)
            if bool(use_events):
                try:
                    event_count = events_df.loc[customer, service]
                except KeyError:
                    continue
                sales_order_form['A{}'.format(row_delta + 16)] = service
                sales_order_form['E{}'.format(row_delta + 16)] = event_count
                sales_order_form['F{}'.format(row_delta + 16)] = event_price
                sales_order_form['G{}'.format(row_delta + 16)] = event_price * event_count
                sales_order_form['H{}'.format(row_delta + 16)] = cost_center
                sales_order_form['I{}'.format(row_delta + 16)] = account_num
                SAP_sheet.append(
                    [1, 3500, date_prepared, posting_date, customer_number, sales_order_num, None, service, cost_center,
                     None, account_num, 'Credit', event_count, event_count * event_price, None, None, None, 'O0', 'OH0140000', desc])
                row_delta += 1
                continue

            sales_order_form['A{}'.format(row_delta + 16)] = service
            sales_order_form['E{}'.format(row_delta + 16)] = hours
            sales_order_form['F{}'.format(row_delta + 16)] = fixed_rate
            sales_order_form['G{}'.format(row_delta + 16)] = fixed_rate * hours
            sales_order_form['H{}'.format(row_delta + 16)] = cost_center
            sales_order_form['I{}'.format(row_delta + 16)] = account_num

            SAP_sheet.append([1, 3500, date_prepared, posting_date, customer_number, sales_order_num, None, service, cost_center, None, account_num, 'Credit', hours, hours * fixed_rate, None, None, None, 'O0', 'OH0140000', desc])
            row_delta += 1
        
        [approvers] = doQuery(myConnection, f"SELECT first_approval, second_approval, third_approval FROM customers WHERE cust_name = '{customer}';")
        for i, approver in enumerate(approvers):
            [name] = doQuery(myConnection, f"SELECT user_name FROM my_users WHERE employee_num = {approver};")
            try:
                sales_order_form['A{}'.format(i + 32)] = name
            except ValueError:
                sales_order_form['A{}'.format(i + 32)] = name[0]
            except:
                continue
            sales_order_form['F{}'.format(i + 32)] = approver
        # writer = pd.ExcelWriter(save_path + '\\Line Changes for {}.xlsx'.format(customer))
        # writer.save()
        sales_order_form_workbook.save(save_path + '\\Sales Order for {}.xlsx'.format(customer))
    # updated_labor_df = pd.read_excel(wings_file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'BILLABLE_TIME', 'WORK_DATE', 'ACTUAL_TIME'])
    # updated_labor_df['WORK_DATE'] = pd.to_datetime(updated_labor_df['WORK_DATE'], format='%Y-%m-%d').dt.date
    writer = pd.ExcelWriter(save_path + '\\events_sheet.xlsx')
    events_df.to_excel(writer, sheet_name='Events')
    writer.save()


root = Tk()
root.title('Sales Order and SAP Upload Generator')
root.config(bg='white')

employee_number_label = Label(root, text='Employee Number: ', bg='White').grid(row=0)
employee_number_entry = Entry(root)  # input('Employee Number: ')
employee_number_entry.grid(row=0, column=1)
sales_order_num_entry_label = Label(root, text='Last Sales Order Number: ', bg='White').grid(row=1)
sales_order_num_entry = Entry(root)
sales_order_num_entry.grid(row=1, column=1)

start_date_label = Label(root, text='Start Date (In format MM/DD/YYYY): ', bg='White').grid(row=2)
start_date = Entry(root)
start_date.insert(0, (date.today() - timedelta(1)).strftime('%m/%d/%Y'))
start_date.grid(row=2, column=1)

end_date_label = Label(root, text='End Date (In format MM/DD/YYYY): ', bg='White').grid(row=3)
end_date = Entry(root)
end_date.insert(0, date.today().strftime('%m/%d/%Y'))
end_date.grid(row=3, column=1)

# location_label = Label(root, text='Location: ', bg='White').grid(row=5)
location = StringVar()
location.set('CVG')
location_list = OptionMenu(root, location, 'CVG', 'ILN', 'MIA')
location_list.grid(row=4)

# file_select_button = Button(root, text="Select Wings Data", command=browse_files)
# file_select_button.grid(row=6)

generate_sales_orders_button = Button(root, text='Generate Sales Orders', command=lambda: generate_sales_orders(location.get(), start_date.get(), end_date.get()))
generate_sales_orders_button.grid(row=4, column=1)

root.mainloop()
