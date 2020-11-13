import pandas as pd
from openpyxl import load_workbook
import os
from tkinter import *
from tkinter import filedialog
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import date, timedelta


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
    events_df = None
    directory = r'G:\Line Reports\Reports\{} Daily Events'.format(loc) + '\\'
    print(directory)
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
            if events_df is None:
                events_df = df
            else:
                events_df = events_df.add(df, fill_value=0)
    print(events_df)
    customer_info_dict = pd.read_excel('{} Sales Order Info.xlsx'.format(loc), sheet_name=None)
    # wings_file = browse_files()
    d = r'G:\Line Reports\Reports\Wings Daily Data'
    for file in os.listdir(d):
        if file.endswith('.xlsx'):
            wings_file = d + '\\' + file
            # print(wings_file)
            # hours_df = pd.read_excel(r'G:\Line Reports\Reports\Wings Daily Data\\' + file, sheet_name='Labor',
            #                          usecols=['WORK_ORDER_NUMBER', 'ACTUAL_TIME', 'DESCRIPTION', 'WORK_DATE'])
            break
    save_path = browse_directory()

    maintenance_tracking_sheet = pd.read_excel(wings_file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'BILLABLE_TIME', 'WORK_DATE', 'ACTUAL_TIME', 'DESCRIPTION'])
    maintenance_tracking_sheet['WORK_DATE'] = pd.to_datetime(maintenance_tracking_sheet['WORK_DATE'], format='%Y-%m-%d %H:%M:%S')
    D = {'BILLABLE_TIME': 'sum', 'ACTUAL_TIME': 'sum', 'WORK_ORDER_NUMBER': 'first'}
    maintenance_tracking_sheet.groupby(['DESCRIPTION']).agg(D)
    raw_labor_data = pd.read_excel(wings_file, sheet_name='Labor')

    start_date_1 = min(maintenance_tracking_sheet['WORK_DATE'])
    end_date_1 = max(maintenance_tracking_sheet['WORK_DATE'])
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
            # raw_old_customer_rows = index_match(old_labor_df, 'WORK_ORDER_NUMBER', df['Work Order Numbers'])
        except KeyError:
            continue

        event_rows = None
        for name in events_df.index:
            if name in customerName:
                event_rows = events_df.loc[name]
                break

        if event_rows is None:
            print('Something wrong with {} event rows'.format(customerName))
            continue

        print(name)
        # print(event_rows)
        print(event_rows.to_frame())

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

        for r in dataframe_to_rows(event_rows.to_frame().T, index=False):
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
                line = str(start_date_1)[:-9] + ' - ' + str(end_date_1)[:-9]
                sales_order_form['G{}'.format(i + 9)] = line
            else:
                sales_order_form['G{}'.format(i + 9)] = line
                sales_order_form['G{}'.format(i + 9)] = line
            desc = desc + ' ' + line

        # line_changes_df = pd.DataFrame(columns=['Description', 'Old Hours', 'New Hours', 'Old Revenue', 'New Revenue', 'Old Cost', 'New Cost'])
        for i, description in enumerate(df['Description of Service']):
            hours = 0
            # old_hours = 0
            if type(description) == float:
                break
            if df[description][0] == 'Events':
                try:
                    hours = event_rows.loc[description]
                    # hours = event_rows[event_rows['Event'] == description]['Count'].iloc[0]
                    # old_hours = hours
                except (IndexError, KeyError) as e:
                    print(e)
                    pass
                work_orders = index_match(customer_rows, 'DESCRIPTION', df.loc[1:, description])
                # old_work_orders = index_match(raw_old_customer_rows, 'DESCRIPTION', df.loc[1:, description])
                # work_orders = customer_rows[customer_rows['DESCRIPTION'].isin(df.loc[1:, description])]
                hours += sum(work_orders['BILLABLE_TIME'])
                # old_hours += sum(old_work_orders['BILLABLE_TIME'])
            else:
                work_orders = index_match(customer_rows, 'DESCRIPTION', df[description])
                # old_work_orders = index_match(raw_old_customer_rows, 'DESCRIPTION', df[description])
                # work_orders = customer_rows[customer_rows['DESCRIPTION'].isin(df[description])]
                hours += sum(work_orders['BILLABLE_TIME'])
                # old_hours += sum(old_work_orders['BILLABLE_TIME'])
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
            # if old_hours != hours:
            #     line_changes = [description, old_hours, hours, old_hours * rate, hours * rate, old_hours * 37.89, hours * 37.89]
            #     line_changes_df = line_changes_df.append(pd.Series(line_changes, index=line_changes_df.columns), ignore_index=True)

        for i, approver in enumerate(df['Approver(s) Printed Name']):
            if type(approver) == float:
                break
            sales_order_form['A{}'.format(i + 32)] = approver
            try:
                sales_order_form['F{}'.format(i + 32)] = df['Employee Number'][i]
            except TypeError:
                continue
        # writer = pd.ExcelWriter(save_path + '\\Line Changes for {}.xlsx'.format(customerName))
        # line_changes_df.to_excel(writer, index=False)
        # writer.save()
        sales_order_form_workbook.save(save_path+ '\\Sales Order for {}.xlsx'.format(customerName))
    updated_labor_df = pd.read_excel(wings_file, sheet_name='Labor', usecols=['WORK_ORDER_NUMBER', 'BILLABLE_TIME', 'WORK_DATE', 'ACTUAL_TIME'])
    updated_labor_df['WORK_DATE'] = pd.to_datetime(updated_labor_df['WORK_DATE'], format='%Y-%m-%d').dt.date
    writer = pd.ExcelWriter(save_path + '\\events_sheet.xlsx')
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

start_date_label = Label(root, text='Start Date (In format MM/DD/YYYY): ', bg='White').grid(row=3)
start_date = Entry(root)
start_date.insert(0, (date.today() - timedelta(1)).strftime('%m/%d/%Y'))
start_date.grid(row=3, column=1)

end_date_label = Label(root, text='End Date (In format MM/DD/YYYY): ', bg='White').grid(row=4)
end_date = Entry(root)
end_date.insert(0, date.today().strftime('%m/%d/%Y'))
end_date.grid(row=4, column=1)

# location_label = Label(root, text='Location: ', bg='White').grid(row=5)
location = StringVar()
location.set('CVG')
location_list = OptionMenu(root, location, 'CVG', 'ILN', 'MIA')
location_list.grid(row=5)

# file_select_button = Button(root, text="Select Wings Data", command=browse_files)
# file_select_button.grid(row=6)

generate_sales_orders_button = Button(root, text='Generate Sales Orders', command=lambda: generate_sales_orders(location.get(), start_date.get(), end_date.get()))
generate_sales_orders_button.grid(row=5, column=1)

root.mainloop()
