from datetime import *
import pandas as pd
import os


report_df = pd.DataFrame(columns=['Location', 'Customer', 'Day', 'Week', 'Month']).set_index(['Location', 'Customer'])
customers_dict = {'CVG': ['ABX', 'ATI AMZ', 'ATI DHL', 'DHL', '21 Air', 'Frontier', 'Cargojet', 'Aerologic', 'Commutair', 'Air Georgian', 'MESA', 'Southwest', 'Republic', 'United Airlines', 'National', 'Northern Air Cargo', 'Swift', 'Sky West'], 'ILN': ['ABX', 'ATI', 'Atlas', 'Sun Country'], 'MIA': ['ABX', 'ATI', 'DHL', 'Cargojet', 'Northern Air Cargo', 'Amerijet', 'Sunwing']}
for loc in ['CVG', 'MIA', 'ILN']:
    ranges = [1, 7, 31]
    lengths = ['Day', 'Week', 'Month']
    for i, r in enumerate(ranges):
        events_df = pd.read_excel('Line Maintenance Report.xlsx', sheet_name='Events', index_col=0)
        end = (date.today() - timedelta(1))
        start = (date.today() - timedelta(r))
        directory = r'G:\Line Reports\Reports\{} Daily Events'.format(loc) + '\\'
        for file in os.listdir(directory):
            try:
                if file[:3] != loc:
                    continue
                d = pd.to_datetime(file[4:14], format='%Y-%m-%d').date()
                if start <= d <= end:
                    df = pd.read_excel(directory + file, sheet_name='Events', index_col=0)
                    df = df.fillna(0)
                    events_df += df
                    print(file)
            except:
                continue
        for customer in customers_dict[loc]:
            turns = events_df.loc[customer, 'Turns']
            report_df.loc[(loc, customer), lengths[i]] = turns
writer = pd.ExcelWriter(r'G:\Finance\Work\ACF\Mary Jo Reports\Report {}.xlsx'.format(date.today()))
report_df.to_excel(writer, sheet_name='Report')
writer.save()
print(report_df)
