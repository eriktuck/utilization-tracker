import pandas as pd
import numpy as np

july_hours = 'data/july2020.xlsx'
employee_crosswalk = 'data/employee_crosswalk.csv'

df = pd.read_excel(july_hours,
                   header=20)

dx = pd.read_csv(employee_crosswalk, index_col=0)
employee_map = dx.to_dict()

drop_cols = [col for col in df.columns if col.startswith('Unnamed')]

df.drop(drop_cols, axis=1, inplace=True)

def fill_down(df, columns_to_fill):
    value_dict = dict(zip(columns_to_fill, [None]*len(columns_to_fill)))
    filled_df = pd.DataFrame(columns=df.columns)
    for _, row in df.iterrows():
        for column in columns_to_fill:        
            if row[column] is np.nan:
                row[column] = value_dict[column]
            else:
                value_dict[column] = row[column]
        filled_df = filled_df.append(row, sort=False)
    return filled_df


df = fill_down(df, ['Project', 'Employee'])

df.dropna(axis='rows', subset=['Date'], inplace=True)

df.drop(['Project.1', 'UDT10', 'Comments'], axis=1, inplace=True)

mapper = {'Project': 'Activity Name',
        'Employee': 'User Name',
        'Date': 'Entry Date',
        'Hours': 'Hours Worked'}
df.rename(columns=mapper, inplace=True)

df['User Name'] = df['User Name'].map(employee_map['User Name'])

df = df[['User Name', 'Entry Date', 'Activity Name', 'Hours Worked']]

df.sort_values(['User Name', 'Entry Date'], inplace=True)

df['Time Off Hrs'] = 0

df.to_csv('data/paste_into_google_sheet.csv', index=False)