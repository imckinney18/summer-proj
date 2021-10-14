import xlrd
import pandas as pd
import numpy as np
import os

EXCEL_PATH = './blastdata.xlsx' # need the name of the excel file

sheet_names = [ # need the names of the individual sheets (trials)
    'AVERAGE',
    'Internal Break Down',
    '2 Bang',
    'Single Strand Roll Up Interior',
    '450gr Interior',
    '300gr Interior',
    'Jelly Interior',
    '3 Strand Exterior',
    'Carl G GUNNER',
    'CARL G AG',
    'AT4'
]


def read_sheet_by_name(sheet_name):
    return pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, engine='openpyxl')

def read_column_into_numpy(sheet, column_name):
    return sheet[column_name].dropna().to_numpy()

def extract_average_max_cumulative(sheet_name, column_names = None):
    sheet = read_sheet_by_name(sheet_name)

    if column_names is None:
        column_names = list(set(sheet.columns).difference(['Time (msec)', 'PEAK PRESSURE', 'AVG', '19.1']))
        print(column_names)

    master_array = np.array([read_column_into_numpy(sheet, name) for name in column_names])

    means = list(np.mean(master_array, axis=1))
    maxes = list(np.max(master_array, axis=1))
    cumus = list(np.sum(master_array, axis=1))

    all_three = tuple(zip(means, maxes, cumus))

    mapping = tuple(zip(column_names, all_three))

    values_to_return = {
        mapping[i][0] : mapping[i][1] for i in range(len(mapping))
    }
    
    return values_to_return

sheet_name = 'Carl G GUNNER'
# sheet_name = '2 Bang'
trial_names = [f'CHARGE {i}' for i in range(1, 16)]
# trial_names = list(set([i for i in range(1, 97)]).difference(set([7, 14, 16])))
d = extract_average_max_cumulative(sheet_name, trial_names)

print(d)