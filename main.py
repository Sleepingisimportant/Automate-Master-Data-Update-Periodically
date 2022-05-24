'''
Author: Sleepingisimportant (GitHub)
This program is to be used to automate the process of updating the master file periodically. For more information, please see "README.txt"
'''

import pandas as pd
from datetime import datetime
import os

# Here assuming that monthly update file name is with a certain pattern.
report_name_current_month = r"Monthly Update {}.xls".format(datetime.today().strftime("%m-%Y"))

# Paths of monthly Update file and master file.
update_file = os.path.join(os.path.dirname(__file__),
                           "Example/Monthly_update_report/{}".format(report_name_current_month))
master_file = os.path.join(os.path.dirname(__file__), "Example/Master.xls")

# Transform file content to dataframe
df_update = pd.read_excel(update_file, sheet_name="Sheet1", engine='openpyxl')
df_master = pd.read_excel(master_file, sheet_name="Sheet1", engine='openpyxl')

# Update Master with update file by ID
for i in range(len(df_update.index)):
    searched_ID = df_update['ID'][i]

    # If ID exists, update date to Master file
    if searched_ID in df_master.ID.values:
        value_to_be_update = df_update['Date'][i]
        index = df_master.index[df_master['ID'] == searched_ID]
        df_master.loc[index, 'Date'] = value_to_be_update

    # If ID does not exists, append new row to Master file
    else:
        df_master = pd.concat([df_master, df_update.iloc[[i]]], ignore_index=True)

# Set specific date format in the to-be-generated master file
df_master["Date"] = df_master["Date"].dt.strftime("%m/%d/%y")

# Set output(new master file) path and name
out_path = os.path.join(os.path.dirname(__file__), "Example/Updated Master {}.xls".format(datetime.today().date()))

# Generate and save new master file
writer = pd.ExcelWriter(out_path, engine='openpyxl', date_format='mm dd yyyy')
df_master.to_excel(writer, sheet_name='Created {}'.format(datetime.today().date()))
writer.save()
