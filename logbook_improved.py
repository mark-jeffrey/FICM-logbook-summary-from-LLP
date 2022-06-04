# ICU procedure logbook summariser
'''Takes Lifelong Learning Platform generated logbook excel sheet 
and summarises ICU relevant procedures'''

# Import libraries
import datetime as dt
import pandas as pd

# Variables
filename = 'logbook_complete.xlsx'
start_date = '2021-08-04'
end_date = '2022-06-04'


# Data imports
# Import date, procedure and supervision level from LLP output
print('Importing logbook...')
anaes_log = pd.read_excel(
    filename,
    sheet_name='LOGBOOK_CASE_ANAESTHETIC',
    usecols='B,Y:Z',
    dtype={
        'Procedure Type': 'string',
        'Procedure Supervision': 'string'}
)

proc_log = pd.read_excel(
    filename,
    sheet_name='LOGBOOK_STAND_ALONE_PROCEDURE',
    usecols='B,E:H',
    dtype={
        'Procedure Type (Anaesthesia)': 'string',
        'Procedure Type (Medicine)': 'string',
        'Procedure Type (Pain)': 'string',
        'Supervision': 'string'}
)

icu_log = pd.read_excel(
    filename,
    sheet_name='LOGBOOK_CASE_INTENSIVE',
    usecols='B,J:K',
    dtype={
        'Event': 'string',
        'Supervision': 'string'}
)

# Process logs prior to concat
print('Processing...')
# anaes_log = anaes_log.dropna()
anaes_log.rename(
    columns={'Procedure Supervision': 'Supervision'}, inplace=True)

# Combine procedure type columns from proc_log into one
proc_log['Procedure Type'] = proc_log['Procedure Type (Medicine)'].fillna(
    proc_log['Procedure Type (Pain)']).fillna(
        proc_log['Procedure Type (Anaesthesia)']
)

# Drop old columns
proc_log = proc_log.drop([
    'Procedure Type (Anaesthesia)',
    'Procedure Type (Medicine)',
    'Procedure Type (Pain)'],
    axis=1
)

# Separate ICU events
icu_log[['Event1', 'Event2', 'Event3']
        ] = icu_log['Event'].str.split(',', expand=True)

icu_log1 = icu_log[['Date', 'Event1', 'Supervision']].dropna()
icu_log2 = icu_log[['Date', 'Event2', 'Supervision']].dropna()
icu_log3 = icu_log[['Date', 'Event3', 'Supervision']].dropna()

icu_log1.rename(columns={'Event1': 'Procedure Type'}, inplace=True)
icu_log2.rename(columns={'Event2': 'Procedure Type'}, inplace=True)
icu_log3.rename(columns={'Event3': 'Procedure Type'}, inplace=True)

# Join and format ICU logs
icu_log_final = pd.concat([icu_log1, icu_log2, icu_log3])
icu_log_final = icu_log_final.apply(lambda x: x.astype(str).str.strip())
icu_log_final = icu_log_final[icu_log_final['Procedure Type'].isin(
    ['intra-hospital-transfer', 'inter-hospital-transfer'])]

# Join logs
log = pd.concat([anaes_log, proc_log, icu_log_final])


# Cleaning
print('Cleaning...')
# lower case and strip whitespace
log = log.apply(lambda x: x.astype(str).str.lower())
log = log.apply(lambda x: x.astype(str).str.strip())

# remove nan and rename duplicate entries
#log.drop(log[log['Procedure Type'] == '<na>'].index, inplace=True)
log.loc[log['Procedure Type'] == 'transfer-inter-hospital',
        'Procedure Type'] = 'inter-hospital-transfer'
log.loc[log['Procedure Type'] == 'transfer-intra-hospital',
        'Procedure Type'] = 'intra-hospital-transfer'
log.loc[log['Procedure Type'] == 'pa-catheter',
        'Procedure Type'] = 'pulmonary-artery-catheter'
log.loc[log['Procedure Type'] == 'rsi',
        'Procedure Type'] = 'rapid-sequence-induction'
log.loc[log['Procedure Type'] == 'airway-protection',
        'Procedure Type'] = 'rapid-sequence-induction'


# Date filtered report
print('Filtering for selected dates...')
date_fileterd_log = log.copy()
date_fileterd_log['Date'] = pd.to_datetime(
    date_fileterd_log['Date'], format="%d %B %Y")
date_fileterd_log.set_index('Date', inplace=True)
date_fileterd_log = date_fileterd_log.loc[start_date:end_date]
date_fileterd_log.reset_index(inplace=True)


# Complete report
print('Creating reports...')
report = pd.pivot_table(
    data=log,
    index='Procedure Type',
    columns='Supervision',
    aggfunc='count',
    margins=True,
    fill_value=0
)

# Date filtered report
filtered_report = pd.pivot_table(
    data=date_fileterd_log,
    index='Procedure Type',
    columns='Supervision',
    aggfunc='count',
    margins=True,
    fill_value=0
)

# Write to Excel files
print('Writing reports to excel...')
with pd.ExcelWriter(f"ICU procedure summary.xlsx") as writer:
    report.to_excel(writer, sheet_name='Events')

with pd.ExcelWriter(f"ICU procedure summary past year only.xlsx") as writer:
    filtered_report.to_excel(writer, sheet_name='Events')

print('Summary completed')
