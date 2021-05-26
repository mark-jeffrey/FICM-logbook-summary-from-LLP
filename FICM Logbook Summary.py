# !/usr/bin/env python
# coding: utf-8

#  # FICM Logbook Summary Report

# Creates a FICM Logbook Summary Report
# from an Anaesthetics LLP Logbook excel export

# Import Libararies

import pandas as pd

# User Defined Variables

name = 'Mark Jeffrey'
logbook = 'my logbook.xlsx'
start_date = '2018-8-1'
end_date = '2021-2-7'

# Define functions


def get_event_local(str):  # counts no of 'str' under local supervision
    return len(icu_log_local[icu_log_local['Event'].str.contains(str)])


def get_event_distant(str):  # counts no of 'str' under distant supervision
    return len(icu_log_distant[icu_log_distant['Event'].str.contains(str)])


def get_event_total(str):  # counts total no of 'str'
    return len(icu_log[icu_log['Event'].str.contains(str)])


def get_procedures_total(str):  # counts total no of 'str'
    return len(procedures_all[procedures_all['Procedure Type'] == str])


def get_procedures_local(str):  # counts no of 'str' under local supervision
    return len(procedures_all_local[
               procedures_all_local['Procedure Type'] == str])


def get_procedures_distant(str):  # counts no of str under distant supervision
    return len(procedures_all_distant[
               procedures_all_distant['Procedure Type'] == str])


# Import each sheet into a DataFrame
# Set Case ID as Index
# Change 'Date' to Datetime
# Filter by dates

anaesthetic_log = pd.read_excel(logbook, sheet_name='LOGBOOK_CASE_ANAESTHETIC', index_col=0)
anaesthetic_log['Date'] = pd.to_datetime(anaesthetic_log['Date'], format="%d %B %Y")
anaesthetic_log = anaesthetic_log[(anaesthetic_log['Date'] > start_date) & (anaesthetic_log['Date'] < end_date)]

procedure_log = pd.read_excel(logbook, sheet_name='LOGBOOK_STAND_ALONE_PROCEDURE',index_col=0)
procedure_log['Date'] = pd.to_datetime(procedure_log['Date'], format="%d %B %Y")
procedure_log = procedure_log[(procedure_log['Date'] > start_date) & (procedure_log['Date'] < end_date)]

session_log = pd.read_excel(logbook, sheet_name='LOGBOOK_SESSION',index_col=0)
session_log['Date'] = pd.to_datetime(session_log['Date'], format="%d %B %Y")
session_log = session_log[(session_log['Date'] > start_date) & (session_log['Date'] < end_date)]

icu_log = pd.read_excel(logbook, sheet_name='LOGBOOK_CASE_INTENSIVE',index_col=0)
icu_log['Date'] = pd.to_datetime(icu_log['Date'], format="%d %B %Y")
icu_log = icu_log[(icu_log['Date'] > start_date) & (icu_log['Date'] < end_date)]


# Divide DataFrame by levels of supervision
# local = immediate and local
# distant = distant and solo
icu_log_local = icu_log[(icu_log['Supervision'] == 'Immediate') | (
                         icu_log['Supervision'] == 'Local')]
icu_log_distant = icu_log[(icu_log['Supervision'] == 'Distant') | (
                           icu_log['Supervision'] == 'Solo')]


# Data for events table
event_local = [
    get_event_local('ward-review'),
    get_event_local('admission'),
    get_event_local('lead-ward-round'),
    get_event_local('cardiac-arrest'),
    get_event_local('trauma-team'),
    get_event_local('intra-hospital-transfer'),
    get_event_local('inter-hospital-transfer'),
    get_event_local('discussion-with-relatives'),
    get_event_local('end-of-life-care')
]

event_distant = [
    get_event_distant('ward-review'),
    get_event_distant('admission'),
    get_event_distant('lead-ward-round'),
    get_event_distant('cardiac-arrest'),
    get_event_distant('trauma-team'),
    get_event_distant('intra-hospital-transfer'),
    get_event_distant('inter-hospital-transfer'),
    get_event_distant('discussion-with-relatives'),
    get_event_distant('end-of-life-care')
]

event_total = [
    get_event_total('ward-review'),
    get_event_total('admission'),
    get_event_total('lead-ward-round'),
    get_event_total('cardiac-arrest'),
    get_event_total('trauma-team'),
    get_event_total('intra-hospital-transfer'),
    get_event_total('inter-hospital-transfer'),
    get_event_total('discussion-with-relatives'),
    get_event_total('end-of-life-care')
]

#  Create Events table
Events = pd.DataFrame(
    data=[event_local, event_distant, event_total],
    index=['Local Supervision', 'Distant Supervision', 'Total'],
    columns=['Ward review', 'Admission', 'Lead ward round', 'Cardiac arrest',
             'Trauma team', 'Intra-hospital transfer',
             'Inter-hosptial transfer', 'Discussion with relatives',
             'End of life care/donation']
)

# Format events table
Events = Events.T
Events.index.names = ['Events']


# Procedures Table

# Procedures Multi-index
outside = [
    'Airways and Lungs', 'Airways and Lungs', 'Airways and Lungs',
    'Airways and Lungs', 'Airways and Lungs', 'Airways and Lungs',
    'Cardiovascular', 'Cardiovascular', 'Cardiovascular', 'Cardiovascular',
    'Cardiovascular', 'Cardiovascular', 'Cardiovascular',
    'Abdomen', 'Abdomen', 'Abdomen',
    'CNS', 'CNS'
]

inside = [
    'Emergency Intubation',
    'Percutaneous Tracheostomy',
    'Bronchoscopy',
    'Chest Drain - Seldinger',
    'Chest Drain - Surgical',
    'Lung Ultrasound',
    'Arterial cannulation',
    'Central venous access – IJ',
    'Central venous access – SC',
    'Central venous access – Femoral',
    'Pulmonary artery catheter',
    'Non-invasive CO monitoring',
    'Echocardiogram',
    'Ascitic drain/tap',
    'Sengstaken tube placement',
    'Abdominal ultrasound/FAST',
    'Lumbar puncture',
    'Brainstem death testing'
]

hier_index = list(zip(outside, inside))
hier_index = pd.MultiIndex.from_tuples(hier_index)


# Creates table for each procedure column with supervision
# Removes rows with missing entries
# Renames columns to "Procedure Type" and "Supervision"
anaesthetic_procedure_log = procedure_log[
    ['Procedure Type (Anaesthesia)', 'Supervision']].dropna().rename(
        columns={"Procedure Type (Anaesthesia)": "Procedure Type"})
medicine_procedure_log = procedure_log[
    ['Procedure Type (Medicine)', 'Supervision']].dropna().rename(
        columns={"Procedure Type (Medicine)": "Procedure Type"})
pain_procedure_log = procedure_log[
    ['Procedure Type (Pain)', 'Supervision']].dropna().rename(
        columns={"Procedure Type (Pain)": "Procedure Type"})
anaesthetic_sheet_procedure_log = anaesthetic_log[
    ['Procedure Type', 'Procedure Supervision']].dropna().rename(
        columns={"Procedure Supervision": "Supervision"})


# Contatencates above into a single table with 2 columns
# "Procedure Type" and "Supervision"
procedures_all = pd.concat([
    anaesthetic_procedure_log,
    medicine_procedure_log,
    pain_procedure_log,
    anaesthetic_sheet_procedure_log]
)

# Makes Supervision column all lower case to avoid duplication
procedures_all['Supervision'] = procedures_all['Supervision'].str.lower()

# Subdivides all procedures by level of supervision
procedures_all_local = procedures_all[
    (procedures_all['Supervision'] == 'supervised') | (
     procedures_all['Supervision'] == 'observed')]
procedures_all_distant = procedures_all[
     procedures_all['Supervision'] == 'solo']


#  Data for Procedures table
procedures_total = [
    len(procedures_all[(
        procedures_all['Procedure Type'] == 'rsi') | (
        procedures_all['Procedure Type'] == 'emergency-intubation') | (
        procedures_all['Procedure Type'] == 'airway-protection')]),
    get_procedures_total('percutaneous-tracheostomy'),
    get_procedures_total('bronchoscopy'),
    get_procedures_total('intercostal-drain:seldinger'),
    get_procedures_total('intercostal-drain:open'),
    get_procedures_total('lung-ultrasound'),
    get_procedures_total('arterial-cannulation'),
    get_procedures_total('central-venous-access–internal-jugular'),
    get_procedures_total('central-venous-access–subclavian'),
    get_procedures_total('central-venous-access–femoral'),
    get_procedures_total('pulmonary-artery-catheter'),
    get_procedures_total('non-invasive-co-monitoring'),
    get_procedures_total('echocardiogram'),
    len(procedures_all[(
        procedures_all['Procedure Type'] == 'ascitic-tap') | (
        procedures_all['Procedure Type'] == 'abdominal-paracentesis')]),
    get_procedures_total('sengstacken-tube-placement'),
    get_procedures_total('abdominal-ultrasound/fast'),
    get_procedures_total('lumbar-puncture'),
    get_procedures_total('brainstem-death-testing')
]

procedures_total_local = [
    len(procedures_all_local[(
        procedures_all_local['Procedure Type'] == 'rsi') | (
        procedures_all_local['Procedure Type'] == 'emergency-intubation') | (
        procedures_all_local['Procedure Type'] == 'airway-protection')]),
    get_procedures_local('percutaneous-tracheostomy'),
    get_procedures_local('bronchoscopy'),
    get_procedures_local('intercostal-drain:seldinger'),
    get_procedures_local('intercostal-drain:open'),
    get_procedures_local('lung-ultrasound'),
    get_procedures_local('arterial-cannulation'),
    get_procedures_local('central-venous-access–internal-jugular'),
    get_procedures_local('central-venous-access–subclavian'),
    get_procedures_local('central-venous-access–femoral'),
    get_procedures_local('pulmonary-artery-catheter'),
    get_procedures_local('non-invasive-co-monitoring'),
    get_procedures_local('echocardiogram'),
    len(procedures_all_local[(
        procedures_all_local['Procedure Type'] == 'ascitic-tap') | (
        procedures_all_local['Procedure Type'] == 'abdominal-paracentesis')]),
    get_procedures_local('sengstacken-tube-placement'),
    get_procedures_local('abdominal-ultrasound/fast'),
    get_procedures_local('lumbar-puncture'),
    get_procedures_local('brainstem-death-testing')
    ]

procedures_total_distant = [
    len(procedures_all_distant[(
        procedures_all_distant['Procedure Type'] == 'rsi') | (
        procedures_all_distant['Procedure Type'] == 'emergency-intubation') | (
        procedures_all_distant['Procedure Type'] == 'airway-protection')]),
    get_procedures_distant('percutaneous-tracheostomy'),
    get_procedures_distant('bronchoscopy'),
    get_procedures_distant('intercostal-drain:seldinger'),
    get_procedures_distant('intercostal-drain:open'),
    get_procedures_distant('lung-ultrasound'),
    get_procedures_distant('arterial-cannulation'),
    get_procedures_distant('central-venous-access–internal-jugular'),
    get_procedures_distant('central-venous-access–subclavian'),
    get_procedures_distant('central-venous-access–femoral'),
    get_procedures_distant('pulmonary-artery-catheter'),
    get_procedures_distant('non-invasive-co-monitoring'),
    get_procedures_distant('echocardiogram'),
    len(procedures_all_distant[(
        procedures_all_distant['Procedure Type'] == 'ascitic-tap') | (
        procedures_all_distant['Procedure Type'] == 'abdominal-paracentesis')]),
    get_procedures_distant('sengstacken-tube-placement'),
    get_procedures_distant('abdominal-ultrasound/fast'),
    get_procedures_distant('lumbar-puncture'),
    get_procedures_distant('brainstem-death-testing')
    ]


# Procedures Table
Procedures = pd.DataFrame(
    data=[procedures_total_local, procedures_total_distant, procedures_total],
    index=['Local Supervision', 'Distant Supervision', 'Total'],
    columns=hier_index
)

# Format procedures table
Procedures = Procedures.T
Procedures.index.names = (['System', 'Procedure'])


# Export to Excel
with pd.ExcelWriter(f"{name} FICM Logbook Summary {start_date} to {end_date}.xlsx") as writer:
    Events.to_excel(writer, sheet_name='Events')
    Procedures.to_excel(writer, sheet_name='Procedures')
