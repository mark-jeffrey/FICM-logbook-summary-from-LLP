# !/usr/bin/env python
# coding: utf-8

#  # FICM Logbook Summary Report

# Creates a FICM Logbook Summary Report from an Anaesthetics LLP Logbook excel export

# Import Libararies

import pandas as pd

# User Defined Variables

name = 'ICM Trainee'
logbook = 'logbook_export.xlsx'
start_date = 'start'
end_date = 'end'


# ## Events Table

# Import each sheet into a DataFrame
# Set Case ID as Index

anaesthetic_log = pd.read_excel(
    logbook, sheet_name='LOGBOOK_CASE_ANAESTHETIC', index_col=0)
procedure_log = pd.read_excel(
    logbook, sheet_name='LOGBOOK_STAND_ALONE_PROCEDURE', index_col=0)
session_log = pd.read_excel(
    logbook, sheet_name='LOGBOOK_SESSION', index_col=0)
icu_log = pd.read_excel(
    logbook, sheet_name='LOGBOOK_CASE_INTENSIVE', index_col=0)

# Divide DataFrame by levels of supervision
# local = immediate and local
# distant = distant and solo
# teaching not included in ICU cases - ?need to include in Notes on LLP

icu_log_local = icu_log[(icu_log['Supervision'] == 'Immediate') | (icu_log['Supervision'] == 'Local')]
icu_log_distant = icu_log[(icu_log['Supervision'] == 'Distant') | (icu_log['Supervision'] == 'Solo')]
icu_log_teaching = icu_log[icu_log['Supervision'] == 'Teaching']


# Data for events table
# More abstraction by re-using num function?
def num(str):
    return len(icu_log_local[icu_log_local['Event'].str.contains(str)])


event_local = [
    num('ward-review'),
    num('admission'),
    num('lead-ward-round'),
    num('cardiac-arrest'),
    num('trauma-team'),
    num('intra-hospital-transfer'),
    num('inter-hospital-transfer'),
    num('discussion-with-relatives'),
    num('end-of-life-care')
]


def num(str):
    return len(icu_log_distant[icu_log_distant['Event'].str.contains(str)])


event_distant = [
    num('ward-review'),
    num('admission'),
    num('lead-ward-round'),
    num('cardiac-arrest'),
    num('trauma-team'),
    num('intra-hospital-transfer'),
    num('inter-hospital-transfer'),
    num('discussion-with-relatives'),
    num('end-of-life-care')
]


def num(str):
    return len(icu_log[icu_log['Event'].str.contains(str)])


event_total = [
    num('ward-review'),
    num('admission'),
    num('lead-ward-round'),
    num('cardiac-arrest'),
    num('trauma-team'),
    num('intra-hospital-transfer'),
    num('inter-hospital-transfer'),
    num('discussion-with-relatives'),
    num('end-of-life-care')
]


#  Creates Events table
Events = pd.DataFrame(
    data=[event_local, event_distant, event_total],
    index=['Local Supervision', 'Distant Supervision', 'Total'],
    columns=['Ward review', 'Admission', 'Lead ward round', 'Cardiac arrest', 'Trauma team', 'Intra-hospital transfer', 'Inter-hosptial transfer', 'Discussion with relatives', 'End of life care/donation']
)

# Transpose axes for correct layout
Events = Events.T
Events.index.names = ['Events']


#  ## Procedures Table

# Procedures Multi-index

outside = [
    'Airways and Lungs', 'Airways and Lungs', 'Airways and Lungs', 'Airways and Lungs', 'Airways and Lungs', 'Airways and Lungs',
    'Cardiovascular', 'Cardiovascular', 'Cardiovascular', 'Cardiovascular', 'Cardiovascular', 'Cardiovascular', 'Cardiovascular',
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

anaesthetic_procedure_log = procedure_log[['Procedure Type (Anaesthesia)', 'Supervision']].dropna().rename(columns={"Procedure Type (Anaesthesia)": "Procedure Type"})
medicine_procedure_log = procedure_log[['Procedure Type (Medicine)', 'Supervision']].dropna().rename(columns={"Procedure Type (Medicine)": "Procedure Type"})
pain_procedure_log = procedure_log[['Procedure Type (Pain)', 'Supervision']].dropna().rename(columns={"Procedure Type (Pain)": "Procedure Type"})
anaesthetic_sheet_procedure_log = anaesthetic_log[['Procedure Type', 'Procedure Supervision']].dropna().rename(columns={"Procedure Supervision": "Supervision"})

# Contatencates above into a single table with 2 columns "Procedure Type" and "Supervision"

procedures_all = pd.concat([
    anaesthetic_procedure_log,
    medicine_procedure_log,
    pain_procedure_log,
    anaesthetic_sheet_procedure_log]
)

# Makes Supervision column all lower case to avoid duplication

procedures_all['Supervision'] = procedures_all['Supervision'].str.lower()

# Subdivides all procedures by level of supervision

procedures_all_local = procedures_all[(procedures_all['Supervision'] == 'supervised') | (procedures_all['Supervision'] == 'observed')]
procedures_all_distant = procedures_all[procedures_all['Supervision'] == 'solo']


#  Data for Procedures table
def num(str):
    return len(procedures_all[procedures_all['Procedure Type'] == str])


procedures_total = [
    len(procedures_all[(procedures_all['Procedure Type'] == 'rsi') | (procedures_all['Procedure Type'] == 'emergency-intubation') | (procedures_all['Procedure Type'] == 'airway-protection')]),
    num('percutaneous-tracheostomy'),
    num('bronchoscopy'),
    num('intercostal-drain:seldinger'),
    num('intercostal-drain:open'),
    num('lung-ultrasound'),
    num('arterial-cannulation'),
    num('central-venous-access–internal-jugular'),
    num('central-venous-access–subclavian'),
    num('central-venous-access–femoral'),
    num('pulmonary-artery-catheter'),
    num('non-invasive-co-monitoring'),
    num('echocardiogram'),
    len(procedures_all[(procedures_all['Procedure Type'] == 'ascitic-tap') | (procedures_all['Procedure Type'] == 'abdominal-paracentesis')]),
    num('sengstacken-tube-placement'),
    num('abdominal-ultrasound/fast'),
    num('lumbar-puncture'),
    num('brainstem-death-testing')
    ]


def num(str):
    return len(procedures_all_local[procedures_all_local['Procedure Type'] == str])


procedures_total_local = [
    len(procedures_all_local[(procedures_all_local['Procedure Type'] == 'rsi') | (procedures_all_local['Procedure Type'] == 'emergency-intubation') | (procedures_all_local['Procedure Type'] == 'airway-protection')]),
    num('percutaneous-tracheostomy'),
    num('bronchoscopy'),
    num('intercostal-drain:seldinger'),
    num('intercostal-drain:open'),
    num('lung-ultrasound'),
    num('arterial-cannulation'),
    num('central-venous-access–internal-jugular'),
    num('central-venous-access–subclavian'),
    num('central-venous-access–femoral'),
    num('pulmonary-artery-catheter'),
    num('non-invasive-co-monitoring'),
    num('echocardiogram'),
    len(procedures_all_local[(procedures_all_local['Procedure Type'] == 'ascitic-tap') | (procedures_all_local['Procedure Type'] == 'abdominal-paracentesis')]),
    num('sengstacken-tube-placement'),
    num('abdominal-ultrasound/fast'),
    num('lumbar-puncture'),
    num('brainstem-death-testing')
    ]


def num(str):
    return len(procedures_all_distant[procedures_all_distant['Procedure Type'] == str])


procedures_total_distant = [
    len(procedures_all_distant[(procedures_all_distant['Procedure Type'] == 'rsi') | (procedures_all_distant['Procedure Type'] == 'emergency-intubation') | (procedures_all_distant['Procedure Type'] == 'airway-protection')]),
    num('percutaneous-tracheostomy'),
    num('bronchoscopy'),
    num('intercostal-drain:seldinger'),
    num('intercostal-drain:open'),
    num('lung-ultrasound'),
    num('arterial-cannulation'),
    num('central-venous-access–internal-jugular'),
    num('central-venous-access–subclavian'),
    num('central-venous-access–femoral'),
    num('pulmonary-artery-catheter'),
    num('non-invasive-co-monitoring'),
    num('echocardiogram'),
    len(procedures_all_distant[(procedures_all_distant['Procedure Type'] == 'ascitic-tap') | (procedures_all_distant['Procedure Type'] == 'abdominal-paracentesis')]),
    num('sengstacken-tube-placement'),
    num('abdominal-ultrasound/fast'),
    num('lumbar-puncture'),
    num('brainstem-death-testing')
    ]


# Procedures Table

Procedures = pd.DataFrame(
    data=[procedures_total_local, procedures_total_distant, procedures_total],
    index=['Local Supervision', 'Distant Supervision', 'Total'],
    columns=hier_index
)

# Transpose table
Procedures = Procedures.T
Procedures.index.names = (['System', 'Procedure'])


# ## Export to Excel

with pd.ExcelWriter(f"{name} FICM Logbook Summary {start_date} to {end_date}.xlsx") as writer:
    Events.to_excel(writer, sheet_name='Events')
    Procedures.to_excel(writer, sheet_name='Procedures')
