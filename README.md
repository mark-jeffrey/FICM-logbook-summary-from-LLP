# FICM-logbook-summary-from-LLP
Creates a FICM Logbook Summary from a Lifelong Learning Platform logbook export. 
Useful for Dual Anaesthetics and ICM Trainees in the UK who maintain an LLP Anaesthetics logbook but need to produce reports for ICM ARCPs. 

1. Export LLP logbook as excel file
2. If you want a report of a specific time period, currently you will need to manually delete unwanted dates from the excel file prior to running the script
3. Save the python script and logbook in same directory
4. Either rename the exported logbook as 'logbook_export.xlsx', or edit the 'logbook' user defined variable in the code to match the name of your file
5. Edit the 'name', 'start_date' and 'end_date' user defined variables - note these currently only effect the name of the output save file
6. Running the script will then generate an excel file in the current directory with 'Events' and 'Procedures' tables on seperate sheets

## Webapp
Developed by @Stuk with pyodide
https://stuk.github.io/pyder/
https://github.com/Stuk/pyder
