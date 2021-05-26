# FICM-logbook-summary-from-LLP
Creates a FICM Logbook Summary from a Lifelong Learning Platform logbook export. 
Useful for Dual Anaesthetics and ICM Trainees in the UK who maintain an LLP Anaesthetics logbook but need to produce reports for ICM ARCPs. 

1. Export LLP logbook as excel file
3. Save the python script and logbook in same directory
4. Either rename the exported logbook as 'logbook_export.xlsx', or edit the 'logbook' user defined variable in the code to match the name of your file
5. Edit the 'name', 'start_date' and 'end_date' user defined variables
6. Running the script will then generate an excel file in the current directory with 'Events' and 'Procedures' tables on seperate sheets

## Webapp
Uses an older version of the script which doesn't incorporate date selection - so delete rows manually for unwanted date periods prior to running the script

https://hobnobs9444.github.io/logbook/

Developed by [@Stuk](https://github.com/Stuk) with pyodide
