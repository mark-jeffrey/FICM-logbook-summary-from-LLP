# logbook

Creates a FICM Logbook Summary from a Lifelong Learning Platform logbook export. Useful for Dual Anaesthetics and ICM Trainees in the UK who maintain an LLP Anaesthetics logbook but need to produce reports for ICM ARCPs.


## logbook_improved.py
Use this if executing as python script (more accurate than the current version on the webapp - I will update this eventually). 

1. Export LLP logbook as excel file
2. Save the python script and logbook in same directory
3. Change filename varible to match the excel sheet filename
4. Change start_date and end_date for date filtered report
5. Run the script


## Webapp

Uses an logbook.py which doesn't incorporate date selection - delete rows manually for unwanted date periods prior to running the script

https://hobnobs9444.github.io/logbook/

Developed by @Stuk with pyodide
