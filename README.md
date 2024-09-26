SQL Attendance Data Processor
This Python script connects to a SQL Server database, retrieves employee attendance data, and processes it into formatted Excel sheets. The script includes a graphical user interface (GUI) for selecting a date range and specific employee IDs.

Features:
Database Connection: Connect to an SQL server using user-provided credentials.
Retrieve Data: Pull employee and attendance data based on user-selected parameters.
Excel Export: Export data into well-formatted Excel sheets with automatic column sizing, alternating row colors, and custom headers.
User Interface: Interactive GUI for selecting the date range and employee IDs for data extraction.

Requirements:
Python 3.x
Required Libraries:
pandas
pyodbc
openpyxl
tkinter
tkcalendar
Install the necessary libraries using the following command:

bash

pip install pandas pyodbc openpyxl tkcalendar
Usage Instructions:
Clone the Repository:

bash

git clone https://github.com/MianAshfaq/SQL-Attendance-Data-Processor.git
cd SQL-Attendance-Data-Processor
Update Connection Parameters:

Open the Python script (attendance_data_processor.py) and update the following details with your SQL Server configuration:

python

server = '<SQL_SERVER_IP>'
database = '<DATABASE_NAME>'
username = '<SQL_SERVER_USERNAME>'
password = '<SQL_SERVER_PASSWORD>'
Ensure you replace the placeholder text with your actual SQL Server IP, database name, username, and password.

Run the Script:

Execute the script using Python:

bash

python attendance_data.py
GUI Interaction:

After running the script, a graphical interface will pop up where you can:

Select the start and end dates for the attendance report.
Select specific employee IDs or enter them manually.
Output: The attendance data will be saved as an Excel file (attendance_data_v7.xlsx) in the working directory. Each employee will have a separate sheet in the workbook with formatted IN/OUT attendance records.

Customization Options:
SQL Query Customization: You can modify the SQL queries in the script to fetch different fields or add additional filters based on your requirements.
Excel Formatting: Customize the Excel output by editing the process_and_save_data function to change styles or formats as per your needs.
