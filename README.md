# EMEA ES Payroll Bulk Directive

Project Title: EMEA ES Payroll Bulk Directive

# Description
The end user must fill out an input form with data that must be transformed into a format recognized by South African Revenue Services and stored in a database or repository. The South African Revenue Services must get this data in bulk in order for them to review it and take appropriate action.

Proposed Solution:
Implement an integrated and user-friendly software system that facilitates seamless data entry and storage, automates the conversion process to meet SARS standards, and enables efficient management of the entire workflow. This solution not only enhances data accuracy and security but also diminishes reliance on external parties, thereby reducing associated costs for Kumba and Platinum Payroll. By automating this process, Anglo American was able to reduce time, improve accuracy, and control costs while submitting Tax Directive requests to the South African Revenue Services. This eliminated the need for a third party to do the activity on behalf of Kumba and Platinum Payroll.

# Prerequisites and Dependencies
Complete the excel template with correct and valid information/data.

# Code Explaination.
Data_Record_Generator Script:
This segment of code defines a Python class DataRecord with methods for processing different fields of a data record. Here's a breakdown of the main components:
1. Import Statements - Imports the pandas library as pd for data manipulation. Imports the datetime class from the datetime module for working with dates and times.
2. Class Definition (DataRecord) - The class has several methods, each responsible for processing a specific field from a row of data.
3. Method Definitions - Each method takes a row parameter, which is assumed to be a pandas DataFrame row. The methods return processed values for their respective fields, often applying formatting or transformations.

File_Generator Script:
This script is designed to process Excel files containing employee payroll data and generate corresponding text files for tax directives in South Africa. It supports two types of tax directives: IRP3S and IRP3A. Here's a high-level overview of the code:
1. Import necessary modules and suppress warnings.
2. Check if the script version matches the one used by the GSS Automation Team.
3. Ask the user to select an Excel file to process.
4. Read the header record from the selected Excel file.
5. Check the 'Information type' value to determine if it's IRP3S or IRP3A.
6. Define functions to generate header records, data records, and trailer records based on the information type.
7. Prepare logging to track the execution and store logs in a text file.
8. Process the data records sheet, calculating sums for certain fields, and generate data and trailer records.
9. Save the generated header, data, and trailer records in a text file.
10. Send an email to the automation team with the log file attached.
The code is organized into sections based on their functionality, such as logging preparation, reading data from Excel, and generating header, data, and trailer records.

The primary functions are:
* generate_header_record: Generates a header record for the text file based on the information type (IRP3S or IRP3A).
* generate_data_record: Generates a data record for each employee in the Excel file, which includes various personal and payroll-related details.
* generate_trailer_record: Generates a trailer record for the text file with calculated sums based on the information type (IRP3S or IRP3A).
* send_email_with_log: Sends an email to the automation team with the log file attached.
  
The code uses the pandas library to read and process Excel files, os and sys modules to interact with the operating system and environment variables, and logging to track execution and store the log in a text file. The win32com.client module is used to interact with Microsoft Outlook to send an email with the log file attached.

# Outcome Achieved:
1. Improved Efficiency: The implementation of this integrated automation has significantly improved the efficiency of the payroll data processing and submission process. By automating the conversion process and streamlining the workflow, the time required to submit Tax Directive requests to the South African Revenue Services has been reduced.
2. Enhanced Data Accuracy: This automation ensures that the data entered by the end user is accurate and meets the standards required by the South African Revenue Services. This has reduced the likelihood of errors in the submission process.
3. Cost Savings: By eliminating the need for a third party to handle the activity on behalf of Kumba and Platinum Payroll, the automation has helped reduce costs associated with outsourcing this task.
4. Improved Data Security: This automation provides a secure way to store and manage the payroll data, ensuring that sensitive information is protected.
5. User-Friendly Interface: This automation has a user-friendly interface that makes it easy for end users to fill out the input form and submit the data.
