import logging
from datetime import datetime
import time
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import sys
import win32com.client as win32

user_name = os.getlogin()
# Setting variables to check is this version matches with the GSS Automation Team's control
application = "ES_Payroll Bulk Directive"
version = "v04"
path = f"C:/Users/{user_name}/Box/Automation Script Versions/versions.xlsx"
df = pd.read_excel(path)
filter_criteria = (df['app'] == application) & (df['versão'] == version)
start_time = None

if not filter_criteria.any():
    input('Outdated app, talk to the automation team. Press ENTER to close the code \n')
    quit()

root = tk.Tk()
root.withdraw()  # Hide the main window
# Ask the user to select a text file
text_file_path = filedialog.askopenfilename(title="Select Text file", filetypes=[("Text files", "*.txt")])
root.destroy()  # Destroy the Tkinter window

if not text_file_path:
    print("No file selected. Exiting...")
    sys.exit()

################################ LOG PREPARATION ##################################

# Get the path of the directory where the script is located
script_directory = os.path.dirname(os.path.abspath(__file__))

# Create the path for the LogControl folder
log_control_path = os.path.join(script_directory, 'LogControl')
# If the LogControl folder doesn't exist, create it
if not os.path.exists(log_control_path):
    os.makedirs(log_control_path)

# Create the full path to the log file within the LogControl folder
log_file_name = f"ExecutionLog_{datetime.now().strftime('%d%m%Y%H%M')}"+".txt"
log_file_path = os.path.join(log_control_path, log_file_name)

# configures the logging module to log messages to a file.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path)
    ]
)

try:
    # Read the text file into a list of lines
    with open(text_file_path, "r") as file:
        lines = file.readlines()

    # Extract header, data, and trailer records
    header_record = lines[0].strip()
    data_records = [line.strip() for line in lines[1:-1]]
    trailer_record = lines[-1].strip()
    logging.info('The response text file has been read and processed successfully')
    status_automation = "Successfully"
except Exception as e:
    logging.error(f"An error occurred when reading and processing the file: Here is the error {e}")
    status_automation = "Failed"

#  extracts specific fields from the header_record and stores them in a list
header_fields = []
try:
    header_fields = [
        header_record[0:1].strip(),
        header_record[1:9].strip(),
        header_record[9:17].strip(),
        header_record[17:18].strip(),
        header_record[18:19].strip(),
        header_record[19:27].strip(),
        header_record[27:35].strip(),
        header_record[35:49].strip(),
        header_record[49:63].strip(),
        header_record[63:77].strip(),
        header_record[77:79].strip(),
        header_record[79:80].strip()
    ]
    logging.info("The header fields has been extracted successfully")
    status_automation = "Successfully"
except Exception as e:
    logging.error(f"An error occurred while extracting header fields: Here is the error {e}")
    status_automation = "Failed"

# iterates over each data_record in the data_records list and extracts specific fields from each data_record. It then appends the extracted fields as a list to the data_fields list
data_fields = []
try:
    for data_record in data_records:
        data_fields.append([
            data_record[0:1].strip(),  # sec_id
            data_record[1:21].strip(),  # req_seq_num
            data_record[21:36].strip(),  # appl_id
            data_record[36:40].strip(),  # it_area
            data_record[40:50].strip(),  # it_ref_no
            data_record[50:65].strip(),  # dir_id
            data_record[65:67].strip(),  # req_status
            data_record[67:75].strip(),  # issue_date
            data_record[75:83].strip(),  # validity_start
            data_record[83:91].strip(),  # validity_end
            data_record[91:106].strip(),  # gross_amount
            data_record[106:114].strip(),  # accrual_date
            data_record[114:118].strip(),  # source_code
            data_record[118:133].strip(),  # serv_rend_loc_amt
            data_record[133:137].strip(),  # serv_rend_loc_source_code
            data_record[137:152].strip(),  # serv_rend_for_amt
            data_record[152:156].strip(),  # serv_rend_for_source_code
            data_record[156:171].strip(),  # tax_withheld
            data_record[171:175].strip(),  # year_of_assessment
            data_record[175:190].strip(),  # pre_1march1998_amt
            data_record[190:205].strip(),  # trf_amt
            data_record[205:220].strip(),  # mem_fund_contr
            data_record[220:235].strip(),  # exc_contrib_amt
            data_record[235:250].strip(),  # taxed_transf_non_memb_spouse
            data_record[250:265].strip(),  # exempt_services
            data_record[265:280].strip(),  # aipf_member_contributions
            data_record[280:295].strip(),  # S10_1_O_ii_exempt_amount
            data_record[295:310].strip(),  # deemed_prov_fund_contrib
            data_record[310:325].strip(),  # total_benefit
            data_record[325:327].strip(),  # tax_rate
            data_record[327:342].strip(),  # tax_free_portion
            data_record[342:357].strip(),  # gross_amount_paye
            data_record[357:358].strip(),  # deduction_frequency
            data_record[358:373].strip(),  # allowed_contributions
            data_record[373:386].strip(),  # approved_deemed_remuneration
            data_record[386:401].strip(),  # it88l_ref_no
            data_record[401:416].strip(),  # assessed_tax_amount
            data_record[416:435].strip(),  # assessed_tax_prn
            data_record[435:450].strip(),  # admin_penalty
            data_record[450:469].strip(),  # admin_penalty_prn
            data_record[469:514].strip(),  # provisional_tax_amount
            data_record[514:532].strip(),  # provisional_tax_period
            data_record[532:589].strip(),  # provisional_tax_prn
        ])
        logging.info("The data fields has been extracted successfully")
        logging.info(f"Number of data record to be generated: {len(data_records)}")
        total_data_records = len(data_records)
        status_automation = "Successfully"
except Exception as e:
    logging.error(f"An error occurred while extracting data fields: Here is the error {e}")
    status_automation = "Failed"

# extracts specific fields from the trailer_record and stores them in a list
trailer_fields = []
try:
    trailer_fields = [
        trailer_record[0:1].strip(),
        trailer_record[1:9].strip(),
        trailer_record[9:29].strip(),
        trailer_record[29:49].strip(),
        trailer_record[49:69].strip(),
        trailer_record[69:89].strip(),
        trailer_record[89:109].strip(),
        trailer_record[109:125].strip(),
        trailer_record[125:145].strip(),
        trailer_record[145:165].strip(),
        trailer_record[165:185].strip(),
        trailer_record[185:205].strip(),
        trailer_record[205:225].strip(),
        trailer_record[225:245].strip(),
        trailer_record[245:265].strip(),
        trailer_record[265:285].strip(),
        trailer_record[285:305].strip(),
        trailer_record[305:325].strip(),
    ]
    logging.info("The tailer fields has been extracted successfully")
    status_automation = "Successfully"
except Exception as e:
    logging.error(f"An error occurred while extracting trailer fields: Here is the error {e}")
    status_automation = "Failed"

try:
    # Create a DataFrame for header, data, and trailer records using the three lists
    header_df = pd.DataFrame([header_fields], columns=["File section identifier", "Information type", "Information sub-type", "Test data indicator", "File series control field", "External system identification", "Interface version number", "Unique file identifier", "Date and time of file creation", "Unique file identifier of the file from which the response was generated", "Source file processing status", "Tax Directive Request Type"])
    data_df = pd.DataFrame(data_fields, columns=["File section identifier", "Directive request ID number", "Directive application ID", "The Income Tax area to which this taxpayer belongs", "Income Tax reference number", "Directive ID (Original directive request)", "Request status", "Date of directive issue", "The start date of the validity period of this directive", "The end date of the validity period of this directive", "Gross amount of lump sum", "Date of accrual of lump sum", "The lump sum source code", "Services rendered local amount", "Services rendered local source code", "Services rendered abroad (foreign) amount", "Services rendered foreign source code", "Tax Withheld", "The assessment of tax year to which this tax directive applies", "Vested right pre-2 March 1998", "Amount Transferred", "Own contribution to a provident fund (up to 1 March 2016)", "Contributions not previously allowed as a deduction", "Transferred divorce benefit previously taxed", "Amount_exempt_based_on_services_outside_the_Republic", "AIPF member transfer contributions", "Amount exempt in terms of section 10(1)(o)(ii)", "Deemed provident fund contributions (After tax pension benefit)", "Full benefit used to purchase an annuity", "Tax free portion of the gross lump sum gratuity/remuneration", "Tax free portion of the gross lump sum gratuity/remuneration", "PAYE amount to be deducted from gross remuneration", "Frequency of deducting PAYE amount from gross lump sum gratuity/remuneration", "Contributions allowed as exemption from lump sum", "Approved monthly deemed remuneration", "IT 88L reference number", "Tax amount to be deducted for outstanding Assessed tax", "Assessed Tax Payment Reference Number", "Administrative Penalty", "Administrative Penalty Payment Reference Number", "Provisional Tax amount to be deducted for outstanding Provisional Tax", "Period for which Provisional tax is outstanding", "Provisional Tax Payment Reference Number"])
    trailer_df = pd.DataFrame([trailer_fields], columns=["File section identifier", "Number of records in this file", "Gross amount of lump sum", "PAYE amount to be deducted from gross remuneration", "Tax free portion of the gross lump sum gratuity/remuneration", "Tax amount to be deducted for outstanding Assessed tax", "Provisional Tax amount to be deducted for outstanding Provisional Tax", "Tax free portion of the gross lump sum gratuity/remuneration", "Vested right pre-2 March 1998", "Amount Transferred", "Own contribution to a provident fund (up to 1 March 2016)", "Contributions not previously allowed as a deduction", "Transferred divorce benefit previously taxed", "Amount_exempt_based_on_services_outside_the_Republic", "AIPF member transfer contributions", "Amount exempt in terms of section 10(1)(o)(ii)", "Administrative Penalty Payment Reference Number", "Full benefit used to purchase an annuity"])
    logging.info("DataFrames has been created successfully")
    status_automation = "Successfully"
except Exception as e:
    logging.error(f"An error occurred while creating DataFrames: Here is the error {e}")
    status_automation = "Failed"

# Creates an Excel file and writes three DataFrames to it as separate sheets.
response_excel_file = f"C:\\Users\\{user_name}\\PycharmProjects\\EMEA_ES_Payroll%20Bulk%20Directive\\EMEA_ES_Payroll%20Bulk%20Directive Response\\response_excel_file.xlsx"
try:
    with pd.ExcelWriter(response_excel_file) as writer:
        header_df.to_excel(writer, sheet_name="Header Record", index=False)
        data_df.to_excel(writer, sheet_name="Data Record", index=False)
        trailer_df.to_excel(writer, sheet_name="Trailer Record", index=False)
        logging.info("An excel file has been created successfully")
        status_automation = "Successfully"
except Exception as e:
    logging.error(f"An error occurred while writing to Excel: {e}")
    status_automation = "Failed"

# sys.stdout = sys.__stdout__
start_time = time.time()
end_time = time.time()
execution_duration = round(end_time - start_time, 2)
# Create an Outlook application object and Create a new email
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.Subject = 'Automation Team - Automation Log'
email_body = "EMEA ES Payroll Bulk Directive" + "_" + str(datetime.today()) + "_" + str(status_automation) + "_" + str(execution_duration) + "_" + str(total_data_records) + "_" + "number of Data Record generated"

email.HTMLBody = email_body
email_recipients = ['banele.madikane@angloamerican.com', 'breno.andrade@angloamerican.com']
email.To = '; '.join(email_recipients)

# Attach the log file
attachment = os.path.abspath(log_file_path)
email.Attachments.Add(attachment)
email.Send()