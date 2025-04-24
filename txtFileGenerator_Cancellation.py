import pandas as pd
from datetime import datetime
import sys
import os
from data_record_generator import DataRecord
import warnings
import win32com.client as win32
from tkinter import filedialog
from tkinter import Tk
import logging
import time

# Suppress warnings from the openpyxl library, this warning does not affect the functionality of the code.
warnings.filterwarnings("ignore", category=UserWarning)

## Setting variables to check is this version matches with the GSS Automation Team's control
# application = "ES_Payroll Bulk Directive"
# version = "v04"
# user_name = os.getlogin()
# path = f"C:/Users/{user_name}/Box/Automation Script Versions/versions.xlsx"
# df = pd.read_excel(path)
# filter_criteria = (df['app'] == application) & (df['versão'] == version)
# start_time = None
#
# if not filter_criteria.any():
#     input('Outdated app, talk to the automation team. Press ENTER to close the code \n')
#     quit()

# Initialize Tkinter
root = Tk()
root.withdraw()  # Hide the main window
# Ask the user to select an Excel file
excel_file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls")])
root.destroy()  # Destroy the Tkinter window

if not excel_file_path:
    logging.info("No file selected. Exiting...")
    sys.exit()

user_name = os.getlogin()
start_time = time.time()

    ################################ HEADER RECORD FUNCTION ##################################
def generate_header_record(row):
    try:
        # Extract required data from the Excel row
        sec_id = 'H'.ljust(1)
        info_type = str(row['Information type']).ljust(8)
        info_subtype = ' '.ljust(8)
        test_data = row['Test data indicator']
        file_series_ctl = 'S'
        ext_sys = str(row['External system identification']).ljust(8)
        ver_no = str(row['Interface version number']).rjust(8)
        own_file_id = (f"AA{datetime.now().strftime('%d%m%Y%H%M')}").ljust(14) # Uses format: 'AA'<DDMMYYYY><HHMM> for unique identifier
        gen_time = datetime.now().strftime('%Y%m%d%H%M%S')
        # tax_calc_ind = row['Tax Directive Request Type']

        # Combine all the parts into one header string
        header_record = f"{sec_id}{info_type}{info_subtype}{test_data}{file_series_ctl}{ext_sys}{ver_no}{own_file_id}{gen_time}"
        logging.info('Header Record generated successfully.')
        return header_record
    except Exception as e:
        logging.error(f"Error during Header Record generation.Here it follows the error: {e}")


################################ DATA RECORD FUNCTION ##################################
def generate_data_record(row):
    try:
        # Class created to generate each variable of the data record
        data_record = DataRecord()

        # Generating Data Fields
        sec_id = data_record.process_file_section_identifier(row)
        reg_seq_num = data_record.process_directive_request_id_number(row)
        dir_id = data_record.process_directive_id(row)
        it_ref_no = data_record.process_income_tax_reference_number(row)
        tp_id = data_record.process_taxpayer_sa_id_number(row)
        tp_other_id = data_record.process_taxpayer_passport_no_permit_no(row)
        fsca_regis_no = data_record.process_fsca_registration_number(row)
        fund_number = data_record.process_approved_fund_number(row)
        insurer_fsca_regis_no = data_record.process_insurer_fsca_regisstered_number(row)
        cancel_reason = data_record.process_directive_cancellation_reason(row)
        contact_person = data_record.process_contact_person(row)
        dial_code_contact_person = data_record.process_dial_code_contact_person(row)
        tel_contact_person = data_record.process_tel_contact_person(row)

        # Combine all the parts into one header string
        data_record = f"{sec_id}{reg_seq_num}{dir_id}{it_ref_no}{tp_id}{tp_other_id}{fsca_regis_no}{fund_number}{insurer_fsca_regis_no}{cancel_reason}{contact_person}{dial_code_contact_person}{tel_contact_person}"
        logging.info(f"Line number {str(row['Line_Num'])} from Data Record generated successfully.")
        return data_record
    except Exception as e:
        logging.error(f"Error in line number {str(row['Line_Num'])} during Data Record generation. Here it follows the error: {e}")


################################ TRAILER RECORD FUNCTION ##################################

def generate_trailer_record(rec_nb):
    try:
        # Trailer fields
        sec_id = 'T'
        rec_no = str(rec_nb).rjust(8, '0')


        # Combine all the parts into one trailer string
        trailer_record = f"{sec_id}{rec_no}"
        logging.info('Trailer Record generated successfully.')
        return trailer_record
    except Exception as e:
        logging.error(f"Error during Trailer Record generation.Here it follows the error: {e}")

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

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ]
)


try:
    ################################ READ DATA FROM MS EXCEL and WRITE TO TXT ##################################
    # Read Header Record
    header_data = pd.read_excel(excel_file_path, sheet_name='Header Record')

    # Read Data Record
    data_data = pd.read_excel(excel_file_path, sheet_name='Data Record')
    data_data['Line_Num'] = range(1, len(data_data) + 1)

    # Get total number of records in the data record section of the file.
    rec_nb = len(data_data)
    logging.info('ES Payroll Bulk Tax Directive Template - Cancellation.xlsx" opened successfully.')

    # Iterate through rows in each sheet and generate records
    header_records = [generate_header_record(row) for index, row in header_data.iterrows()]
    logging.info(f"Number of data record to be generated: {len(data_data)}")
    generated_records = len(data_data)
    data_records = [generate_data_record(row) for index, row in data_data.iterrows()]
    trailer_records = [generate_trailer_record(rec_nb)]

    # Read Header Record
    header_data = pd.read_excel(excel_file_path, sheet_name='Header Record')

    # Check the external system identification value
    ext_sys_id = header_data.iloc[0]['External system identification']
    file_path = f"C:\\Users\\{user_name}\\Documents\\EMEA_ES_Payroll%20Bulk%20Directive\\Cancellation of Processed Directives"
    # file_path = f"C:\\Users\\Public\\Documents\\EMEA_ES_Payroll%20Bulk%20Directive\\EMEA_ES_Payroll%20Bulk%20Directive\\Cancellation of Processed Directives"

    # Define the counter file based on the external system identification
    if ext_sys_id == 'KUMBAIRO':
        counter_file = os.path.join(file_path, "kumba_counter_cancellations.txt")
        if not os.path.exists(counter_file):
            with open(counter_file, "w") as file:
                file.write("0")
    elif ext_sys_id == 'RUSTPLAT':
        counter_file = os.path.join(file_path, "platinum_counter_cancellations.txt")
        if not os.path.exists(counter_file):
            with open(counter_file, "w") as file:
                file.write("0")
    elif ext_sys_id == 'MODIKWA1':
        counter_file = os.path.join(file_path, "modikwa_counter_retrenchment.txt")
        if not os.path.exists(counter_file):
            with open(counter_file, "w") as file:
                file.write("0")
    elif ext_sys_id == 'ANGLOAME':
        counter_file = os.path.join(file_path, "GSS_counter_retrenchment.txt")
        if not os.path.exists(counter_file):
            with open(counter_file, "w") as file:
                file.write("0")
    elif ext_sys_id == 'ANGLOCOR':
        counter_file = os.path.join(file_path, "ACSSA_counter_retrenchment.txt")
        if not os.path.exists(counter_file):
            with open(counter_file, "w") as file:
                file.write("0")
    else:
        logging.info(f"Unsupported external system identification: {ext_sys_id}. Exiting...")
        sys.exit(1)

    # Read the last used number from the file
    with open(counter_file, "r") as file:
        last_used_number = int(file.read())

    # Increment the number by 1
    new_number = last_used_number + 1

    # Write the new number back to the file
    with open(counter_file, "w") as file:
        file.write(str(new_number))

    # Generate the new file name

    file_name = os.path.join(file_path, f"IRP3CRQ.R{new_number:06}.txt")


    # Write the header record, trailer record to the specified file
    with open(file_name, 'w') as file:
        file.write('\n'.join(header_records))
        file.write('\n')
        file.write('\n'.join(data_records))
        file.write('\n')
        file.write('\n'.join(trailer_records))

    # Inform the user that the header record, data record, trailer record has been saved
    logging.info(f"Records saved to {file_name}")
    status_automation = "Successfully"

except Exception as e:
    logging.info(f"Error opening the database, please check the data content again. Here it follows the error: {e}")
    status_automation = "Failed"


sys.stdout = sys.__stdout__
end_time = time.time()
execution_duration = round(end_time - start_time, 2)

# Create an Outlook application object and Create a new email
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.Subject = 'Automation Team - Automation Log'
email_body = "EMEA ES Payroll Bulk Directive" + "_" + str(datetime.today()) + "_" + str(status_automation) + "_" + str(execution_duration) + "_" + str(generated_records) + "_" + "number of Data Record generated"

email.HTMLBody = email_body
email_recipients = ['banele.madikane@angloamerican.com']
email.To = '; '.join(email_recipients)

# Attach the log file
attachment = os.path.abspath(log_file_path)
email.Attachments.Add(attachment)
email.Send()

