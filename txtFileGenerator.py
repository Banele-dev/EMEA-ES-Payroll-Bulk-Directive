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
application = "ES_Payroll Bulk Directive"
version = "v01"
user_name = os.getlogin()
path = f"C:/Users/{user_name}/Box/Automation Script Versions/versions.xlsx"
df = pd.read_excel(path)
filter_criteria = (df['app'] == application) & (df['versão'] == version)
start_time = None

if not filter_criteria.any():
    input('Outdated app, talk to the automation team. Press ENTER to close the code \n')
    quit()

# Initialize Tkinter
root = Tk()
root.withdraw()  # Hide the main window
# Ask the user to select an Excel file
excel_file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls")])
root.destroy()  # Destroy the Tkinter window

if not excel_file_path:
    logging.info("No file selected. Exiting...")
    sys.exit()

start_time = time.time()

# Read Header Record
header_data = pd.read_excel(excel_file_path, sheet_name='Header Record')

# Check the external system identification value
info_type = header_data.iloc[0]['Information type']

if info_type == 'IRP3S':
    ################################ HEADER RECORD FUNCTION ##################################
    # Function to generate the header record with all entries left-justified and blank-padded
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
            tax_calc_ind = row['Tax Directive Request Type']

            # Combine all the parts into one header string
            header_record = f"{sec_id}{info_type}{info_subtype}{test_data}{file_series_ctl}{ext_sys}{ver_no}{own_file_id}{gen_time}{tax_calc_ind}"
            logging.info('Header Record generated successfully.')
            status_automation = "Successful"
            return header_record
        except Exception as e:
            logging.error(f"Error during Header Record generation.Here it follows the error: {e}")
            status_automation = "Failed"


    ################################ DATA RECORD FUNCTION ##################################
    def generate_data_record(row):
        try:
            # Class created to generate each variable of the data record
            data_record = DataRecord()

            # Generating Data Fields
            sec_id = data_record.process_file_section_identifier(row)
            reg_seq_num = data_record.process_request_seq_number(row)
            paye_ref_no = data_record.process_employer_paye_reference_number(row)
            emp_name = data_record.process_employer_name(row)
            emp_physical_address = data_record.process_employer_physical_address(row)
            emp_physical_post_code = data_record.process_employer_physical_post_code(row)
            emp_post_address = data_record.process_employer_post_address(row)
            emp_post_code = data_record.process_employer_post_code(row)
            emp_dial_code = data_record.process_employer_dial_code(row)
            emp_tel_no = data_record.process_employer_tel_no(row)
            emp_contact_person = data_record.process_employer_contact_person(row)
            email_address_administrator = data_record.process_email_address_administrator(row)
            it_ref_no = data_record.process_income_tax_reference_number(row)
            no_it_ref_reason_text = data_record.process_no_it_ref_reason_text(row)
            tp_id = data_record.process_taxpayer_sa_id_number(row)
            tp_other_id = data_record.process_taxpayer_passport_no_permit_no(row)
            passport_country = data_record.process_passport_country(row)
            tp_employee_no = data_record.process_taxpayer_employee_no(row)
            tp_dob = data_record.process_taxpayer_dob(row)
            tp_surname = data_record.process_taxpayer_surname(row)
            tp_inits = data_record.process_taxpayer_inits(row)
            tp_firstnames = data_record.process_taxpayer_firstnames(row)
            tp_res_address = data_record.process_tp_res_address(row)
            tp_res_code = data_record.process_tp_res_code(row)
            tp_post_address = data_record.process_tp_post_address(row)
            tp_post_code = data_record.process_tp_post_code(row)
            tax_year = data_record.process_tax_year(row)
            dir_reason = data_record.process_dir_reason(row)
            tp_annual_income = data_record.process_tp_annual_income(row)
            date_of_accrual = data_record.process_date_of_accrual(row)
            empl_tax_resident_ind = data_record.process_empl_tax_resident_ind(row)
            S10_1_O_II_IND = data_record.process_S10_1_O_II_IND(row)
            SERV_REND_ABROAD_IND = data_record.process_SERV_REND_ABROAD_IND(row)
            SERV_REND_ABROAD_AMT = data_record.process_SERV_REND_ABROAD_AMT(row)
            TAX_WITHHELD_IND = data_record.process_TAX_WITHHELD_IND(row)
            TAX_WITHHELD_AMT = data_record.process_TAX_WITHHELD_AMT(row)
            start_date_qual_per = data_record.process_start_date_qual_per(row)
            end_date_qual_per = data_record.process_end_date_qual_per(row)
            tot_work_day_qual_per = data_record.process_tot_work_day_qual_per(row)
            work_days_outside_sa_qual_per = data_record.process_work_days_outside_sa_qual_per(row)
            start_date_srce_per = data_record.process_start_date_srce_per(row)
            end_date_srce_per = data_record.process_end_date_srce_per(row)
            tot_work_days_srce_per = data_record.process_tot_work_days_srce_per(row)
            work_days_outside_sa_srce_per = data_record.process_work_days_outside_sa_srce_per(row)
            yoa_year = data_record.process_yoa_year(row)
            yoa_tot_work_days = data_record.process_yoa_tot_work_days(row)
            yoa_work_days_outside_sa = data_record.process_yoa_work_days_outside_sa(row)
            yoa_deemed_accrual = data_record.process_yoa_deemed_accrual(row)
            yoa_used_prior_to_vesting = data_record.process_yoa_used_prior_to_vesting(row)
            yoa_portion_gain_qual_exempt = data_record.process_yoa_portion_gain_qual_exempt(row)
            gross_value = data_record.process_gross_value(row)
            S101OII_exempt_amount = data_record.process_S101OII_exempt_amount(row)
            taxable_portion = data_record.process_taxable_portion(row)
            declaration_ind = data_record.process_declaration_ind(row)
            paper_resp = data_record.process_paper_resp(row)

            # Combine all the parts into one header string
            data_record = f"{sec_id}{reg_seq_num}{paye_ref_no}{emp_name}{emp_physical_address}{emp_physical_post_code}{emp_post_address}{emp_post_code}{emp_dial_code}{emp_tel_no}{emp_contact_person}{email_address_administrator}{it_ref_no}{no_it_ref_reason_text}{tp_id}{tp_other_id}{passport_country}{tp_employee_no}{tp_dob}{tp_surname}{tp_inits}{tp_firstnames}{tp_res_address}{tp_res_code}{tp_post_address}{tp_post_code}{tax_year}{dir_reason}{tp_annual_income}{date_of_accrual}{empl_tax_resident_ind}{S10_1_O_II_IND}{SERV_REND_ABROAD_IND}{SERV_REND_ABROAD_AMT}{TAX_WITHHELD_IND}{TAX_WITHHELD_AMT}{start_date_qual_per}{end_date_qual_per}{tot_work_day_qual_per}{work_days_outside_sa_qual_per}{start_date_srce_per}{end_date_srce_per}{tot_work_days_srce_per}{work_days_outside_sa_srce_per}{yoa_year}{yoa_tot_work_days}{yoa_work_days_outside_sa}{yoa_deemed_accrual}{yoa_used_prior_to_vesting}{yoa_portion_gain_qual_exempt}{gross_value}{S101OII_exempt_amount}{taxable_portion}{declaration_ind}{paper_resp}"
            logging.info(f"Line number {str(row['Line_Num'])} from Data Record generated successfully.")
            status_automation = "Successfully"
            return data_record
        except Exception as e:
            logging.error(f"Error in line number {str(row['Line_Num'])} during Data Record generation. Here it follows the error: {e}")
            status_automation = "Failed"


    ################################ TRAILER RECORD FUNCTION ##################################

    def generate_trailer_record(rec_nb, annual_income_sum, gross_value_sum, exempt_amount_sum, taxable_portion_sum):
        try:
            # Trailer fields
            sec_id = 'T'
            rec_no = str(rec_nb).rjust(8, '0')
            annual_income = '{:0>16}'.format(int(annual_income_sum * 100))
            # first convert the gross_value_sum to cents by multiplying it by 100 and converting it to an integer. Then, we use string formatting to create a string with a width of 20 characters, right-justified, and filled with zeros. The {:0>20} format specification means to pad the string with zeros on the left until it reaches a width of 20 characters.
            gross_value = '{:0>20}'.format(int(gross_value_sum * 100))
            exempt_amount = '{:0>20}'.format(int(exempt_amount_sum * 100))
            taxable_portion = '{:0>20}'.format(int(taxable_portion_sum * 100))


            # Combine all the parts into one trailer string
            trailer_record = f"{sec_id}{rec_no}{annual_income}{gross_value}{exempt_amount}{taxable_portion}"
            logging.info('Trailer Record generated successfully.')
            status_automation = "Successfully"
            return trailer_record
        except Exception as e:
            logging.error(f"Error during Trailer Record generation.Here it follows the error: {e}")
            status_automation = "Failed"

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

        # Calculating: Annual Income Sum - # Gross Value Sum - # Exemption Amount Sum - # Taxable Portion Sum
        annual_income_sum = data_data['Taxpayer annual income'].sum()
        gross_value_sum = data_data['Gross Value of gain'].sum()
        exempt_amount_sum = data_data['S10-1-O-II-EXEMPT-AMOUNT'].sum()
        taxable_portion_sum = data_data['Taxable portion'].sum()
        logging.info('Database "ES Payroll Bulk Tax Directive Template.xlsx" opened successfully.')

        # Iterate through rows in each sheet and generate records
        header_records = [generate_header_record(row) for index, row in header_data.iterrows()]
        logging.info(f"Number of data record to be generated: {len(data_data)}")
        generated_records = len(data_data)
        data_records = [generate_data_record(row) for index, row in data_data.iterrows()]
        trailer_records = [generate_trailer_record(rec_nb, annual_income_sum, gross_value_sum, exempt_amount_sum, taxable_portion_sum)]

        # Read Header Record
        header_data = pd.read_excel(excel_file_path, sheet_name='Header Record')

        # Check the external system identification value
        ext_sys_id = header_data.iloc[0]['External system identification']
        # file_path = f"C:\\Users\\{user_name}\\Desktop"
        file_path = f"C:/Users/{user_name}/PycharmProjects/EMEA_ES_Payroll%20Bulk%20Directive/EMEA_ES_Payroll%20Bulk%20Directive"

        # Define the counter file based on the external system identification
        if ext_sys_id == 'KUMBAIRO':
            counter_file = os.path.join(file_path, "kumba_counter_sharepayments.txt")
            if not os.path.exists(counter_file):
                with open(counter_file, "w") as file:
                    file.write("0")
        elif ext_sys_id == 'RUSTPLAT':
            counter_file = os.path.join(file_path, "platinum_counter_sharepayments.txt")
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

        file_name = os.path.join(file_path, f"IRP3S.R{new_number:06}.txt")


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

elif info_type == 'IRP3A':
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
            tax_calc_ind = row['Tax Directive Request Type']

            # Combine all the parts into one header string
            header_record = f"{sec_id}{info_type}{info_subtype}{test_data}{file_series_ctl}{ext_sys}{ver_no}{own_file_id}{gen_time}{tax_calc_ind}"
            logging.info('Header Record generated successfully.')
            status_automation = "Successfully"
            return header_record
        except Exception as e:
            logging.error(f"Error during Header Record generation.Here it follows the error: {e}")
            status_automation = "Failed"


    ################################ DATA RECORD FUNCTION ##################################
    def generate_data_record(row):

        try:
            #Class created to generate each varible of the data record
            data_record = DataRecord()

            # Generating Data Fields
            sec_id = data_record.process_file_section_identifier(row)
            reg_seq_num = data_record.process_request_seq_number(row)
            paye_ref_no = data_record.process_employer_paye_reference_number(row)
            emp_name = data_record.process_employer_name(row)
            emp_physical_address = data_record.process_employer_physical_address(row)
            emp_physical_post_code = data_record.process_employer_physical_post_code(row)
            emp_post_address = data_record.process_employer_post_address(row)
            emp_post_code = data_record.process_employer_post_code(row)
            emp_dial_code = data_record.process_employer_dial_code(row)
            emp_tel_no = data_record.process_employer_tel_no(row)
            emp_contact_person = data_record.process_employer_contact_person(row)
            email_address_administrator = data_record.process_email_address_administrator(row)
            it_ref_no = data_record.process_income_tax_reference_number(row)
            no_it_ref_reason_text = data_record.process_no_it_ref_reason_text(row)
            tp_id = data_record.process_taxpayer_sa_id_number(row)
            tp_other_id = data_record.process_taxpayer_passport_no_permit_no(row)
            passport_country = data_record.process_passport_country(row)
            tp_employee_no = data_record.process_taxpayer_employee_no(row)
            tp_dob = data_record.process_taxpayer_dob(row)
            tp_surname = data_record.process_taxpayer_surname(row)
            tp_inits = data_record.process_taxpayer_inits(row)
            tp_firstnames = data_record.process_taxpayer_firstnames(row)
            tp_res_address = data_record.process_tp_res_address(row)
            tp_res_code = data_record.process_tp_res_code(row)
            tp_post_address = data_record.process_tp_post_address(row)
            tp_post_code = data_record.process_tp_post_code(row)
            tax_year = data_record.process_tax_year(row)
            dir_reason = data_record.process_dir_reason(row)
            tp_annual_income = data_record.process_tp_annual_income(row)
            date_of_accrual = data_record.process_date_of_accrual(row)
            severance_benef_payable = data_record.process_severance_benef_payable(row)
            employ_owned_policy_amount = data_record.process_employ_owned_policy_amount(row)
            SECT_10_1_GB_III_COMP = data_record.process_SECT_10_1_GB_III_COMP(row)
            leave_payment = data_record.process_leave_payment(row)
            notice_payment = data_record.process_notice_payment(row)
            arbitration_ccma_award = data_record.process_arbitration_ccma_award(row)
            other_amount_nature = data_record.process_other_amount_nature(row)
            other_amount = data_record.process_other_amount(row)
            gross_amount_payable = data_record.process_gross_amount_payable(row)
            declaration_ind = data_record.process_declaration_ind(row)
            paper_resp = data_record.process_paper_resp(row)

            # Combine all the parts into one header string
            data_record = f"{sec_id}{reg_seq_num}{paye_ref_no}{emp_name}{emp_physical_address}{emp_physical_post_code}{emp_post_address}{emp_post_code}{emp_dial_code}{emp_tel_no}{emp_contact_person}{email_address_administrator}{it_ref_no}{no_it_ref_reason_text}{tp_id}{tp_other_id}{passport_country}{tp_employee_no}{tp_dob}{tp_surname}{tp_inits}{tp_firstnames}{tp_res_address}{tp_res_code}{tp_post_address}{tp_post_code}{tax_year}{dir_reason}{tp_annual_income}{date_of_accrual}{severance_benef_payable}{employ_owned_policy_amount}{SECT_10_1_GB_III_COMP}{leave_payment}{notice_payment}{arbitration_ccma_award}{other_amount_nature}{other_amount}{gross_amount_payable}{declaration_ind}{paper_resp}"
            logging.info(f"Line number {str(row['Line_Num'])} from Data Record generated successfully.")
            status_automation = "Successfully"
            return data_record
        except Exception as e:
            logging.error(f"Error in line number {str(row['Line_Num'])} during Data Record generation. Here it follows the error: {e}")
            status_automation = "Failed"


    ################################ TRAILER RECORD FUNCTION ##################################

    def generate_trailer_record(rec_nb, annual_income_sum, severance_benef_pay_sum, employ_owned_policy_amt_sum, SECT_10_1_GB_III_COMP_SUM, leave_payment_sum, notice_payment_sum, arbitration_ccma_award_sum, other_amount_sum, gross_amount_payable_sum):
        try:
            # Trailer fields
            sec_id = 'T'
            rec_no = str(rec_nb).rjust(8, '0')
            annual_income = '{:0>16}'.format(int(annual_income_sum))
            # annual_income = '{:0>16}'.format(int(round(annual_income_sum)))
            # first convert the gross_value_sum to cents by multiplying it by 100 and converting it to an integer. Then, we use string formatting to create a string with a width of 20 characters, right-justified, and filled with zeros. The {:0>20} format specification means to pad the string with zeros on the left until it reaches a width of 20 characters.
            severance_benef_pay = '{:0>20}'.format(int(severance_benef_pay_sum * 100))
            employ_owned_policy_amt = '{:0>20}'.format(int(employ_owned_policy_amt_sum * 100))
            SECT_10_1_GB_III_COMP = '{:0>20}'.format(int(SECT_10_1_GB_III_COMP_SUM * 100))
            leave_payment = '{:0>20}'.format(int(leave_payment_sum * 100))
            notice_payment = '{:0>20}'.format(int(notice_payment_sum * 100))
            arbitration_ccma_award = '{:0>20}'.format(int(arbitration_ccma_award_sum * 100))
            other_amount = '{:0>20}'.format(int(other_amount_sum * 100))
            gross_amount_payable = '{:0>20}'.format(int(gross_amount_payable_sum * 100))

            # Combine all the parts into one trailer string
            trailer_record = f"{sec_id}{rec_no}{annual_income}{severance_benef_pay}{employ_owned_policy_amt}{SECT_10_1_GB_III_COMP}{leave_payment}{notice_payment}{arbitration_ccma_award}{other_amount}{gross_amount_payable}"
            logging.info('Trailer Record generated successfully.')
            status_automation = "Successfully"
            return trailer_record
        except Exception as e:
            logging.error(f"Error during Trailer Record generation.Here it follows the error: {e}")
            status_automation = "Failed"

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

        # Calculating: Annual Income Sum - # Gross Value Sum - # Exemption Amount Sum - # Taxable Portion Sum
        annual_income_sum = data_data['Taxpayer annual income'].sum()
        severance_benef_pay_sum = data_data['SEVERANCE-BENEF-PAYABLE'].sum()
        employ_owned_policy_amt_sum = data_data['EMPLOY-OWNED-POLICY-AMOUNT'].sum()
        SECT_10_1_GB_III_COMP_SUM = data_data['SECT-10-1-GB-III-COMP'].sum()
        leave_payment_sum = data_data['LEAVE-PAYMENT'].sum()
        notice_payment_sum = data_data['NOTICE-PAYMENT'].sum()
        arbitration_ccma_award_sum = data_data['ARBITRATION-CCMA-AWARD'].sum()
        other_amount_sum = data_data['OTHER-AMOUNT'].sum()
        gross_amount_payable_sum = data_data['GROSS-AMOUNT-PAYABLE'].sum()
        logging.info('Database "ES Payroll Bulk Tax Directive Template.xlsx" opened successfully.')

        # Iterate through rows in each sheet and generate records
        header_records = [generate_header_record(row) for index, row in header_data.iterrows()]
        logging.info(f"Number of data record to be generated: {len(data_data)}")
        generated_records = len(data_data)
        data_records = [generate_data_record(row) for index, row in data_data.iterrows()]
        trailer_records = [generate_trailer_record(rec_nb, annual_income_sum, severance_benef_pay_sum, employ_owned_policy_amt_sum, SECT_10_1_GB_III_COMP_SUM, leave_payment_sum, notice_payment_sum, arbitration_ccma_award_sum, other_amount_sum, gross_amount_payable_sum)]

        # Read Header Record
        header_data = pd.read_excel(excel_file_path, sheet_name='Header Record')

        # Check the external system identification value
        ext_sys_id = header_data.iloc[0]['External system identification']
        # file_path = f"C:\\Users\\{user_name}\\Desktop"
        file_path = f"C:/Users/{user_name}/PycharmProjects/EMEA_ES_Payroll%20Bulk%20Directive/EMEA_ES_Payroll%20Bulk%20Directive"

        # Define the counter file based on the external system identification
        if ext_sys_id == 'KUMBAIRO':
            counter_file = os.path.join(file_path, "kumba_counter_retrenchment.txt")
            if not os.path.exists(counter_file):
                with open(counter_file, "w") as file:
                    file.write("0")
        elif ext_sys_id == 'RUSTPLAT':
            counter_file = os.path.join(file_path, "platinum_counter_retrenchment.txt")
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

        file_name = os.path.join(file_path, f"IRP3A.R{new_number:06}.txt")


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

else:
    logging.info(f"Unsupported information type: {info_type}. Exiting...")
    status_automation = "Failed"
    sys.exit(1)

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

