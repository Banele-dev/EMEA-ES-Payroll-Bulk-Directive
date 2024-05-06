import pandas as pd
from datetime import datetime

class DataRecord:
    def __init__(self):
        pass

    def process_file_section_identifier(self, row):
        return row['File section identifier'] if pd.notnull(row['File section identifier']) else 'R'

    def process_request_seq_number(self, row):
        return str(row['Directive request ID number'] if pd.notnull(row['Directive request ID number']) else ' ').ljust(20)

    def process_employer_paye_reference_number(self, row):
        return row['Employer’s PAYE reference number']

    def process_employer_name(self, row):
        return str(row['Employer’s/Trade name']).ljust(120)

    def process_employer_physical_address(self, row):
        address = row['Employer’s physical business address']
        if pd.notnull(address):
            chunks = [address[i:i + 35] for i in range(0, len(address), 35)]
            formatted_address = ''.join(chunks)
            return formatted_address.ljust(35 * 4)[:35 * 4]
        else:
            return ' ' * (35 * 4)

    def process_employer_physical_post_code(self, row):
        return str(row['Employer’s physical business postal code']).ljust(10)[:10] if pd.notnull(row['Employer’s physical business postal code']) else ' ' * 10


    def process_employer_post_address(self, row):
        address = row['Employer’s postal address']
        if pd.notnull(address):
            chunks = [address[i:i + 35] for i in range(0, len(address), 35)]
            formatted_address = ''.join(chunks)
            return formatted_address.ljust(35 * 4)[:35 * 4]
        else:
            return ' ' * (35 * 4)

    def process_employer_post_code(self, row):
        return str(row['Employer’s postal code']).ljust(10)[:10] if pd.notnull(row['Employer’s postal code']) else ' ' * 10

    def process_employer_dial_code(self, row):
        return str(row['Employer’s dialling code']).replace("(", "").replace(")", "").ljust(10)[:10] if pd.notnull(row['Employer’s dialling code']) else ' ' * 10

    def process_employer_tel_no(self, row):
        return str(row['Employer’s telephone number']).replace(" ", "").ljust(10)[:10] if pd.notnull(row['Employer’s telephone number']) else ' ' * 10

    def process_employer_contact_person(self, row):
        return str(row['Contact person at the employer']).ljust(120)[:120] if pd.notnull(row['Contact person at the employer']) else ' ' * 120

    def process_email_address_administrator(self, row):
        return str(row['Employer email address']).ljust(50)[:50] if pd.notnull(row['Employer email address']) else ' ' * 50

    def process_income_tax_reference_number(self, row):
        return str(row['Income Tax reference number']).zfill(10) if pd.notnull(row['Income Tax reference number']) else ' ' * 10

    def process_no_it_ref_reason_text(self, row):
        return str(row['NO-IT-REF-REASON-TEXT']).ljust(65)[:65] if pd.notnull(row['NO-IT-REF-REASON-TEXT']) else ' ' * 65

    def process_taxpayer_sa_id_number(self, row):
        return row['Taxpayer SA ID number'] if pd.notnull(row['Taxpayer SA ID number']) else ' '

    def process_taxpayer_passport_no_permit_no(self, row):
        if pd.isnull(row['Taxpayer SA ID number']):
            return str(row['Taxpayer Passport No / Permit No']).ljust(18)[:18] if pd.notnull(row['Taxpayer Passport No / Permit No']) else ' ' * 18
        else:
            return ' ' * 18

    def process_passport_country(self, row):
        if pd.isnull(row['Taxpayer SA ID number']):
            return str(row['Taxpayer Passport Country / Country of Origin'])[:3] if pd.notnull(row['Taxpayer Passport Country / Country of Origin']) else ' ' * 3
        else:
            return ' ' * 3

    def process_taxpayer_employee_no(self, row):
        return str(row['Taxpayer employee number']).ljust(20)[:20] if pd.notnull(row['Taxpayer employee number']) else ' ' * 20

    def process_taxpayer_dob(self, row):
        return datetime.strptime(str(row['Taxpayer date of birth']), '%Y.%m.%d').strftime('%Y%m%d') if pd.notnull(row['Taxpayer date of birth']) else ' ' * 8

    def process_taxpayer_surname(self, row):
        return str(row['Taxpayer surname']).strip().ljust(120)[:120] if pd.notnull(row['Taxpayer surname']) else ' ' * 120

    def process_taxpayer_inits(self, row):
        return str(row['Taxpayer initials']).strip().ljust(5)[:5] if pd.notnull(row['Taxpayer initials']) else ' ' * 5

    def process_taxpayer_firstnames(self, row):
        return str(row['Taxpayer first names']).strip().ljust(90)[:90] if pd.notnull(row['Taxpayer first names']) else ' ' * 90

    def process_tp_res_address(self, row):
        address = row['TP-RES-ADDRESS']
        if pd.notnull(address):
            chunks = [address[i:i + 35] for i in range(0, len(address), 35)]
            formatted_address = ''.join(chunks)
            return formatted_address.ljust(35 * 4)[:35 * 4]
        else:
            return ' ' * (35 * 4)

    def process_tp_res_code(self, row):
        return str(row['Taxpayer residential postal code']).strip().ljust(10)[:10] if pd.notnull(row['Taxpayer residential postal code']) else ' ' * 10

    def process_tp_post_address(self, row):
        address = row['Taxpayer postal address']
        if pd.notnull(address):
            chunks = [address[i:i + 35] for i in range(0, len(address), 35)]
            formatted_address = ''.join(chunks)
            return formatted_address.ljust(35 * 4)[:35 * 4]
        else:
            return ' ' * (35 * 4)

    def process_tp_post_code(self, row):
        return str(row['Taxpayer postal code'])[:10].strip().ljust(10) if pd.notnull(row['Taxpayer postal code']) else ' ' * 10

    def process_tax_year(self, row):
        return str(row['Tax year for which the directive is requested']) if pd.notnull(row['Tax year for which the directive is requested']) and len(str(row['Tax year for which the directive is requested'])) == 4 and str(row['Tax year for which the directive is requested']).isdigit() else ' ' * 4

    def process_dir_reason(self, row):
        return str(row['Reason for directive'])[:2] if pd.notnull(row['Reason for directive']) else ' ' * 2

    def process_tp_annual_income(self, row):
        taxpayer_income = row['Taxpayer annual income'] if pd.notnull(row['Taxpayer annual income']) else ' '
        if pd.isnull(taxpayer_income):
            tp_annual_income = '0'.rjust(13, '0')
        else:
            try:
                # Round the amount to the nearest Rand and remove decimal part
                rounded_income = round(float(taxpayer_income))
                tp_annual_income = str(rounded_income).split('.')[0].rjust(13, '0')
            except ValueError:
                tp_annual_income = '0'.rjust(13, '0')
        return tp_annual_income


    def process_date_of_accrual(self, row):
        return datetime.strptime(str(row['Date when gross amount accrued']), '%Y.%m.%d').strftime('%Y%m%d') if pd.notnull(row['Date when gross amount accrued']) else ' ' * 8

    def process_empl_tax_resident_ind(self, row):
        return row['Is the employee a tax resident?'] if pd.notnull(row['Is the employee a tax resident?']) else ' '

    def process_S10_1_O_II_IND(self, row):
        return row['Is the exemption in terms of section 10(1)(o)(ii) applicable?'] if pd.notnull(row['Is the exemption in terms of section 10(1)(o)(ii) applicable?']) else ' '

    def process_SERV_REND_ABROAD_IND(self, row):
        return row['Were there any services rendered abroad?'] if pd.notnull(row['Were there any services rendered abroad?']) else ' '

    def process_SERV_REND_ABROAD_AMT(self, row):
        serv_rend_abroad_amt = row['If yes (to the above question), indicate the amount'] if pd.notnull(row['If yes (to the above question), indicate the amount']) else 0
        return str(int(serv_rend_abroad_amt * 100)).zfill(15)

    def process_TAX_WITHHELD_IND(self, row):
        return row['Was there any Tax withheld?'] if pd.notnull(row['Was there any Tax withheld?']) else ' '

    def process_TAX_WITHHELD_AMT(self, row):
        tax_withheld_amt = row['If yes (to the above question), indicate the amount2'] if pd.notnull(row['If yes (to the above question), indicate the amount2']) else 0
        return str(int(tax_withheld_amt * 100)).zfill(15)

    def process_start_date_qual_per(self, row):
        return datetime.strptime(str(row['START-DATE-QUAL-PER']), '%Y.%m.%d').strftime('%Y%m%d') + '0' * (8 * 14) if pd.notnull(row['START-DATE-QUAL-PER']) else '0' * (8 * 15)

    def process_end_date_qual_per(self, row):
        return datetime.strptime(str(row['END-DATE-QUAL-PER']), '%Y.%m.%d').strftime('%Y%m%d') + '0' * (8 * 14) if pd.notnull(row['END-DATE-QUAL-PER']) else '0' * (8 * 15)

    def process_tot_work_day_qual_per(self, row):
        return str(row['TOT-WORK-DAYS-QUAL-PER']).zfill(4)[-4:] + '0' * (4 * 14) if pd.notnull(row['TOT-WORK-DAYS-QUAL-PER']) else '0' * (4 * 15)

    def process_work_days_outside_sa_qual_per(self, row):
        return str(row['WORK-DAYS-OUTSIDE-SA-QUAL-PER']).zfill(4)[-4:] + '0' * (4 * 14) if pd.notnull(row['WORK-DAYS-OUTSIDE-SA-QUAL-PER']) else '0' * (4 * 15)

    def process_start_date_srce_per(self, row):
        return datetime.strptime(str(row['Start Date of the source period']), '%Y.%m.%d').strftime('%Y%m%d') if pd.notnull(row['Start Date of the source period']) else '0' * 8

    def process_end_date_srce_per(self, row):
        return datetime.strptime(str(row['End Date of the source period']), '%Y.%m.%d').strftime('%Y%m%d') if pd.notnull(row['End Date of the source period']) else '0' * 8

    def process_tot_work_days_srce_per(self, row):
        return str(row['TOT-WORK-DAYS-SRCE-PER']).zfill(4)[-4:] if pd.notnull(row['TOT-WORK-DAYS-SRCE-PER']) else '0' * 4

    def process_work_days_outside_sa_srce_per(self, row):
        return str(row['WORK-DAYS-OUTSIDE-SA-SRCE-PER']).zfill(4)[-4:] if pd.notnull(row['WORK-DAYS-OUTSIDE-SA-SRCE-PER']) else '0' * 4

    def process_yoa_year(self, row):
        return str(row['Year of Assessment in source period']) + '0' * (4 * 15) if pd.notnull(row['Year of Assessment in source period']) and len(str(row['Year of Assessment in source period'])) == 4 and str(row['Year of Assessment in source period']).isdigit() else '0' * (4 * 16)

    def process_yoa_tot_work_days(self, row):
        return str(row['Total workdays in source period']).zfill(3)[-3:] + '0' * (3 * 15) if pd.notnull(row['Total workdays in source period']) else '0' * (3 * 16)

    def process_yoa_work_days_outside_sa(self, row):
        return str(row['YOA-WORK-DAYS-OUTSIDE-SA']).zfill(3)[-3:] + '0' * (3 * 15) if pd.notnull(row['YOA-WORK-DAYS-OUTSIDE-SA']) else '0' * (3 * 16)

    # this function takes a row of data, retrieves the value from the column 'YOA-DEEMED-ACCRUAL', multiplies it by 100, converts it to an integer, converts it to a string, pads it with leading zeros to make it 15 characters long, and then returns the resulting string.
    def process_yoa_deemed_accrual(self, row):
        yoa_deemed_accrual = row['YOA-DEEMED-ACCRUAL'] if pd.notnull(row['YOA-DEEMED-ACCRUAL']) else 0
        return str(int(yoa_deemed_accrual * 100)).zfill(15) + '0' * (15 * 15)


    def process_yoa_used_prior_to_vesting(self, row):
        yoa_used_prior_to_vesting = row['YOA-USED-PRIOR-TO-VESTING'] if pd.notnull(row['YOA-USED-PRIOR-TO-VESTING']) else 0
        return str(int(yoa_used_prior_to_vesting * 100)).zfill(15) + '0' * (15 * 15)

    def process_yoa_portion_gain_qual_exempt(self, row):
        yoa_portion_gain_qual_exempt = row['YOA-PORTION-GAIN-QUAL-EXEMPT'] if pd.notnull(row['YOA-PORTION-GAIN-QUAL-EXEMPT']) else 0
        return str(int(yoa_portion_gain_qual_exempt * 100)).zfill(15) + '0' * (15 * 15)

    def process_gross_value(self, row):
        gross_value = row['Gross Value of gain'] if pd.notnull(row['Gross Value of gain']) else 0
        return str(int(gross_value * 100)).zfill(15)

    def process_S101OII_exempt_amount(self, row):
        s101OII_exempt_amount = row['S10-1-O-II-EXEMPT-AMOUNT'] if pd.notnull(row['S10-1-O-II-EXEMPT-AMOUNT']) else 0
        return str(int(s101OII_exempt_amount * 100)).zfill(15)

    def process_taxable_portion(self, row):
        taxable_portion = row['Taxable portion'] if pd.notnull(row['Taxable portion']) else 0
        return str(int(taxable_portion * 100)).zfill(15)

    def process_severance_benef_payable(self, row):
        severance_benef_payable = row['SEVERANCE-BENEF-PAYABLE'] if pd.notnull(row['SEVERANCE-BENEF-PAYABLE']) else 0
        return str(int(severance_benef_payable * 100)).zfill(15)

    def process_employ_owned_policy_amount(self, row):
        employ_owned_policy_amount = row['EMPLOY-OWNED-POLICY-AMOUNT'] if pd.notnull(row['EMPLOY-OWNED-POLICY-AMOUNT']) else 0
        return str(int(employ_owned_policy_amount * 100)).zfill(15)

    def process_SECT_10_1_GB_III_COMP(self, row):
        SECT_10_1_GB_III_COMP = row['SECT-10-1-GB-III-COMP'] if pd.notnull(row['SECT-10-1-GB-III-COMP']) else 0
        return str(int(SECT_10_1_GB_III_COMP * 100)).zfill(15)

    def process_leave_payment(self, row):
        leave_payment = row['LEAVE-PAYMENT'] if pd.notnull(row['LEAVE-PAYMENT']) else 0
        return str(int(leave_payment * 100)).zfill(15)

    def process_notice_payment(self, row):
        notice_payment = row['NOTICE-PAYMENT'] if pd.notnull(row['NOTICE-PAYMENT']) else 0
        return str(int(notice_payment * 100)).zfill(15)

    def process_arbitration_ccma_award(self, row):
        arbitration_ccma_award = row['ARBITRATION-CCMA-AWARD'] if pd.notnull(row['ARBITRATION-CCMA-AWARD']) else 0
        return str(int(arbitration_ccma_award * 100)).zfill(15)

    def process_other_amount_nature(self, row):
        other_amount_nature = row['OTHER-AMOUNT-NATURE']
        if pd.notnull(other_amount_nature):
            chunks = [other_amount_nature[i:i + 20] for i in range(0, len(other_amount_nature), 20)]
            formatted_nature = ''.join(chunks)
            return formatted_nature.ljust(20 * 7)[:20 * 7]
        else:
            return ' ' * (20 * 7)

    def process_other_amount(self, row):
        other_amount = row['OTHER-AMOUNT'] if pd.notnull(row['OTHER-AMOUNT']) else 0
        return str(int(other_amount * 100)).zfill(15) + '0' * (15 * 6)

    def process_gross_amount_payable(self, row):
        gross_amount_payable = row['GROSS-AMOUNT-PAYABLE'] if pd.notnull(row['GROSS-AMOUNT-PAYABLE']) else 0
        return str(int(gross_amount_payable * 100)).zfill(15)

    def process_directive_request_id_number(self, row):
        return str(row['Directive request ID number'] if pd.notnull(row['Directive request ID number']) else ' ').ljust(20)

    # def process_directive_id(self, row):
    #     return str(row['Directive ID (Original directive request)']).zfill(15)[:15] if pd.notnull(row['Directive ID (Original directive request)']) else '0' * 15

    def process_directive_id(self, row):
        return str(row['Directive ID (Original directive request)']).ljust(15)[:15] if pd.notnull(row['Directive ID (Original directive request)']) else ' ' * 15

    def process_fsca_registration_number(self, row):
        return str(row['FSCA registration number'] if pd.notnull(row['FSCA registration number ']) else ' ').ljust(19)

    def process_approved_fund_number(self, row):
        return str(row['Approved fund number']).ljust(11)[:11] if pd.notnull(row['Approved fund number']) else ' ' * 11

    def process_insurer_fsca_regisstered_number(self, row):
        return str(row['FSCA registered insurer number']).ljust(12)[:12] if pd.notnull(row['FSCA registered insurer number']) else ' ' * 12

    def process_directive_cancellation_reason(self, row):
        return str(row['Directive cancellation reason']).ljust(120)[:120] if pd.notnull(row['Directive cancellation reason']) else ' ' * 120

    def process_contact_person(self, row):
        return str(row['Contact Person']).ljust(120)[:120] if pd.notnull(row['Contact Person']) else ' ' * 120

    def process_dial_code_contact_person(self, row):
        return str(row['Contact Person dialing code']).replace("(", "").replace(")", "").ljust(10)[:10] if pd.notnull(row['Contact Person dialing code']) else ' ' * 10


    def process_tel_contact_person(self, row):
        return str(row['Contact Person telephone number']).replace(" ", "").ljust(10)[:10] if pd.notnull(row['Contact Person telephone number']) else ' ' * 10

    def process_declaration_ind(self, row):
        return row['DECLARATION-IND'] if pd.notnull(row['DECLARATION-IND']) else 'Y'

    def process_paper_resp(self, row):
        return row['Paper response indicator'] if pd.notnull(row['Paper response indicator']) else 'N'
