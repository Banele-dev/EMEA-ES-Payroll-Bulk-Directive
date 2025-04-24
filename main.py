import subprocess
import warnings
import os
from datetime import datetime
import logging
import time
import sys


# Suppress warnings from the openpyxl library, this warning does not affect the functionality of the code.
warnings.filterwarnings("ignore", category=UserWarning)

## Setting variables to check is this version matches with the GSS Automation Team's control
# application = "ES_Payroll Bulk Directive"
# version = "v04"
# user_name = os.getlogin()
# path = f"C:/Users/{user_name}/Box/Automation Script Versions/versions.xlsx"
# df = pd.read_excel(path)
# filter_criteria = (df['app'] == application) & (df['versão'] == version)
# # start_time = None
#
# if not filter_criteria.any():
#     input('Outdated app, talk to the automation team. Press ENTER to close the code \n')
#     quit()

user_name = os.getlogin()
start_time = time.time()

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

def run_FileGenerator():
    try:
        script_directory = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_directory, "txtFileGenerator.py")
        subprocess.run(["python", script_path])
        logging.info("txtFileGenerator is running successfully")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running txtFileGenerator.py: {e}")

def run_FileGenerator_Cancellation():
    try:
        script_directory = os.path.dirname(os.path.abspath(__file__))
        subprocess.run(["python", "txtFileGenerator_Cancellation.py"], cwd=script_directory)
        logging.info("txtFileGenerator_Cancellation.py is running successfully")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running txtFileGenerator_Cancellation.py: {e}")

def run_ResponseFileReader():
    try:
        script_directory = os.path.dirname(os.path.abspath(__file__))
        subprocess.run(["python", "txtResponseFileReader.py"], cwd=script_directory)
        logging.info("txtResponseFileReader.py is running successfully")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running txtResponseFileReader.py: {e}")

if __name__ == "__main__":
    print("Choose an option:")
    print("1. Generate share payment/retrenchment directives")
    print("2. Cancel directives")
    print("3. Read directive response file")

    try:
        choice = int(input("Enter the number of your choice \n"))
    except ValueError:
        logging.error("Invalid input. Please enter a valid number.")

    if choice == 1:
        run_FileGenerator()
    elif choice == 2:
        run_FileGenerator_Cancellation()
    elif choice == 3:
        run_ResponseFileReader()
    else:
        print("Invalid option. Please choose a number between 1 and 3.")

print("FINISHED")

sys.exit(0)
