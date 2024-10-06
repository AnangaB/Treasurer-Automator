import os
from datetime import datetime
import pytz 
import pandas as pd
import json
import sys
from pypdf import PdfWriter, PdfReader
from fill_reimbursement_form import get_single_filled_req_form

#file path for common data, which is data used in all forms
common_data_file_path = r'./Reimbursement Data/common_info.json'

""" Returns Requests df and Common Info df, after receiving files from terminal and accessing stored files. Program will exit, if improper data is received. 
    The requests data is a dataset with information to fill reimbursement forms with. 
    The common_info consists of data, that all reimbursement forms will contain.
"""
def get_requests_and_common_data():
    if len(sys.argv) != 4:
        print("Error: Invalid Input Passed in terminal. A command should be of the form:\n python fill_up_cheque_req.py dir/input_file.csv output_file_dir")
        sys.exit(1)

    # get file path of excel file that contains reimburesements values that need to be added into forms
    reimburesement_requests_file_path = sys.argv[1]
    
    try:
        # Attempt to open the file
        with open(reimburesement_requests_file_path, 'r') as file:
            print("The Reimbursement Request File has been found!")
    except FileNotFoundError:
        print("The Reimbursement Request File does not exist.")

    # Open and read the common data JSON file, which contains a few data that all reimburesment forms will have
    with open(common_data_file_path, 'r') as json_file:
        common_info = json.load(json_file)

    common_info["date_today"]=  datetime.now(pytz.timezone('US/Pacific')).strftime('%Y-%m-%d')

    #get csv of reimbursement requests
    requests = pd.read_csv(reimburesement_requests_file_path)
    
    execs_data_file_path =  sys.argv[2]
    execs_data = pd.read_csv(execs_data_file_path)

    return common_info, execs_data, requests

    """ Creates and saves a pdf, that contains all filled in cheque requestions for funds being taken from core.
    """
def complete_all_core_requests(core_requests, common_info, execs_data):
    
    #will contain all forms for core_requests reimbursements
    all_core_filled_forms = PdfWriter()
    
    def fill_form(payee_details):
        filtered_exec = execs_data[execs_data["Name"] == payee_details["Requester"]]
        # Check if the result is not empty
        exec_email = ""
        if not filtered_exec.empty:
            exec_email = filtered_exec["Email"].values[0] 

        filled_form = get_single_filled_req_form(common_info, payee_details, exec_email)
        all_core_filled_forms.add_page(filled_form)
    core_requests.apply(fill_form, axis=1)

    combined_output_path = sys.argv[3]
    output_file_name = str(common_info["date_today"]) + " " + str(len(core_requests)) + " core reimbursements" + ".pdf"
    # Write all filled forms into a single PDF
    with open(os.path.join(combined_output_path,output_file_name), 'wb') as out_pdf:
        all_core_filled_forms.write(out_pdf)

def main():

    common_info,execs_data, requests = get_requests_and_common_data()

    core_requests = requests.loc[(requests["Fund Type"] == "Core") & ((requests["Form Created"]) == False) ]
    if(len(core_requests) > 0):
        complete_all_core_requests(core_requests, common_info, execs_data )
        
if __name__ == '__main__':
    main()
