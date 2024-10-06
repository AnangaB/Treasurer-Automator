import os
from datetime import datetime
import pytz 
import pandas as pd
import json
import sys
from pypdf import PdfWriter, PdfReader

#file path for common data, which is data used in all forms
common_data_file_path = r'./Reimbursement Data/common_info.json'
# file path 
raw_cheque_req_form_path =r'./Reimbursement Data/Blank Forms/raw_cheque_req_form.pdf'

""" Returns Requests df and Common Info df, after receiving files from terminal and accessing stored files. Program will exit, if improper data is received. 
    The requests data is a dataset with information to fill reimbursement forms with. 
    The common_info consists of data, that all reimbursement forms will contain.
"""
def get_requests_and_common_data():
    if len(sys.argv) != 2:
        print("Error: Invalid Input Passed in terminal. A command should be of the form:\n python fill_up_cheque_req.py dir/FileName.csv")
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

    return common_info, requests

def main():

    common_info, requests = get_requests_and_common_data()
    print(requests)
    print(requests.columns)

    # pdf_writer will contain all forms
    pdf_writer = PdfWriter()
    
    #the blank reimburement form, that will repeatedly get refernced to and filled out
    blank_reimbursement_form_pdf = PdfReader(raw_cheque_req_form_path)

    def fill_out_form(payee_details):
        
        form_data = {"Today s Date":common_info["date_today"],
                    'Group Name' : common_info["group_name"],
                    'Cheque Payable To print legibly': payee_details["Requester"],
                    'In The Amount Of':(payee_details["Amount"]),
                    'Describe the request andor provide additional information if necessary': payee_details["Reason"], 
                    'Requested By': common_info["requested_by"], 
                    'Position':common_info["position"],
                    'Picked up by':payee_details["Cheque Picked Up By"],
                    'Email':payee_details["Email"]
                    }
        
        #create a filled in cheque req form and append to pdf_writer
        filled_cheque_req_form =  PdfWriter(blank_reimbursement_form_pdf)

        filled_cheque_req_form.update_page_form_field_values(
            filled_cheque_req_form.pages[0], form_data,auto_regenerate=False
        )
        pdf_writer.add_page(filled_cheque_req_form.pages[0])

        # Mark this row as filled
        requests.loc[payee_details.name, 'Form Created'] = True

    requests.apply(fill_out_form, axis=1)

    # Write all filled forms into a single PDF
    combined_output_path = "output forms/combined_forms2.pdf"
    with open(combined_output_path, 'wb') as out_pdf:
        pdf_writer.write(out_pdf)
        

if __name__ == '__main__':
    main()

