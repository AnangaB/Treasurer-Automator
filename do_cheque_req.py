import os
from datetime import datetime
import pytz 
import pandas as pd
import json
import sys
from pypdf import PdfWriter, PdfReader
from fill_reimbursement_form import get_single_filled_req_form
from pdf2image import convert_from_path


#file path for common data, which is data used in all forms
common_data_file_path = r'./Reimbursement Data/common_info.json'

#file path for grant summary page
grant_summary_page = r"./Reimbursement Data\Blank Forms/SFSS-Form-EventProjectSummary.pdf"

""" Returns Requests df and Common Info df, after receiving files from terminal and accessing stored files. Program will exit, if improper data is received. 
    The requests data is a dataset with information to fill reimbursement forms with. 
    The common_info consists of data, that all reimbursement forms will contain.

    The args passed onto the program are:
        - reimburesement_requests_file_path = sys.argv[1]
        - datasets containing data of full names and emails of exec = sys.argv[2]
        - past meeting minutes directory = sys.argv[3]
        - output dir for forms = sys.argv[4]



"""
def get_data_via_args():
    if len(sys.argv) != 5:
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
    
    meeting_mins_dir = sys.argv[3]
    output_dir = sys.argv[4]

    return common_info, execs_data, requests, meeting_mins_dir, output_dir

    """ Creates a pdf, that contains all filled in cheque requestions for funds
    """
def complete_all_individual_requests(requests, common_info, execs_data, meeting_mins_folder):
    
    #will contain all forms for requests reimbursements
    all_filled_forms = PdfWriter()
    
    def fill_form(payee_details):
        filtered_exec = execs_data[execs_data["Name"] == payee_details["Requester"]]
        # Check if the result is not empty
        exec_email = ""
        if not filtered_exec.empty:
            exec_email = filtered_exec["Email"].values[0] 

        filled_form = get_single_filled_req_form(common_info, payee_details, exec_email)
        all_filled_forms.add_page(filled_form)
        #add blank page after
        all_filled_forms.add_blank_page()
        #append relevant meeting minutes pdf as well
        mins_file_path = os.path.join(meeting_mins_folder, payee_details["Minutes File"])
        all_filled_forms.append(mins_file_path)
        #add blank page, if 
        if len(all_filled_forms.pages) % 2 != 0:
            all_filled_forms.add_blank_page()


    requests.apply(fill_form, axis=1)
    return all_filled_forms

def append_grant_summary_page_to_pdf(pdf_writer):
        #add blank page, if 
        if len(pdf_writer.pages) % 2 != 0:
            pdf_writer.add_blank_page()
        pdf_writer.append(grant_summary_page)

        return pdf_writer



def main():

    common_info,execs_data, requests, meeting_mins_dir, output_dir = get_data_via_args()

    core_requests = requests.loc[(requests["Fund Type"] == "Core") & ((requests["Form Created"]) == False) ]
    if(len(core_requests) > 0):
        filled_out_core_forms = complete_all_individual_requests(core_requests, common_info, execs_data, meeting_mins_dir )
        output_file_name = str(common_info["date_today"]) + " " + str(len(core_requests)) + " core reimbursements" + ".pdf"
        # Write all filled forms into a single PDF
        with open(os.path.join(output_dir,output_file_name), 'wb') as out_pdf:
            filled_out_core_forms.write(out_pdf)
            
            

    #do all grant requests
    grant_requests = requests.loc[(requests["Fund Type"] == "Grant") & ((requests["Form Created"]) == False) ]
    grant_ids = grant_requests["Grant ID"].dropna().unique()

    for id in grant_ids:
        id_specific_requests = grant_requests[grant_requests["Grant ID"] == id]
        if(len(id_specific_requests) > 0):
            filled_out_grant_cheque_reqs = complete_all_individual_requests(id_specific_requests, common_info, execs_data, meeting_mins_dir )
            
            #append grant event summary page to the back
            filled_out_grant_cheque_reqs = append_grant_summary_page_to_pdf(filled_out_grant_cheque_reqs)
            output_file_name = str(common_info["date_today"]) + "-grant-" +  str(int(id)) + "-reimbursements" + ".pdf"
            # Write all filled forms into a single PDF
            with open(os.path.join(output_dir,output_file_name), 'wb') as out_pdf:
                filled_out_grant_cheque_reqs.write(out_pdf)


        
if __name__ == '__main__':
    main()

