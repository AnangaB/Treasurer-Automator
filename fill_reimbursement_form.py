from pypdf import PdfWriter, PdfReader
import pypdf
from pypdf.generic import NameObject
# file path for the blank cheque req
raw_cheque_req_form_path =r'./Reimbursement Data/Blank Forms/raw_cheque_req_form.pdf'

 #the blank reimburement form, that will repeatedly get refernced to and filled out
blank_reimbursement_form_pdf = PdfReader(raw_cheque_req_form_path)

"""Fills a single cheque req form and returns it as a pypdf PageObject
"""
def get_single_filled_req_form(common_info, payee_details,exec_email):

    #fix the forms first

    form_data = {'Today s Date':(common_info["date_today"]),
                'Group Name' : common_info["group_name"],
                'Cheque Payable To print legibly': payee_details["Requester"],
                'In The Amount Of':(payee_details["Amount"]),
                'Describe the request andor provide additional information if necessary': str(payee_details["Reason"]), 
                'Requested By': common_info["requested_by"], 
                'Position':common_info["position"],
                'Picked up by':payee_details["Requester"],
                'Email':exec_email
                }

#

    #create a filled in cheque req form and append to pdf_writer
    filled_cheque_req_form =  PdfWriter(blank_reimbursement_form_pdf)
   
    # Update the fields with the form_data
    filled_cheque_req_form.update_page_form_field_values(
        filled_cheque_req_form.pages[0], form_data,
           auto_regenerate=False
    )


    
    return filled_cheque_req_form.pages[0]
