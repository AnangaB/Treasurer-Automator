import os
from fillpdf import fillpdfs
from datetime import datetime
import pytz 
import pandas as pd
import json
from PyPDF2 import PdfWriter, PdfReader

# Open and read the common data JSON file, which contains a few data that all reimburesment forms will have
with open(r'Common Data/common_info.json', 'r') as json_file:
    common_info = json.load(json_file)

common_info["date_today"]=  datetime.now(pytz.timezone('US/Pacific')).strftime('%Y-%m-%d')

#get csv of reimbursemets
requests = pd.read_csv(r"cheque req requests/cheque_req_requests.csv")
requests = requests.reset_index(drop=True)

# Create a PDF writer to combine forms
pdf_writer = PdfWriter()


def fill_out_form(payee_details):
    raw_form ="Common Data/raw_cheque_req_pdf.pdf"
    

    form_data = {"Today s Date":common_info["date_today"],
                'Group Name' : common_info["group_name"],
                'Cheque Payable To print legibly': common_info["requested_by"],
                'In The Amount Of':(payee_details["Amount"]),
                'Describe the request andor provide additional information if necessary': payee_details["Reason"], 
                'Requested By': common_info["requested_by"], 
                'Position':common_info["position"],
                'Picked up by':"",
                'Street Address':payee_details["Street Address"],
                'Email':'Email',
                'City, Province':payee_details["City And Province"],
                'Postal Code': payee_details["Postal Code"]
                }
                                                                                                                                                                      #Fill the PDF form into a temporary output path
    temp_output_path = f"temp_{payee_details.name}.pdf"
    fillpdfs.write_fillable_pdf(raw_form, temp_output_path, form_data, flatten=False)

    # Add the filled PDF to the writer
    with open(temp_output_path, 'rb') as f:
        pdf_reader = PdfReader(f)
        pdf_writer.add_page(pdf_reader.pages[0])  # Add the filled form page to the writer

    # Mark this row as filled
    requests.loc[payee_details.name, 'Form Created'] = True


    
print(requests.columns)

requests.apply(fill_out_form, axis=1)


# Write all filled forms into a single PDF
combined_output_path = "output forms/combined_forms.pdf"
with open(combined_output_path, 'wb') as out_pdf:
    pdf_writer.write(out_pdf)
    

#overwrite requests
requests.to_csv(r"cheque req requests/cheque_req_requests.csv")
