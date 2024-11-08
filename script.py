import fitz  # PyMuPDF
import os
import time
import requests

def get_latest_file_in_folder(dir_path, file_type):
  # Get a list of all the files in the directory with the specified file type
  files = [f for f in os.listdir(dir_path) if f.endswith(file_type)]

  # Create a tuple for each file that includes the file name and its modification time
  file_info = [(f, os.stat(os.path.join(dir_path, f)).st_mtime) for f in files]

  # Sort the list of tuples by the modification time
  file_info.sort(key=lambda x: x[1])

  # Get the last tuple in the sorted list, which will contain the file name and modification time of the latest file
  latest_file = file_info[-1][0]
  latest_time = file_info[-1][1]

  # Return the file name and modification time
  return latest_file

download_location = os.path.expanduser("~")+"/Downloads/"
file_name = get_latest_file_in_folder(download_location, ".pdf")
print(file_name)


# def ChangeFileName():
#     file_path = get_latest_file_in_folder(os.getcwd(), '.pdf')
#     pdf_document = fitz.open(file_path)
    
#     for page_num in range(pdf_document.page_count):
#         page = pdf_document[page_num]
#         text = page.get_text()
#         # print(f"Page {page_num + 1}:\n{text}\n")

#     pdf_document.close()
#     text = text.replace("\n", "")
#     text_lsit = text.split('Proof of DOB')
#     p_number = text_lsit[1].split('Date')[0]

#     p_number = p_number.replace("P - ", "")

#     other_stuff = text_lsit[1].split('Branch ID')[0]

#     other_stuff = other_stuff.split("INDIVIDUAL")[1]
#     name_and_phoone = other_stuff.split('AADHAAR')[0]

#     name = name_and_phoone.split("91-")[0]
#     phone_no = name_and_phoone.split("91-")[1]


#     email = other_stuff.split("E-mail ID")[1]

#     file_name = f"{p_number}={name}={phone_no}={email}.pdf"
#     print(file_name)

#     return file_name

# ChangeFileName()



