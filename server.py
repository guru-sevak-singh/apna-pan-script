from flask import Flask, request
from flask_cors import CORS
import time
import os
from fnmatch import fnmatch
import fitz  # PyMuPDF



app = Flask(__name__)
CORS(app=app)




def get_file(fiel_name):
    fle_str = f"*{fiel_name}*"
    directory_path = os.getcwd()  # Replace with the path to your folder

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if fnmatch(file, fle_str):
                p_number = file.split(".")[0]
                return p_number
    
    return "no_file"


def chck_for_file(file):
    return_value = False
    for n in range(5):
        if os.path.exists(file):
            return_value = True
            break
        else:
            time.sleep(1)
            continue
    return return_value


def read_pdf(file_path):
    pdf_document = fitz.open(file_path)

    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text = page.get_text()
        # print(f"Page {page_num + 1}:\n{text}\n")

    pdf_document.close()
    
    text_lsit = text.split('Proof of DOB')
    text_lsit = text_lsit[1].split('Date')
    text = text_lsit[0]
    text = text.replace("\n", "")
    text = text.replace(" ", "")
    text = text.replace("P-", "")
    os.rename(file_path, f"{text}.pdf")
    
    return text


# print(get_file('301kjhgfghjk56'))

@app.route('/get_ack_no')
def index():
    name = request.args.get('name')
    email = request.args.get('email')
    phone_number = request.args.get('phone_number')
    time.sleep(3)
    file_name_text = f"{name}={phone_number}={email}"
    for n in range(3):
        p_number = get_file(file_name_text)
        if p_number != "no_file":
            p_number = p_number.split("=")[0]
            break
        else:
            continue

    return {"p_number": p_number}

@app.route('/get_file_p_number')
# ######### /get_file_p_number?file_name=guru
def get_file_p_number():
    f_name = request.args.get("file_name")
    file_name = f_name + ".pdf"
    is_file = chck_for_file(file_name)
    if is_file == True:
        p_number = read_pdf(file_name)
        return {'p_number': p_number}
    else:
        return {"p_number": "no_file_found"}
    


if __name__ == "__main__":
    app.run(port='3389')