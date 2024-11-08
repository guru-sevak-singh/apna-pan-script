import fitz  # PyMuPDF
import os



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

# Replace 'your_pdf_file.pdf' with the path to your PDF file
pdf_file_path = 'guru.pdf'
print(read_pdf(pdf_file_path))

