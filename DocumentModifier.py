import re
from docx import Document
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import threading  # Import threading module for concurrent processing

# Global variable to signal thread to stop
stop_processing = False

# Assuming the keyword_mapping is defined as provided in your question
keyword_mapping = {
    'Account_Type': 'Account Type',
    'Arbitrator_Reference_Letter': 'Arbitrator Reference Letter',
    'APPL_FATHERNAME': 'Appl Fathername',
    'Appl_Name': 'Appl_Name',
    'Arbitrator_Address': 'Arbitrator Address',
    'Arbitrator_Name': 'Arbitrator Name',
    'Appointof_Arbitrator_Letter_Date': 'Appoint-of Arbitrator Letter Date',
    'AGR_DATE': 'Agr_Date',
    'Ac_NO': 'Ac No',   
	'APPL_ADDRESS___PART_1': 'Appl Address _ Part 1',
    'Bank_Address': 'Bank_Address',
    'Bank_name': 'Bank Name',
    'Co_Borrower_Name': 'Co_Borrower Name',
	'CoApplicant_1': 'Co-Applicant_1',
	'CO_FATHERNAME1': 'Co_Fathername-1',
	'COB_ADDRESS___PART_1': 'Cob Address _ Part 1',
	'COBORROWER_2': 'Co-Borrower -2',
	'CoApplicant_2':'Co-Applicant_2',
	'CO_FATHERNAME2': 'Co_Fathername-2',
	'COB_ADDRESS___PART_2': 'Cob Address _ Part 2',
	'CoBorrower_3': 'Co-Borrower -3',
	'CoApplicant_3': 'Co-Applicant_3',
	'Co_Fathername3':'Co_Fathername-3',
	'Cob_Address___Part_3': 'Cob Address _ Part 3',
	'CoBorrower_4': 'Co-Borrower -4',
	'CoApplicant_4': 'Co-Applicant_4',
	'Co_Fathername4': 'Co_Fathername-4',
	'Cob_Address___Part_4': 'Cob Address _ Part 4',
	'CoBorrower_5': 'Co-Borrower -5',
	'CoApplicant_5': 'Co-Applicant_5',
	'Co_Fathername5': 'Co_Fathername-5',
	'Cob_Address___Part_5': 'Cob Address _ Part 5',
	'CoBorrower_6': 'Co-Borrower -6',
	'CoApplicant_6': 'Co-Applicant_6',
	'Co_Fathername6' : 'Co_Fathername-6',
	'Cob_Address___Part_6': 'Cob Address _ Part 6',
	'CoBorrower_7' : 'Co-Borrower -7',
	'CoApplicant_7' : 'Co-Applicant_7',
	'Co_Fathername7' : 'Co_Fathername-7',
    'Cob_Address___Part_7': 'Cob Address _ Part 7',
    'DEST_ACC_HOLDER': 'Dest_Acc_Holder',
    'Disc_Letter_Number': 'Disc_Letter_Number',
    'FCL_Amount': 'FCL Amount',
    'FCL_Amount_in_words': 'FCL Amount in words',
    'FCL_Date': 'FCL Date',
    'File_Number': 'File Number',
	'Guar_Name_1': 'Guar_Name 1',
	'Guarantor_1': 'Guarantor_1',
	'Guar_Fathername_1': 'Guar Fathername -1',
	'Guar_Address___Part_1': 'Guar Address _ Part 1',
    'Loan_No': 'Loan No',
    'Lrn_Date': 'Lrn Date',
    'Lrn_Amt' : 'Lrn Amt',
    'LRN_Amount_in_words': 'LRN Amount in words',
    'Product_Name': 'Product Name',
    'RefLetter_Ref_Number': 'RefLetter_Ref_Number',
    'PRODUCT': 'Product',
    'S__W__H_GUAR': 'S / W / H_GUAR',
    'S__W__H_APPL': 'S / W / H_APPL',
	'S__W__H_CB1': 'S / W / H_CB1',
	'S__W__H_CB2': 'S / W / H_CB2',
	'S__W__H_CB3': 'S / W / H_CB3',
    'S__W__H_CB4': 'S / W / H_CB4',
    'S__W__H_CB5': 'S / W / H_CB5',
    'S__W__H_CB6': 'S / W / H_CB6',
    'S__W__H_CB7': 'S / W / H_CB7',
    'Total_Amt_Finance': 'Total_Amt_Finance',
    'Total_Amt_Finance_in_words': 'Total_Amt_Finance in words',    
    'IFSC': 'Ifsc',
    'LRN_DATE': 'Lrn Date',
    'Statement_of_Claim_Schedule_date': 'Statement of Claim Schedule date',
    'SOC_Letter_Number': 'SOC_Letter_Number',
    'TOTAL_AMT_FINANCE':'Total_Amt_Finance',
    #'SRO_ADDRESS':'',
    #'SRO_NAME':'',   
}


def replace_keywords_in_paragraphs(paragraph, row, keyword_mapping):
    try:
        # Find all matches of the pattern in the paragraph text
        matches = re.findall(r'«([a-zA-Z_0-9]+)»', paragraph.text)
        
        # Iterate through matches and replace values with Excel data
        for match in matches:
            mapped_column_name = keyword_mapping.get(match, match)
            if mapped_column_name in row:
                replacement_value = str(row[mapped_column_name])
                if replacement_value == 'nan':
                    replacement_value =""
                paragraph.text = paragraph.text.replace(f'«{match}»', replacement_value)
    except Exception as e:
        print(f"Error in replace_keywords_in_paragraphs: {e}")
        error_message = f"Error: {str(e)}"
        status_label.config(text=error_message)
            

def replace_keywords_in_tables(table, excel_data, keyword_mapping):
    try:
        for row in table.rows:
            for cell in row.cells:
                # Find all matches of the pattern in the cell text
                matches = re.findall(r'«([a-zA-Z_0-9]+)»', cell.text)
                
                # Iterate through matches and replace values with Excel data
                for match in matches:
                    mapped_column_name = keyword_mapping.get(match, match)
                    if mapped_column_name in excel_data:
                        replacement_value = str(excel_data[mapped_column_name])
                        if replacement_value == 'nan':
                            replacement_value=""
                        cell.text = cell.text.replace(f'«{match}»', replacement_value)
    except Exception as e:
        print(f"Error in replace_keywords_in_tables: {e}")
        error_message = f"Error: {str(e)}"
        status_label.config(text=error_message)                  


def replace_keywords_in_word_document(word_template_path, excel_file_path, output_directory, keyword_mapping, sheet_name):

    global stop_processing  # Use the global variable
    # Read the specific sheet from the Excel file
    excel_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    for _, row in excel_data.iterrows():
        if stop_processing:
            print("Processing stopped.")
            status_label.config(text="Processing stopped.")
            break  # Exit the loop if stop_processing is True

        doc = Document(word_template_path)

        # Replace keywords in paragraphs
        for paragraph in doc.paragraphs:

            replace_keywords_in_paragraphs(paragraph, row, keyword_mapping)

        # Replace keywords in tables
        for table in doc.tables:
            replace_keywords_in_tables(table, row, keyword_mapping)

        # Extract Loan No from the Excel data
        loan_no = str(row['Loan No'])

        # Save the modified document to the specified output path
        output_filename = f"{output_directory}/{loan_no}.docx"
        doc.save(output_filename)

        print(f"The modified document has been saved to: {output_filename}")

    if not stop_processing:
        print("Document processing completed.")
        status_label.config(text="Document processing completed.")

def stop_processing_thread():
    global stop_processing
    stop_processing = True
    print("Stopping document processing...")


def browse_excel_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(0, filename)


def browse_word_template():
    filename = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    word_template_entry.delete(0, tk.END)
    word_template_entry.insert(0, filename)


def browse_output_folder():
    foldername = filedialog.askdirectory()
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, foldername)


def process_document():
    global stop_processing
    stop_processing = False
    try:
        excel_file_path = excel_file_entry.get()
        word_template_path = word_template_entry.get()
        output_folder_path = output_folder_entry.get()
        sheet_name = sheet_name_entry.get()

        if not excel_file_path or not word_template_path or not output_folder_path or not sheet_name:
            status_label.config(text="All fields must be filled in.")
            raise ValueError("All fields must be filled in.")
        
  
        # Check if the sheet name exists in the Excel file
        excel_data_sheets = pd.ExcelFile(excel_file_path).sheet_names
        if sheet_name not in excel_data_sheets:
            status_label.config(text=f"Sheet name '{sheet_name}' not found in the Excel file.")
            raise ValueError(f"Sheet name '{sheet_name}' not found in the Excel file.")


        print("Processing Document...")
        status_label.config(text="Processing document...")
        processing_thread = threading.Thread(target=replace_keywords_in_word_document, args=(word_template_path, excel_file_path, output_folder_path, keyword_mapping, sheet_name))
        processing_thread.start()
        #replace_keywords_in_word_document(word_template_path, excel_file_path, output_folder_path, keyword_mapping, sheet_name)

    except Exception as e:
        print(f"Error in process_document: {e}")

# UI setup
root = tk.Tk()
root.title("Document Modifier")

# Labels
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Label(root, text="Word Template:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Label(root, text="Output Folder:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
tk.Label(root, text="Sheet Name:").grid(row=3, column=0, padx=5, pady=5, sticky="e")

# Entry widgets
excel_file_entry = tk.Entry(root, width=50)
word_template_entry = tk.Entry(root, width=50)
output_folder_entry = tk.Entry(root, width=50)
sheet_name_entry = tk.Entry(root, width=30)

excel_file_entry.grid(row=0, column=1, padx=5, pady=5)
word_template_entry.grid(row=1, column=1, padx=5, pady=5)
output_folder_entry.grid(row=2, column=1, padx=5, pady=5)
sheet_name_entry.grid(row=3, column=1, padx=5, pady=5)

# Buttons
tk.Button(root, text="Browse", command=browse_excel_file).grid(row=0, column=2, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_word_template).grid(row=1, column=2, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_output_folder).grid(row=2, column=2, padx=5, pady=5)
tk.Button(root, text="Process Document", command=process_document).grid(row=4, column=1, pady=10)
# Button to stop processing
tk.Button(root, text="Stop Processing", command=stop_processing_thread).grid(row=5, column=1, pady=10)

# Status label
status_label = tk.Label(root, text="", fg="blue")
status_label.grid(row=6, column=1, pady=5)

root.mainloop()
