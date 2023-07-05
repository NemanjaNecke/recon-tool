import tkinter as tk
from tkinter import Tk, filedialog
import re
from datetime import datetime, timedelta
import os
import pandas as pd
import tabula
import numpy as np
from pathlib import Path
from PyPDF2 import PdfReader


class Reconciler():
    def __init__(self, root):
        self.root = root
        self.root.geometry("700x500")
        self.root.title("Reconciler")
        # Choose type of recon
        
        self.choose_label = tk.Label(self.root, text="Choose type of a reconiliation")
        self.choose_cp_button = tk.Button(self.root, text="CP Recon", command = self.choose_cp_recon, bg='green', fg='white')
        self.choose_cp_button.pack()
        self.choose_edc_button = tk.Button(self.root, text="EDC Recon", command=self.choose_edc, bg='red', fg='white')
        self.choose_edc_button.pack()
        self.choose_cp_trans_button = tk.Button(self.root, text="CP Trans Recon", command = self.choose_cp_transaction, bg='green', fg='red')
        self.choose_cp_trans_button.pack()
        self.choose_fund_settle_button = tk.Button(self.root, text='Funding vs Settlement', command= self.choose_fund_settle)
        self.choose_fund_settle_button.pack()
        self.convert_to_excel = tk.Button(self.root, text="Convert PDF to Excel", command=self.choose_pdf_to_excel)
        self.convert_to_excel.pack()
        self.home_dir = os.path.expanduser('~/OneDrive - NCR Corporation/Desktop')
        
    # Convert pdf to excel tkinter
    choose_convert_switch = True
    def choose_pdf_to_excel(self):
        if self.choose_convert_switch == True:
            self.choose_convert_switch = False
            self.convert_to_excel_label = tk.Label(self.root, text="Convert PDF Table to Excel")
            self.convert_to_excel_label.pack()
            self.pdf_file_load = tk.Button(self.root, text="Choose PDF to convert", command=self.load_pdf_file)
            self.pdf_file_load.pack()
            self.page_num = tk.Entry(self.root)
            self.page_num.pack()
            self.original_file_label = tk.Label(self.root, text='')
            self.original_file_label.pack()
            self.convert = tk.Button(self.root, text="Convert", command=lambda: self.convert_file(self.page_num.get()))
            self.convert.pack()
            self.converted_file = tk.Label(self.root, text="")
            self.converted_file.pack()
        else:
            self.choose_convert_switch = True
            self.convert_to_excel_label.pack_forget()
            self.pdf_file_load.pack_forget()
            self.page_num.pack_forget()
            self.original_file_label.pack_forget()
            self.convert.pack_forget()
            self.converted_file.pack_forget()
    # Load PDF for pdf to excel
    def load_pdf_file(self):
        self.pdf_file = filedialog.askopenfilename(initialdir="/", title="Select PDF",
                                                        filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
        if self.pdf_file:
            self.original_file_label.config(text=f"PDF: {os.path.basename(self.pdf_file)}")        
    # Convert pdf to excel
    def convert_file(self, pagenum):
        pdf_file = tabula.read_pdf(self.pdf_file, pages=f'1-{pagenum}')
        df = pd.concat(pdf_file, ignore_index=True)
        filename = f"PDF to excel.xlsx"
        current_dir = os.getcwd()
        parent_dir = self.home_dir

        # Create a folder named "duplicates" in the parent directory if it doesn't exist
        folder_name = 'PDF to excel'
        folder_path = os.path.join(parent_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        output_file = os.path.join(folder_path, filename)
        df.to_excel(output_file)
        self.converted_file.config(text=f"Converted file saved to {filename}")
    # Funding vs Settlement tkinter
    choose_fund_settle_switch = True
    def choose_fund_settle(self):
        if self.choose_fund_settle_switch == True:
            self.choose_fund_settle_switch = False
            self.fund_settle_label = tk.Label(self.root, text="Compare funding and settlement from NCR Reporting")
            self.fund_settle_label.pack()
            self.funding_report = tk.Button(self.root, text="Choose funding report", command=self.choose_funding)
            self.funding_report.pack()
            self.funding_report_label = tk.Label(self.root, text="Funding: ")
            self.funding_report_label.pack()
            self.settlement_report = tk.Button(self.root, text="Choose settlement report", command=self.choose_settlement_report)
            self.settlement_report.pack()
            self.settlement_report_label = tk.Label(self.root, text="Settlement: ")
            self.settlement_report_label.pack()
            self.file_run_fund_label = tk.Label(self.root, text="")
            self.file_run_fund_label.pack()
            self.run_fund = tk.Button(self.root, text="Run recon", command=self.compare_funding_settlement)
            self.run_fund.pack()
        else:
            self.choose_fund_settle_switch = True
            self.settlement_report_label.pack_forget()
            self.fund_settle_label.pack_forget()
            self.funding_report.pack_forget()
            self.funding_report_label.pack_forget()
            self.settlement_report.pack_forget()
            self.file_run_fund_label.pack_forget()
            self.run_fund.pack_forget()
    # Load Deposits and Settlement files
    def choose_funding(self):
        self.funding_file = filedialog.askopenfilename(initialdir="/", title="Select Deposits Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.funding_file:
            self.funding_report_label.config(text=f"Deposits Excel: {os.path.basename(self.funding_file)}")
    def choose_settlement_report(self):
        self.settlement_fundfile = filedialog.askopenfilename(initialdir="/", title="Select settlement Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.settlement_fundfile:
            self.settlement_report_label.config(text=f"Settlement Excel: {os.path.basename(self.settlement_fundfile)}")
    # Find transactions that are in the settlement, but not in the funding
    def compare_funding_settlement(self):
        funds_list = pd.read_excel(self.funding_file, header=1)
        settlement_list = pd.read_excel(self.settlement_fundfile)
        funds_list['Transaction ID'] = funds_list['Transaction ID'].astype(str)
        settlement_list['Transaction ID'] = settlement_list['Transaction ID'].astype(str)
        funds_list.rename(columns={
            'Card Num': 'Card Number'
        }, inplace=True)
        left_join = pd.merge(settlement_list[['Auth Code', 'Card Number', 'Sales Amount', 'Transaction ID']], funds_list[['Transaction ID', 'Card Number', 'Amount']], how="left", on=['Transaction ID', 'Card Number'], indicator='Match')
        self.name =  settlement_list.iloc[0]['Merchant Name']
        left_join['Transaction ID'] = left_join['Transaction ID'].astype(str).str.replace('.', '', regex=False)
        filename = f"funding-settlement_reconciliation_results_{self.name}.xlsx"
        current_dir = os.getcwd()
        parent_dir = self.home_dir

        # Create a folder named "duplicates" in the parent directory if it doesn't exist
        folder_name = 'Settlement vs Funding'
        folder_path = os.path.join(parent_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        output_file = os.path.join(folder_path, filename)
        left_join.to_excel(output_file,index=False)
        self.file_run_fund_label.config(text=f"Results saved to {filename}")        
    # EDC Recon GUI
    choose_edc_switch = True
    def choose_edc(self):
        if self.choose_edc_switch == True:
            self.choose_edc_switch = False
            self.edc_label = tk.Label(self.root, text="EDC Report")
            self.edc_label.pack()
            self.edc_report = tk.Button(self.root, text="Choose file...", command=self.choose_edc_file)
            self.edc_report.pack()            
            self.edc_page_start_label = tk.Label(self.root, text='Enter start page')
            self.edc_page_start_label.pack()
            self.edc_page_start = tk.Entry(self.root)
            self.edc_page_start.pack()
            self.edc_page_num_label = tk.Label(self.root, text='Enter number of pages with transactions')
            self.edc_page_num_label.pack()
            self.edc_page_num = tk.Entry(self.root)
            self.edc_page_num.pack()

            self.authorization_file_label = tk.Label(self.root, text="Authorization Excel:")
            self.authorization_file_label.pack()
            self.authorization_file_button = tk.Button(self.root, text="Choose file...", command=self.choose_auth_file)
            self.authorization_file_button.pack()            
            self.settlement_file_label = tk.Label(self.root, text="Settlement Excel:")
            self.settlement_file_label.pack()
            self.settlement_file_button = tk.Button(self.root, text="Choose file...", command=self.choose_settlement_file)
            self.settlement_file_button.pack()

            self.file_label = tk.Label(self.root, text="")
            self.file_label.pack()
            self.run_edc = tk.Button(self.root, text="Run recon", command=lambda: self.run_reconciler_edc(self.edc_page_start.get(), self.edc_page_num.get()))
            self.run_edc.pack()
        else:
            self.choose_edc_switch = True
            self.edc_label.pack_forget()
            self.edc_report.pack_forget()
            self.settlement_file_label.pack_forget()
            self.settlement_file_button.pack_forget()
            self.run_edc.pack_forget()
            self.file_label.pack_forget()
            self.edc_page_start_label.pack_forget()
            self.edc_page_start.pack_forget()
            self.edc_page_num.pack_forget()
            self.edc_page_num_label.pack_forget()
            self.authorization_file_button.pack_forget()
            self.authorization_file_label.pack_forget()
            
    # CP recon GUI
    choose_cp_switch = True
    def choose_cp_recon(self):
        if self.choose_cp_switch == True:
            self.choose_cp_switch = False
            self.is_capn = tk.BooleanVar()
            self.is_capn_checkbox = tk.Checkbutton(self.root, text="CAPN Merchant", variable=self.is_capn)
            self.is_capn_checkbox.pack()
            self.instore_label = tk.Label(self.root, text="Instore Batches PDF:")
            self.instore_label.pack()
            self.instore_button = tk.Button(self.root, text="Choose file...", command=self.choose_instore_file)
            self.instore_button.pack()

            self.ecomm_label = tk.Label(self.root, text="Ecomm Batches PDF:")
            self.ecomm_label.pack()
            self.ecomm_button = tk.Button(self.root, text="Choose file...", command=self.choose_ecomm_file)
            self.ecomm_button.pack()

            self.deposits_label = tk.Label(self.root, text="Deposits Excel:")
            self.deposits_label.pack()
            self.deposits_button = tk.Button(self.root, text="Choose file...", command=self.choose_deposits_file)
            self.deposits_button.pack()

            self.run_button = tk.Button(self.root, text="Run Reconciler", command=self.run_reconciler_cp)
            self.run_button.pack()

            self.exit_button = tk.Button(self.root, text="Exit", command=self.root.quit)
            self.exit_button.pack()

            self.file_label = tk.Label(self.root, text="")
            self.file_label.pack()
        else:
            self.choose_cp_switch = True
            self.instore_button.pack_forget()
            self.ecomm_label.pack_forget()
            self.ecomm_button.pack_forget()
            self.instore_label.pack_forget()
            self.deposits_label.pack_forget()
            self.deposits_button.pack_forget()
            self.run_button.pack_forget()
            self.file_label.pack_forget()
            self.exit_button.pack_forget()
            self.is_capn_checkbox.pack_forget()
        self.error_window = None

    # CP transaction GUI setup
    choose_cp_trans = True
    def choose_cp_transaction(self):
        if self.choose_cp_trans == True:
            self.choose_cp_trans = False
            self.cp_trans_label = tk.Label(self.root, text="CP transactions file")
            self.cp_trans_label.pack()
            self.cp_trans_button = tk.Button(self.root, text="Choose file...", command=self.choose_cp_trans_file)
            self.cp_trans_button.pack()
            self.auth_file_label = tk.Label(self.root, text="Authorization File")
            self.auth_file_label.pack()
            self.cp_auth_button = tk.Button(self.root, text="Choose file...", command=self.choose_cp_auth)
            self.cp_auth_button.pack()
            self.settlement_file_label = tk.Label(self.root, text= "Setlement File")
            self.settlement_file_label.pack()
            self.cp_trans_settle_button = tk.Button(self.root, text="Choose file...", command=self.choose_cp_settlement)
            self.cp_trans_settle_button.pack()
            self.run_button = tk.Button(self.root, text="Run Reconciler", command=self.run_reconciler_cp_trans)
            self.run_button.pack()
            self.trans_file_label = tk.Label(self.root, text="")
            self.trans_file_label.pack()
            self.file_save = tk.Label(self.root, text='')
            self.file_save.pack()
        else:
            self.choose_cp_trans = True
            self.cp_trans_label.pack_forget()
            self.cp_trans_button.pack_forget()
            self.cp_trans_settle_button.pack_forget()
            self.trans_file_label.pack_forget()
            self.settlement_file_label.pack_forget()
            self.auth_file_label.pack_forget()
            self.cp_auth_button.pack_forget()
            self.trans_file_label.pack_forget()
            self.file_save.pack_forget()
            self.run_button.pack_forget()
    # CP vs NPP transaction compare
    def choose_cp_trans_file(self):
        self.cp_trans = filedialog.askopenfilename(initialdir='/', title="Select CP trans file",
                                                   filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("all files", "*.*")))
        if self.cp_trans:
            self.cp_trans_label.config(text=f"CP Transaction file: {os.path.basename(self.cp_trans)}")
    def choose_cp_settlement(self):
        self.cp_settlement = filedialog.askopenfilename(initialdir="/", title="Select settlement Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("all files", "*.*")))
        if self.cp_settlement:
            self.settlement_file_label.config(text=f"Settlement Excel: {os.path.basename(self.cp_settlement)}")
    def choose_cp_auth(self):
        self.cp_auth = filedialog.askopenfilename(initialdir="/", title="Select authorization Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("all files", "*.*")))
        if self.cp_auth:
            self.auth_file_label.config(text=f"Authorization Excel: {os.path.basename(self.cp_auth)}")
    # CP vs NPP deposits vs batches
    def choose_instore_file(self):
        self.instore_file = filedialog.askopenfilename(initialdir="/", title="Select Instore Batches PDF",
                                                       filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
        if self.instore_file:
            self.instore_label.config(text=f"Instore Batches PDF: {os.path.basename(self.instore_file)}")
    def choose_ecomm_file(self):
        self.ecomm_file = filedialog.askopenfilename(initialdir="/", title="Select Ecomm Batches PDF",
                                                     filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
        if self.ecomm_file:
            self.ecomm_label.config(text=f"Ecomm Batches PDF: {os.path.basename(self.ecomm_file)}")

    def choose_deposits_file(self):
        self.deposits_file = filedialog.askopenfilename(initialdir="/", title="Select Deposits Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.deposits_file:
            self.deposits_label.config(text=f"Deposits Excel: {os.path.basename(self.deposits_file)}")
    # EDC Recon file load
    def choose_edc_file(self):
        self.edc_file = filedialog.askopenfilename(initialdir="/", title="Select EDC PDF",
                                                       filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
        if self.edc_file:
            self.edc_label.config(text=f"EDC PDF: {os.path.basename(self.edc_file)}")

    def choose_settlement_file(self):
        self.settlement_file = filedialog.askopenfilename(initialdir="/", title="Select settlement Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.settlement_file:
            self.settlement_file_label.config(text=f"Settlement Excel: {os.path.basename(self.settlement_file)}")

    def choose_auth_file(self):
        self.authorization_file = filedialog.askopenfilename(initialdir="/", title="Select authorization Excel",
                                                        filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.authorization_file:
            self.authorization_file_label.config(text=f"Authorization Excel: {os.path.basename(self.authorization_file)}")
    # Reconcile CP vs NPP transactions
    def run_reconciler_cp_trans(self):
        cp_transaction_list = pd.read_html(self.cp_trans)
        cp_transaction_list = pd.concat(cp_transaction_list)
        cp_transaction_list['AccountNumberLast4'] = cp_transaction_list['AccountNumberLast4'].astype(str).str.zfill(4)
        settlement_list = pd.read_excel(self.cp_settlement, dtype={'Transaction ID': str})
        authorization_list = pd.read_excel(self.cp_auth, dtype={'Transaction ID': str})
        left_join = pd.merge(cp_transaction_list[['AccountNumberLast4', 'ApprovedAmount', 'Auth Code']], authorization_list[['Auth Code', 'Transaction ID', 'Card Number']], how="left", on='Auth Code', indicator='status auth')
        left_join = pd.merge(left_join, settlement_list[['Auth Code', 'Transaction ID', 'Card Number']], how="left", on='Auth Code', indicator='status settle')
        self.name =  settlement_list.iloc[0]['Merchant Name']
        filename = f"transaction_reconciliation_results_{self.name}.xlsx"
        left_join.rename(columns={
            'Transaction ID_x': 'Transaction ID authorization', 
            'Card Number_x': 'Card Number authorization',
            'Transaction ID_y': 'Transaction ID settlement', 
            'Card Number_y': 'Card Number settlement'
        }, inplace = True)
        results = {
            'left_only': "Didn't reach NPP",
            'right_only': "Match",
            'both': "Match"
        }
        left_join['status auth'] = left_join['status auth'].map(results).astype(str)
        left_join['status settle'] = left_join['status settle'].map(results).astype(str)
        current_dir = os.getcwd()
        parent_dir = self.home_dir

        # Create a folder named "duplicates" in the parent directory if it doesn't exist
        folder_name = 'PDF to excel'
        folder_path = os.path.join(parent_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        output_file = os.path.join(folder_path, filename)
        left_join.to_excel(output_file,index=False)
        self.file_save.config(text=f"Results saved to {filename}")
    # EDC Reconciler
    import re

    def run_reconciler_edc(self, startpage, pagenum):
        authorization_trans_list = pd.read_excel(self.authorization_file)
        settlement_trans_list = pd.read_excel(self.settlement_file)
        pdf = tabula.read_pdf(self.edc_file, pages=f'{startpage}-{pagenum}', stream=True)
        # Create a Pandas DataFrame from the PDF
        df = pd.concat(pdf, ignore_index=True)

        card_cols = df.filter(regex='Card').columns
        card_number_regex = r'(?<!\d)(?:X{11}\d{4}|\d{4}X{11})(?!\d)|\b(?:\S*X{11}\d{4}\S*|\S*\d{4}X{11}\S*)\b'
        df['Card Number'] = df[card_cols].apply(lambda x: ''.join(re.findall(card_number_regex, ' '.join(x.astype(str)))), axis=1)

        df.drop(card_cols, axis=1, inplace=True)

        # Exclude 'Check' column
        check_cols = df.filter(like='Check').columns
        df_without_check = df.drop(columns=check_cols)
        # Extract all Auth Codes from the entire document
        auth_code_regex = r'\b[A-Z0-9]+\b'
        df_without_check['Auth'] = df_without_check.astype(str).apply(lambda x: ', '.join(re.findall(auth_code_regex, ' '.join(x))), axis=1)

        # Extract the desired Auth code (6 characters) from the captured string
        auth_code_extract_regex = r'(?<!\w)(\w{6})(?!\w)'
        df_without_check['Auth'] = df_without_check['Auth'].str.findall(auth_code_extract_regex)

        # Pick the second one if exists, or the first one otherwise
        df_without_check['Auth'] = df_without_check['Auth'].apply(lambda x: x[1] if len(x) > 1 else x[0] if x else np.nan)
                # Extract the date range from the PDF
        pdf = PdfReader(self.edc_file)
        text = ''
        for page in pdf.pages:
            text += page.extract_text()

        date_regex = r'EDC Transaction Report.*?(\d{2}/\d{2}/\d{4})(\s*--\s*(\d{2}/\d{2}/\d{4}))?'
        date_match = re.search(date_regex, text, re.DOTALL)
        if date_match:
            start_date = date_match.group(1)
            end_date = date_match.group(3) if date_match.group(3) else start_date
        else:
            start_date = ''
            end_date = ''
        start_date = start_date.replace('/', '-')
        end_date = end_date.replace('/', '-')
        # Convert settlement_sheet to a pandas DataFrame
        df2 = pd.DataFrame(authorization_trans_list)
        df3 = pd.DataFrame(settlement_trans_list)
        df_without_check['Auth'] = df_without_check['Auth'].astype(str)
        left_join = pd.merge(df_without_check[['Card Number', 'Auth', 'Amount', 'Tip', 'Total']], df2[['Card Number', 'Auth Code', 'Transaction ID']], left_on='Auth', right_on='Auth Code', how='left')
        left_join = pd.merge(left_join, df3[['Card Number', 'Auth Code', 'Transaction ID', 'Sales Amount']], left_on='Auth', right_on='Auth Code', how='left')
        left_join.drop('Card Number_y', axis=1, inplace=True)

        left_join.rename(columns={
            'Card Number_x': 'Card',
            'Transaction ID_x': 'Transaction ID authorization',
            'Auth Code_x': 'Auth code authorization',
            'Transaction ID_y': 'Transaction ID settlement',
            'Auth Code_y': 'Auth Code settlement'
        }, inplace=True)

        columns_to_check = ['Card Number', 'Auth', 'Amount', 'Tip', 'Total']
        left_join.drop_duplicates(subset=columns_to_check, inplace=True)

        self.name = df2.iloc[0]['Merchant Name']
        filename = f"edc_reconciliation_results_{self.name}_{start_date}.xlsx"
        current_dir = os.getcwd()
        parent_dir = self.home_dir
        folder_name = 'EDC Recon'
        folder_path = os.path.join(parent_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        output_file = os.path.join(folder_path, filename)
        
        left_join.to_excel(output_file, index=False)

        self.file_label.config(text=f"Results saved to {filename}")


    # Reconcile Batches and Deposits    
    def run_reconciler_cp(self):
        # Load deposits Excel file
        deposits_df = pd.read_excel(self.deposits_file, header=1)
        deposits_df = deposits_df.drop(deposits_df.tail(1).index)
        
        dba_name = deposits_df['DBA Name'][2]
        # Extract deposit amounts and dates from deposits Excel file
        deposit_df = deposits_df[['Amount', 'Payment Effect Date']].rename(columns={'Payment Effect Date': 'Deposit Date', 'Amount': 'Deposit Amount'})
        deposit_df['Deposit Date'] = pd.to_datetime(deposit_df['Deposit Date'], format='%Y-%m-%d').dt.date
        deposit_df = deposit_df[deposits_df["Amount"] >= 0.1]
        # Load instore batches PDF file
        instore_tables = tabula.read_pdf(self.instore_file, pages='all', multiple_tables=False)
        instore_text = ''
        for table in instore_tables:
            instore_text += table.to_string()

        # Extract batch amounts and dates from instore batches PDF file
        instore_batch_amounts = re.findall(r'Settlement Totals\s+\d+\s+\$\d{1,3}(?:,\d{3})*\.\d{2}\s+\d+\s+(\$\d{1,3}(?:,\d{3})*\.\d{2})', instore_text)
        instore_batch_dates = re.findall(r'(?:POS DOB|ServerEPS DOB).*?(\d{1,2}/\d{1,2}/\d{4})', instore_text, re.DOTALL)
        instore_batch_dates = [datetime.strptime(d, '%m/%d/%Y').date() for d in instore_batch_dates]
        amex_amounts = re.findall(r'American Express Sales\s+\d+\s+\$\d{1,3}(?:,\d{3})*\.\d{2}\s+\d+\s+(\$\d{1,3}(?:,\d{3})*\.\d{2})', instore_text)
        # Load ecomm batches PDF file if it exists
        if hasattr(self, 'ecomm_file') and self.ecomm_file and os.path.isfile(self.ecomm_file):
            ecomm_tables = tabula.read_pdf(self.ecomm_file, pages='all')
            ecomm_text = ''
            for table in ecomm_tables:
                ecomm_text += table.to_string()
            # Extract batch amounts and dates from ecomm batches PDF file
            ecomm_batch_amounts = re.findall(r'Settlement Totals\s+\d+\s+\$\d{1,3}(?:,\d{3})*\.\d{2}\s+\d+\s+(\$\d{1,3}(?:,\d{3})*\.\d{2})', ecomm_text)
            ecomm_batch_dates = re.findall(r'(?:POS DOB|ServerEPS DOB).*?(\d{1,2}/\d{1,2}/\d{4})', ecomm_text, re.DOTALL)
            ecomm_batch_dates = [datetime.strptime(d, '%m/%d/%Y').date() for d in ecomm_batch_dates]
        else:
            ecomm_batch_amounts = []
            ecomm_batch_dates = []

        # Combine the instore and ecomm batch amounts and dates into a single DataFrame
    
        instore_batch_df = pd.DataFrame({'Batch Amount': instore_batch_amounts, 'Batch Date': instore_batch_dates, 'Batch Type': 'Instore'})
        ecomm_batch_df = pd.DataFrame({'Batch Amount': ecomm_batch_amounts, 'Batch Date': ecomm_batch_dates, 'Batch Type': 'Ecomm'})
        combined_batch_df = pd.concat([instore_batch_df, ecomm_batch_df], ignore_index=True)
        combined_batch_df['Batch Amount'] = combined_batch_df['Batch Amount'].str.replace('$', '').str.replace(',', '').astype(float)

        combined_batch_df = combined_batch_df[combined_batch_df['Batch Amount'] > 0]


        results = {
            'left_only': "Found in batches // doesn't match the deposit",
            'right_only': "Check the risk // Found in deposits // doesn't match the batch",
            'both': "Match"
        }

        # Group by same date if there's ecomm batch and sum
        grouped_batch_df = combined_batch_df.groupby('Batch Date')['Batch Amount'].sum().reset_index()

        if self.is_capn.get():
            # Deduct AMEX amounts from batch amounts
            for i, amex_amount in enumerate(amex_amounts):
                if i < len(grouped_batch_df):
                    grouped_batch_df.loc[i, 'Batch Amount'] -= float(amex_amount.replace('$', '').replace(',', ''))

        else:
            # No deduction required
            pass

        # Iterate over instore and ecomm dates to find matching date pairs
        matched_deposits_data = []
        for i, instore_date in enumerate(instore_batch_dates):
            instore_amount = instore_batch_df.loc[i, 'Batch Amount']
            for j, ecomm_date in enumerate(ecomm_batch_dates):
                if ecomm_date == instore_date + pd.DateOffset(days=1) and instore_amount > 0:
                    ecomm_amount = ecomm_batch_df.loc[j, 'Batch Amount']
                    total_amount = instore_amount + ecomm_amount
                    matching_deposit = deposit_df.loc[deposit_df['Deposit Amount'] == total_amount]
                    if len(matching_deposit) > 0:
                        deposit_date = matching_deposit['Deposit Date'].iloc[0]
                        matched_deposits_data.append({'Deposit Amount': total_amount, 'Deposit Date': deposit_date, 'Batch Date': instore_date, 'Batch amount': total_amount})


        matched_deposits_df = pd.DataFrame(matched_deposits_data)

        
        # Create a new column with the date shifted by one day
        instore_batch_df['Next Day'] = instore_batch_df['Batch Date'] + pd.DateOffset(days=1)
        instore_batch_df['Next Day'] = pd.to_datetime(instore_batch_df['Next Day'])
        instore_batch_df['Next Day'] = instore_batch_df['Next Day'].dt.date

        

        # Drop the last row which will have NaN value for Next Day Batch Amount
        instore_batch_df = instore_batch_df.iloc[:-1]
        # Group by consecutive pairs of days and sum the batch amounts
   # Group by consecutive pairs of days and sum the batch amounts
        instore_batch_df['Batch Amount'] = instore_batch_df['Batch Amount'].apply(lambda x: float(str(x).replace('$', '').replace(',', '')))
        # Shift Batch Amount column by one day to get next day amounts and combine days
        instore_batch_df['Next Day Batch Amount'] = instore_batch_df['Batch Amount'].shift(-1)
        instore_batch_df['Combined Batch'] = instore_batch_df.apply(lambda row: round(row['Batch Amount'] + row['Next Day Batch Amount'], 2) if pd.notna(row['Next Day Batch Amount']) and row['Batch Amount'] > 0 and row['Next Day Batch Amount'] > 0 else round(row['Batch Amount'], 2), axis=1)
        instore_batch_df = instore_batch_df.reset_index()

        consecutive = pd.merge(deposit_df, instore_batch_df[['Combined Batch', 'Next Day']], how='outer', left_on='Deposit Amount', right_on=['Combined Batch'], indicator='status recon')
        
        # Merge grouped by date with deposits to check if there's a combined deposit
        grouped_batch_merge = pd.merge(grouped_batch_df, deposit_df, how='outer', right_on='Deposit Amount', left_on=['Batch Amount'], indicator='status recon')
        if self.is_capn.get():
            folder_name = 'CP Reconciliation'
            folder_path = os.path.join(parent_dir, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            desktop_dir = Path.home() / "Desktop"
            grouped_batch_merge_filename = f"{dba_name} recon.xlsx"
            grouped_batch_merge_path = os.path.join(folder_path, filename)
            grouped_batch_merge.to_excel(grouped_batch_merge_path, index=False)
            self.file_label.config(text=f"Results saved to {grouped_batch_merge_filename}")
        else:
            # Merge deposit_df with combined_batch_df on Amount column
            merged_df = pd.merge(combined_batch_df,deposit_df,  how='outer',right_on='Deposit Amount', left_on=['Batch Amount'], indicator='status recon')
            merged_df['status recon'] = merged_df['status recon'].map(results).astype(str)
            consecutive['status recon'] = consecutive['status recon'].map(results).astype(str)

            # Do it after formating values as str
            matched_dates = grouped_batch_merge.loc[grouped_batch_merge['status recon'] == 'Match', 'Batch Date']
            consecutive_match = consecutive.loc[consecutive['status recon'] == 'Match', 'Next Day']

            # Find matched dates in dataframe and exclude 
            merged_df = merged_df[~merged_df['Batch Date'].isin(matched_dates)]
            merged_df = merged_df[~merged_df['Batch Date'].isin(matched_deposits_df)]
            # Sort the merged DataFrame by date
            merged_df = pd.concat([merged_df, matched_dates], sort=False).sort_values(by=['Deposit Date', 'Batch Date'], ascending=True,ignore_index=True)
            if len(consecutive_match) > 0:
                consecutive = consecutive[consecutive['status recon'] == 'Match']
                merged_df = pd.merge(merged_df, consecutive, how='left', on='Deposit Amount')
            merged_df = merged_df.drop_duplicates()
            # Save the combined data to an Excel file named after the 'DBA Name' value
            filename = f"{dba_name} recon.xlsx"
            current_dir = os.getcwd()
            parent_dir = self.home_dir
            
            # Create a folder named "duplicates" in the parent directory if it doesn't exist
            folder_name = 'CP Reconciliation'
            folder_path = os.path.join(parent_dir, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            output_file = os.path.join(folder_path, filename)
            merged_df.to_excel(output_file,index=False)

            # Show where it's saved
            self.file_label.config(text=f"Results saved to {filename}")

root = Tk()
reconciler = Reconciler(root)
root.mainloop()

