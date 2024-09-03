# This file is main file of Backend which contains all the functionality.

from typing import List
import webbrowser
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
import pandas as pd
from io import BytesIO
import subprocess
import shutil
import os

import pdfplumber

# Initiating fastapi
app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allow all methods
    allow_headers=["*"],  # Allow all headers
)

# Path of chrome browser
chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'

# Register Chrome as the browser
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))

# Url to open in chrome browser
url = "http://127.0.0.1:8000/"
webbrowser.get('chrome').open(url)

# This will display main.html on the screen to choose operation.
@app.get("/", response_class=HTMLResponse)
async def read_root():
    with open("main.html") as f:
        return HTMLResponse(content=f.read(), status_code=200)
    
# This will display automation.html on the screen.
@app.get("/automation", response_class=HTMLResponse)
async def read_root():
    with open("automation.html") as f:
        return HTMLResponse(content=f.read(), status_code=200)
    
# This will display data_merge.html on the screen.
@app.get("/data_merge", response_class=HTMLResponse)
async def read_root():
    with open("data_merge.html") as f:
        return HTMLResponse(content=f.read(), status_code=200)
    
# This will display data_join.html on the screen.    
@app.get("/data_join", response_class=HTMLResponse)
async def read_root():
    with open("data_join.html") as f:
        return HTMLResponse(content=f.read(), status_code=200)
    
# This will display pdf.html on the screen.
@app.get("/pdf", response_class=HTMLResponse)
async def read_root():
    with open("pdf.html") as f:
        return HTMLResponse(content=f.read(), status_code=200)
    
# This will display pdfExcel.html on the screen.
@app.get("/excel", response_class=HTMLResponse)
async def read_root():
    with open("pdfExcel.html") as f:
        return HTMLResponse(content=f.read(), status_code=200)

# Automation script will start once the user clicks the button in automation.html file.
@app.post("/run-script")
async def run_script(file: UploadFile = File(...), lower_limit: int = Form(...), upper_limit: int = Form(...),
                      timeToLoad: int = Form(...), fileName: str = Form(...), tab: str = Form(...)):
    try:
        # Save the uploaded file
        file_location = f"temp_{file.filename}"
        with open(file_location, "wb") as f:
            shutil.copyfileobj(file.file, f)

        # Run the script with the uploaded file and limits
        result = subprocess.run(['python', 'automation.py', file_location, str(lower_limit), str(upper_limit), str(timeToLoad), fileName, tab], capture_output=True, text=True)

        # Remove the uploaded file after processing
        os.remove(file_location)

        return JSONResponse(content={'output': result.stdout, 'error': result.stderr})
    except Exception as e:
        return JSONResponse(content={'error': str(e)})
    
# Data Merge will start once the user clicks the button in data_merge.html file.
@app.post("/merge")
async def merge_files(files: list[UploadFile]):
    data_frames = []

    # Read each Excel file into a DataFrame and store it
    for file in files:
        contents = await file.read()
        df = pd.read_excel(BytesIO(contents))
        data_frames.append(df)

    # Concatenate all DataFrames row-wise
    merged_df = pd.concat(data_frames, ignore_index=True)

    # Save the merged DataFrame to an Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False)

    output.seek(0)  # Move the cursor back to the beginning of the file

    # Return the file as a StreamingResponse
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=merged.xlsx"}
    )

# Data Join will start once the user clicks the button in data_join.html file.
@app.post("/join")
async def join_files(file1: UploadFile, file2: UploadFile, file3: UploadFile, column: str = Form(...)):
    
    # Checking column whether LCC or BSP
    if column == 'LCC': 

        # Read the first Excel file into a DataFrame
        contents1 = await file1.read()
        lcc_reco_df = pd.read_excel(BytesIO(contents1))

        # Read the second Excel file into a DataFrame
        contents2 = await file2.read()
        airlines_df = pd.read_excel(BytesIO(contents2))

        # Read the third Excel file into a DataFrame
        contents3 = await file3.read()
        travcom_df = pd.read_excel(BytesIO(contents3))

        # Specify the column names you want to select
        airlines_columns = ["TYPE", "RecordLocator", "Sum of AMOUNT", "P Code", "Transation Date", "Airline Name", "BRANCH", "Name1"]

        travcom_columns = ["TYPE", "TKT NO", "FINALAMOUNT", "INVOICE NO", "DOC_DT", "SLMASTER", "CLIENT NAME", "BRANCH", "PAX NAME"]

        lcc_reco_columns = ["TKTT TYPE", "PNR", "TRAVCOM AMOUNT", "AIRLINE AMOUNT", "Difference", "Exception Remarks", "DOCUMENT NO", "DATE", "AIRLINE", "CLIENT", "Branch", "PAX NAME"]

        required_order = ["TYPE", "TKT NO", "TRAVCOM AMOUNT", "AIRLINE AMOUNT", "Difference", "Sum of AMOUNT", "FINALAMOUNT", "P Code", "Exception Remarks", "DOCUMENT NO", "INVOICE NO", "DATE", "Transation Date", "DOC_DT", "AIRLINE", "Airline Name", "SLMASTER", "CLIENT", "CLIENT NAME", "Branch", "BRANCH", "BRANCH", "PAX NAME", "Name1", "PAX NAME"]

        renaming_order = ["TYPE", "TKT NO", "TRAVCOM AMOUNT", "AIRLINE AMOUNT", "Difference", "Sum of AMOUNT", "FINALAMOUNT", "P Code", "Exception Remarks", "DOCUMENT NO", "INVOICE NO", "DATE", "Transation Date", "DOC_DT", "AIRLINE", "Airline Name", "SLMASTER", "CLIENT", "CLIENT NAME", "Branch", "BRANCH_x", "BRANCH_y", "PAX NAME_x", "Name1", "PAX NAME_y"]

        renaming_order_with_prefix = ["TYPE", "TKT NO", "Reco_TRAVCOM AMOUNT", "Reco_AIRLINE AMOUNT", "Reco_Difference", "Air_Sum of AMOUNT", "Tr_FINALAMOUNT", "Air_P Code", "Reco_Exception Remarks", "Reco_DOCUMENT NO", "Tr_INVOICE NO", "Reco_DATE", "Air_Transation Date", "Tr_DOC_DT", "Reco_AIRLINE", "Air_Airline Name", "Tr_SLMASTER", "Reco_CLIENT", "Tr_CLIENT NAME", "Reco_Branch", "Air_BRANCH", "Tr_BRANCH", "Tr_PAX NAME", "Air_Name1", "Reco_PAX NAME"]

        # Extract the specified columns
        selected_airlines_columns = airlines_df[airlines_columns]
        selected_travcom_columns = travcom_df[travcom_columns]
        selected_lcc_reco_columns = lcc_reco_df[lcc_reco_columns]

        # Renaming the lcc reco and airlines column for similar name.
        selected_lcc_reco_columns = selected_lcc_reco_columns.rename(columns={"TKTT TYPE": "TYPE"})
        selected_lcc_reco_columns = selected_lcc_reco_columns.rename(columns={"PNR": "TKT NO"})
        selected_airlines_columns = selected_airlines_columns.rename(columns={"RecordLocator": "TKT NO"})

        # Perform an outer join based on the specified column
        merged_df = pd.merge(selected_lcc_reco_columns, selected_travcom_columns, on=["TYPE", "TKT NO"], how='outer')
        merged_df = pd.merge(merged_df, selected_airlines_columns, on=["TYPE", "TKT NO"], how='outer')

        merged_df = merged_df[renaming_order]

        merged_df.columns = renaming_order_with_prefix

        # Ordering the columns in dataframe
        final_df = merged_df

    else:

        # Read the first Excel file into a DataFrame
        contents1 = await file1.read()
        bsp_reco_df = pd.read_excel(BytesIO(contents1))

        # Read the second Excel file into a DataFrame
        contents2 = await file2.read()
        statement_df = pd.read_excel(BytesIO(contents2))

        # Read the third Excel file into a DataFrame
        contents3 = await file3.read()
        travcom_df = pd.read_excel(BytesIO(contents3))

        # Renaming the bsp reco column for similar name.
        bsp_reco_df = bsp_reco_df.rename(columns={"TICKET NO": "Ticket No"})

        # Specify the column names you want to select
        statement_columns = ["TKTT TYPE", "Ticket No", "Sum of Gross Amount", "BSP FOP", "Airline Name", "Agent (incl Check Digit)", "Agent IATA Region", "Type Group", "RA NO", "Date of Issue", "Credit Card Number (masked)", "Passenger Name", "PNR"]
        travcom_columns = ["TKTT TYPE", "Ticket No", "Gross Amount", "FOP", "PNR NO", "InvoiceNumber", "MainTag", "InvoiceDate", "ProfileName", "ValidatingCarrier", "TicketingAgentName", "IataNumber", "IataName"]
        bsp_reco_columns = ["TKTT TYPE", "Ticket No", "TRAVCOM AMOUNT", "BSP AMOUNT", "Diff", "CLIENT NAME", "AIRLINE CODE", "EXCEPTION REMARKS", "PAX NAME", "FCM FOP", "CART NO", "PNR NO", "BRANCH", "DOCUMENT NO", "DOC_DATE"]
        required_order = ["TKTT TYPE", "Ticket No", "TRAVCOM AMOUNT", "BSP AMOUNT", "Diff", "Sum of Gross Amount", "Gross Amount", "Type Group", "MainTag", "EXCEPTION REMARKS", "DOCUMENT NO", "InvoiceNumber", "InvoiceDate", "Date of Issue", "DOC_DATE", "CART NO", "FCM FOP", "BSP FOP", "FOP", "PNR", "PNR NO_x", "PNR NO_y", "CLIENT NAME", "ProfileName", "AIRLINE CODE", "Airline Name", "ValidatingCarrier", "BRANCH", "Agent IATA Region", "IataName", "PAX NAME", "Passenger Name", "TicketingAgentName"]
        renaming_order = ["TKTT TYPE", "Ticket No", "TRAVCOM AMOUNT", "BSP AMOUNT", "Diff", "Sum of Gross Amount", "Gross Amount", "Type Group", "MainTag", "EXCEPTION REMARKS", "DOCUMENT NO", "InvoiceNumber", "InvoiceDate", "Date of Issue", "DOC_DATE", "CART NO", "FCM FOP", "BSP FOP", "FOP", "PNR", "PNR NO_x", "PNR NO_y", "CLIENT NAME", "ProfileName", "AIRLINE CODE", "Airline Name", "ValidatingCarrier", "BRANCH", "Agent IATA Region", "IataName", "PAX NAME", "Passenger Name", "TicketingAgentName"]
        renaming_order_with_prefix = ["TKTT TYPE", "Ticket No", "Reco_TRAVCOM AMOUNT", "Reco_BSP AMOUNT", "Reco_Diff", "St_Sum of Gross Amount", "Tr_Gross Amount", "St_Type Group", "Tr_MainTag", "Reco_EXCEPTION REMARKS", "Reco_DOCUMENT NO", "Tr_InvoiceNumber", "Tr_InvoiceDate", "St_Date of Issue", "Reco_DOC_DATE", "Reco_CART NO", "Reco_FCM FOP", "St_BSP FOP", "Tr_FOP", "Reco_PNR", "St_PNR NO", "Tr_PNR NO", "Reco_CLIENT NAME", "Tr_ProfileName", "Reco_AIRLINE CODE", "St_Airline Name", "Tr_ValidatingCarrier", "Reco_BRANCH", "St_Agent IATA Region", "Tr_IataName", "Reco_PAX NAME", "St_Passenger Name", "Tr_TicketingAgentName"]

        # Extract the specified columns
        selected_statement_columns = statement_df[statement_columns]
        selected_travcom_columns = travcom_df[travcom_columns]
        selected_bsp_reco_columns = bsp_reco_df[bsp_reco_columns]

        # Perform an outer join based on the specified column
        merged_df = pd.merge(selected_bsp_reco_columns, selected_travcom_columns, on=["TKTT TYPE", "Ticket No"], how='outer')
        merged_df = pd.merge(merged_df, selected_statement_columns, on=["TKTT TYPE", "Ticket No"], how='outer')

        # Ordering the columns in dataframe
        merged_df = merged_df[renaming_order]

        merged_df.columns = renaming_order_with_prefix

        # Ordering the columns in dataframe
        final_df = merged_df

    # Save the merged DataFrame to an Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)

    output.seek(0)  # Move the cursor back to the beginning of the file

    # Return the file as a StreamingResponse
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=final.xlsx"}
    )

# Pdf Download script will start once the user clicks the button in pdf.html file.
@app.post("/download_pdf")
async def run_script(file: UploadFile = File(...), lower_limit: int = Form(...), upper_limit: int = Form(...),
                      timeToLoad: int = Form(...), tab: str = Form(...)):
    try:

        # Save the uploaded file
        file_location = f"temp_{file.filename}"
        with open(file_location, "wb") as f:
            shutil.copyfileobj(file.file, f)

        # Selecting file name based on tab
        fileName = tab + '.py'

        # Run the script with the uploaded file and limits
        result = subprocess.run(['python', fileName, file_location, str(lower_limit), str(upper_limit), str(timeToLoad)], capture_output=True, text=True)

        # Remove the uploaded file after processing
        os.remove(file_location)

        return JSONResponse(content={'output': result.stdout, 'error': result.stderr})
    
    except Exception as e:

        return JSONResponse(content={'error': str(e)})
    
@app.post("/pdf_to_excel")
async def pdf_to_excel(files: List[UploadFile] = File(...), fileName: str = Form(...)):
    
    try:
        required_list = ['Tax Invoice Number', 'Tax Invoice Date', 'Credit Note Number', 'Credit Note Date', 'Client Name', 'Cart Ref', 'Airline PNR', 'Orig Inv#', 'Orig Inv Date', 'Pax Name', 'From Ticket No', 'Total Fare', 'Add: Meal/Seat/Bag Charge', 'Gross Fare', 'Add: Service Charge', 'Add: Financial Charge', 'Total Charges', 'Less: Trade Discount', 'Add: GST Tax', 'Total', 'Form of Payment', 'Mail ID']
        
        # Currently available, 3 missing
        excel_list = ['Tax Invoice Number', 'Tax Invoice Date', 'Credit Note Number', 'Credit Note Date', 'Cart Ref', 'Airline PNR', 'Orig Inv#', 'Orig Inv Date', 'Total Fare', 'Add: Meal/Seat/Bag Charge', 'Gross Fare', 'Add: Service Charge', 'Add: Financial Charge', 'Total Charges', 'Less: Trade Discount', 'Add: GST Tax', 'Total', 'Form of Payment', 'Mail ID']
        total_columns = len(excel_list)
        def find_next_value_using_pair(pair):
            found_pairs = {}
            found_pairs[pair] = 'Not found'
            for i in range(len(lists) - 1):
                if len(pair) == 3 and lists[i] == pair[0] and lists[i + 1] == pair[1] and lists[i + 2] == pair[2]:
                    if lists[i+3] == 'Booked':
                        found_pairs[pair] = 'Not found'
                    else:
                        found_pairs[pair] = lists[i+3]
                    break
                if len(pair) == 2 and lists[i] == pair[0] and lists[i + 1] == pair[1]:
                    found_pairs[pair] = lists[i+2]
                    break
                if lists[i] == pair:
                    found_pairs[pair] = lists[i+1]
                    break
            # print(found_pairs)
            excel_list.append(found_pairs[pair])
            return None

        def find_next_value_using_pair_with_condition(pair):
            found_pairs = {}
            found_pairs[pair] = 'Not found'
            for i in range(len(lists) - 1):
                if len(pair) == 2 and pair[0] == 'Airline' and pair[1] == 'PNR' and lists[i] == pair[0] and lists[i + 1] == pair[1] and lists[i-2] == 'Ref':
                    found_pairs[pair] = lists[i+2]
                    break
            # print(found_pairs)
            excel_list.append(found_pairs[pair])
            return None


        for file in files:
        # Open the PDF and extract text
            try:
                with pdfplumber.open(file.file) as pdf:
                    text = ""
                    for page in pdf.pages:
                        text += page.extract_text()
                    data = [line.split() for line in text.split('\n') if line.strip() != '']
                    lists = [item for sublist in data for item in sublist]
            except:
                continue

            # Let's find values for both Tax Invoice and Credit Note
            tax_invoice_number = find_next_value_using_pair(('Invoice', 'Number', ':'))
            tax_invoice_date = find_next_value_using_pair(('Invoice', 'Date', ':'))
            credit_note_number = find_next_value_using_pair(('Credit', 'Note:'))
            credit_note_date = find_next_value_using_pair(('Note', ':'))
            client_name = ''
            cart_ref = find_next_value_using_pair(('Cart', 'Ref'))
            pnr = find_next_value_using_pair_with_condition(('Airline', 'PNR'))
            original_invoice = find_next_value_using_pair(('Orig', 'Inv#'))
            original_invoice_date = find_next_value_using_pair(('Orig', 'Inv', 'Date'))
            pax_name = ''
            from_ticket_no = ''
            total_fare = find_next_value_using_pair(('Total', 'Fare:'))
            add_meal_seat_bag_charge = find_next_value_using_pair(('Add:Meal/Seat/Bag', 'Charge:'))
            gross_fare = find_next_value_using_pair(('Gross', 'Fare:'))
            add_service_charge = find_next_value_using_pair(('Service', 'Charge:'))
            add_financial_charge = find_next_value_using_pair(('Financial', 'Charge:'))
            total_charges = find_next_value_using_pair(('Total', 'Charges:'))
            less_trade_discount = find_next_value_using_pair(('Trade', 'Discount:'))
            add_gst_tax = find_next_value_using_pair(('Add:', 'GST', 'Tax'))
            total = find_next_value_using_pair(('Total:'))
            form_of_payment = find_next_value_using_pair(('of', 'Payment', ':'))
            mail_id = find_next_value_using_pair(('Issued', 'By', ':'))

        # print(excel_list)
        # Create an Excel workbook and sheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Extracted Data"

        # Using limit, setting columns and adding the extracted data to sheet
        for i in range(0,len(excel_list), total_columns):
            sheet.append(excel_list[i: i+total_columns])

        # Specify the directory where you want to save the file
        directory = 'C:\\Users\\ganesh.ss\\Downloads'
        # directory = 'C:\\Users\\nithe\\Downloads'

        # Construct the full file path
        file_path = os.path.join(directory, fileName)

        # Increment the counter if the file already exists
        base_filename = file_path

        # Checking for extension
        if not base_filename.endswith('.xlsx'):
            base_filename += '.xlsx'

        # Writing checking filename and renaming filename
        counter = 0
        while os.path.exists(base_filename):
            counter += 1
            if not fileName.endswith('.xlsx'):
                fileName += '.xlsx'
            temp = fileName.split('.')
            base_filename = f"{temp[0]}_{counter}.{temp[1]}"
            base_filename = os.path.join(directory, base_filename)

        # Save the Excel file
        workbook.save(base_filename)
        return JSONResponse(content={'output': 'successfully extracted'})
    
    except Exception as e:

        return JSONResponse(content={'error': str(e)})

# Starting point of the script.
if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
