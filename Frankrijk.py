import streamlit as st
import pandas as pd
import PyPDF2
import io
from zipfile import ZipFile
import pdfplumber

def clean_csv(file):
    lines = file.getvalue().decode('utf-8').splitlines()
    cleaned_lines = [line.strip().strip('"') for line in lines]
    return "\n".join(cleaned_lines)
    
def process_files(main_file, reference_file, invoice_files, start_date, end_date):
    # Load and process the main file
    data = pd.read_csv(main_file)
    data = data[data['Fulfillment Status'] != 'restocked']
    data = data.rename(columns={"Lineitem sku": "SKU"})
    data['SKU'] = data['SKU'].str.replace(r"(-\d+|[A-Z])$", "", regex=True)

    # Load and merge reference data
    reference_file_content = clean_csv(reference_file)
    check = pd.read_csv(io.StringIO(reference_file_content))
    data = pd.merge(data, check[['SKU', 'Alcohol Percentage', 'Excise code']], on='SKU', how='left')

    # Data cleaning and filling missing values
    data['Fulfilled at'] = data['Fulfilled at'].ffill()
    data['Billing Name'] = data['Billing Name'].ffill()
    data['Billing Street'] = data['Billing Street'].ffill()

    # Filtering and formatting data
    df = data[data['Shipping Country'] == 'FR']
    selected_columns = ["Name", "Created at", "Fulfilled at", "Lineitem quantity", "Lineitem name", "Billing Name", "Billing Street", "Alcohol Percentage", "Excise code"]
    new_df = df[selected_columns]
    new_df = new_df.rename(columns={"Name": "Invoice/order", "Created at": "Invoice date", "Fulfilled at": "Delivery date","Lineitem name": "Product name", "Lineitem quantity": "Number of sold items", "Billing Name": "Name of client", "Billing Street": "Address details"  })
    new_df['Invoice date'] = pd.to_datetime(new_df['Invoice date'].str.slice(0, 19), errors='coerce')
    new_df['Delivery date'] = pd.to_datetime(new_df['Delivery date'].str.slice(0, 19), errors='coerce')
    new_df["Plato percentage"] = 0
    # Compute 'Content' column
    new_df['Content'] = (new_df['Product name'].str.extract(r'(\d+)(?!.*\d)').astype(float).astype('Int64')) / 10

    # Ensure 'Content' only has integers (removing .0)
    new_df['Content'] = new_df['Content'].fillna(0).astype(int)  # Replace NaN with 0 and cast to integer

    new_df["Total content"] = new_df["Content"]*new_df["Number of sold items"]

    filtered_df = new_df[(new_df['Delivery date'] >= start_date) & (new_df['Delivery date'] <= end_date)]
    final_data = filtered_df[['Invoice/order', 'Invoice date', 'Delivery date', 'Name of client', 'Address details', 'Product name', 'Excise code', 'Number of sold items', 'Content', 'Total content', 'Alcohol Percentage', 'Plato percentage']]
    final_data = final_data.drop_duplicates()

    output_files = []
    not_found_invoices = []

    for invoice_number in final_data['Invoice/order'].unique():
        invoice_found = False
        for uploaded_file in invoice_files:
            with pdfplumber.open(uploaded_file) as pdf:
                for page_number, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text and str(invoice_number) in text:
                        pdf_writer = PyPDF2.PdfWriter()
                        # Reset stream position and open with PdfReader without using 'with'
                        uploaded_file.seek(0)
                        pdf_reader = PyPDF2.PdfReader(uploaded_file)
                        pdf_writer.add_page(pdf_reader.pages[page_number])

                        pdf_bytes = io.BytesIO()
                        pdf_writer.write(pdf_bytes)
                        pdf_bytes.seek(0)
                        output_files.append((f"FR_{invoice_number}.pdf", pdf_bytes))
                        invoice_found = True
                        break
            if invoice_found:
                break
        if not invoice_found:
            not_found_invoices.append(invoice_number)

    return final_data, output_files, not_found_invoices

def create_zip(pdfs, csv_data=None, csv_name=None):
    zip_buffer = io.BytesIO()
    with ZipFile(zip_buffer, 'a') as zf:
        # Add PDF files to the zip
        for filename, filedata in pdfs:
            zf.writestr(filename, filedata.getvalue())
        # Add CSV file to the zip if provided
        if csv_data and csv_name:
            zf.writestr(csv_name, csv_data)
    zip_buffer.seek(0)
    return zip_buffer

# Streamlit interface
st.title("Frankrijk Accijnsaangifte")
uploaded_file = st.file_uploader("Upload het shopify bestand (csv)", type=['csv'])
reference_file = st.file_uploader("Upload connect_csv", type=['csv'])
invoice_file = st.file_uploader("Upload de PDF bestanden met Invoices", type=['pdf'], accept_multiple_files=True)

start_time = st.text_input("Start datum (YYYY-MM-DD HH:MM:SS)")
end_time = st.text_input("Eind datum (YYYY-MM-DD HH:MM:SS)")

if st.button("Process Files"):
    if uploaded_file and reference_file and invoice_file and start_time and end_time:
        start_date = pd.to_datetime(start_time)
        end_date = pd.to_datetime(end_time)
        try:
            processed_data, pdfs, errors = process_files(uploaded_file, reference_file, invoice_file, start_date, end_date)
            
            # Convert processed data to CSV
            csv = processed_data.to_csv(index=False, sep=';').encode('utf-8')
            formatted_start_time = start_date.strftime('%Y%m%d')
            formatted_end_time = end_date.strftime('%Y%m%d')
            csv_name = f"FR_VINIOWIJNIMPORT_{formatted_start_time}_to_{formatted_end_time}.csv"

            # Create ZIP including both PDFs and CSV
            if pdfs or csv:
                zip_buffer = create_zip(pdfs, csv, csv_name)
                st.download_button("Download alle bestanden als ZIP", zip_buffer, "Frankrijk.zip", "application/zip")
            
            if errors:
                st.error("The following invoices were not found in the PDF: " + ", ".join(str(inv) for inv in errors))
        
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.error("Please upload all files and specify the date range.")
