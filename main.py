import pandas as pd
from docx import Document
from copy import deepcopy
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def generate():
    try:
        # Load the Excel file
        data = pd.read_excel(excel_entry.get())

        # Convert 'PO' to numeric and drop rows with non-numeric 'PO'
        data['PO'] = pd.to_numeric(data['PO'], errors='coerce')
        data = data.dropna(subset=['PO'])

        # Only keep rows with the user-specified PO number
        po_number = int(po_entry.get())
        data = data[data['PO'] == po_number]

        # Load the Word documents
        template1 = Document(template1_entry.get())
        template2 = Document(template2_entry.get())

        # Get today's date
        today = datetime.today().strftime('%Y/%m/%d')

        # Group the DataFrame by the 'PO' column
        groups = data.groupby('PO')

        for po, group in groups:
            # Create a copy of the templates for each group
            doc1 = deepcopy(template1)
            doc2 = deepcopy(template2)

            # Combine the quantities, batch numbers, and reasons in the desired format
            damages_reasons = ', '.join(
                f'{int(row["Damage"])} bags {row["Reason"]} ({row["Batch Number"]})'
                for _, row in group.iterrows()
            )

            # Combine the quantities and reasons for the second document
            group2 = group.groupby('Reason').agg({'Damage': 'sum'}).reset_index()
            damages_reasons2 = ', and '.join(
                f'{int(row["Damage"])} bags are found {row["Reason"]}'
                for _, row in group2.iterrows()
            )

            # Define a dictionary that maps placeholders to replacement text
            replacements1 = {
                '<DATE>': today,
                '<PO>': str(int(po)),  # convert 'PO' to int before converting to string to remove decimal point
                '<INVOICE_NO>': ', '.join(set(group['Invoice NO'].astype(int).astype(str))),
                '<PRODUCT_CODE>': ', '.join(set(group['Product code'].astype(str))),
                '<DAMAGE>': damages_reasons,
                '<BATCH_NUMBER>': ' / '.join(group['Batch Number'].astype(str)),
                '<REASON>': ', '.join(set(group['Reason'].astype(str))),
            }

            replacements2 = {
                '<REASON>': damages_reasons2,
                '<COMPLAINT>': ', '.join(set(group['Complaint'].astype(str))),
                '<PO>': str(int(po)),
                '<INVOICE_NUM>': ', '.join(set(group['Invoice NO'].astype(int).astype(str))),
                '<BAG_NUM>': str(int(group['Damage'].sum())),
                '<VALUE>': str(round(group['Good Value'].sum(), 2)),
            }

            # Replace the placeholders in the documents
            for para in doc1.paragraphs:
                for placeholder, replacement_text in replacements1.items():
                    if placeholder in para.text:
                        for run in para.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, replacement_text)

            for para in doc2.paragraphs:
                for placeholder, replacement_text in replacements2.items():
                    if placeholder in para.text:
                        for run in para.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, replacement_text)

            for table in doc2.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for placeholder, replacement_text in replacements2.items():
                                if placeholder in para.text:
                                    for run in para.runs:
                                        if placeholder in run.text:
                                            run.text = run.text.replace(placeholder, replacement_text)

            # Save the changes with a filename based on the 'PO' value
            doc1.save(f'{output_entry.get()}/COD {int(po)}.docx')  # convert 'PO' to int to remove decimal point
            doc2.save(f'{output_entry.get()}/Email {int(po)}.docx')  # convert 'PO' to int to remove decimal point

        # Show a success message and exit the program
        messagebox.showinfo("Success", "Generation completed successfully!")
        root.quit()
    except Exception as e:
        # Show an error message if anything goes wrong
        messagebox.showerror("Error", str(e))

def select_file(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def select_directory(entry):
    directory_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, directory_path)

root = tk.Tk()

excel_entry = tk.Entry(root, width=50)
excel_entry.pack()
excel_button = tk.Button(root, text="Select Excel File", command=lambda: select_file(excel_entry))
excel_button.pack()

template1_entry = tk.Entry(root, width=50)
template1_entry.pack()
template1_button = tk.Button(root, text="Select COD Template File", command=lambda: select_file(template1_entry))
template1_button.pack()

template2_entry = tk.Entry(root, width=50)
template2_entry.pack()
template2_button = tk.Button(root, text="Select Email Template File", command=lambda: select_file(template2_entry))
template2_button.pack()

output_entry = tk.Entry(root, width=50)
output_entry.pack()
output_button = tk.Button(root, text="Select Output Directory", command=lambda: select_directory(output_entry))
output_button.pack()

po_entry = tk.Entry(root, width=50)
po_entry.pack()
po_label = tk.Label(root, text="Enter PO Number")
po_label.pack()

generate_button = tk.Button(root, text="Generate", command=generate)
generate_button.pack()

root.mainloop()
