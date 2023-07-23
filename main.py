import pandas as pd
from docx import Document
from copy import deepcopy
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def generate():
    # Load the Excel file
    data = pd.read_excel(excel_entry.get())

    # Convert 'PO' to numeric and drop rows with non-numeric 'PO'
    data['PO'] = pd.to_numeric(data['PO'], errors='coerce')
    data = data.dropna(subset=['PO'])

    # Load the Word document
    template = Document(template_entry.get())

    # Get today's date
    today = datetime.today().strftime('%Y/%m/%d')

    # Group the DataFrame by the 'PO' column
    groups = data.groupby('PO')

    for po, group in groups:
        doc = deepcopy(template)  # create a copy of the template for each group

        # Combine the quantities, batch numbers, and reasons in the desired format
        damages_reasons = ', '.join(
            f'{row["Damage"]} ({row["Batch Number"]}) bag(s) {row["Reason"]}'
            for _, row in group.iterrows()
        )

        # Define a dictionary that maps placeholders to replacement text
        replacements = {
            '<DATE>': today,
            '<PO>': str(int(po)),  # convert 'PO' to int before converting to string to remove decimal point
            '<INVOICE_NO>': ', '.join(set(group['Invoice NO'].astype(str))),
            '<PRODUCT_CODE>': ', '.join(set(group['Product code'].astype(str))),
            '<DAMAGE>': damages_reasons,
            '<BATCH_NUMBER>': ', '.join(group['Batch Number'].astype(str)),
            '<REASON>': ', '.join(set(group['Reason'].astype(str))),
        }

        # Replace the placeholders in the document
        for para in doc.paragraphs:
            for placeholder, replacement_text in replacements.items():
                if placeholder in para.text:
                    for run in para.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, replacement_text)

        # Save the changes with a filename based on the 'PO' value
        doc.save(f'{output_entry.get()}/Filled COD Sample {int(po)}.docx')  # convert 'PO' to int to remove decimal point

    # Show a success message and exit the program
    messagebox.showinfo("Success", "Generation completed successfully!")
    root.quit()

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

template_entry = tk.Entry(root, width=50)
template_entry.pack()
template_button = tk.Button(root, text="Select Template File", command=lambda: select_file(template_entry))
template_button.pack()

output_entry = tk.Entry(root, width=50)
output_entry.pack()
output_button = tk.Button(root, text="Select Output Directory", command=lambda: select_directory(output_entry))
output_button.pack()

generate_button = tk.Button(root, text="Generate", command=generate)
generate_button.pack()

root.mainloop()
