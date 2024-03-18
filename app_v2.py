import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter

def choose_pdf_file():
    pdf_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, pdf_file_path)

def choose_docx_destination():
    docx_destination_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    docx_destination_entry.delete(0, tk.END)
    docx_destination_entry.insert(0, docx_destination_path)

def convert_pdf_to_docx():
    pdf_path = pdf_path_entry.get()
    docx_destination = docx_destination_entry.get()

    if pdf_path and docx_destination:
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_destination, start=0, end=None)
            cv.close()
            result_label.config(text="Conversion successful!", fg="green")
        except Exception as e:
            result_label.config(text=f"Error: {e}", fg="red")
    else:
        result_label.config(text="Please select PDF and destination!", fg="red")

def exit_program():
    root.destroy()

root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("400x400")

label = tk.Label(root, text="PDF to DOCX Converter", font=("Arial", 16))
label.pack()

pdf_path_label = tk.Label(root, text="Select PDF File:")
pdf_path_label.pack()
pdf_path_entry = tk.Entry(root, width=50)
pdf_path_entry.pack()
pdf_path_button = tk.Button(root, text="Browse", command=choose_pdf_file)
pdf_path_button.pack()

docx_destination_label = tk.Label(root, text="Select Destination for DOCX:")
docx_destination_label.pack()
docx_destination_entry = tk.Entry(root, width=50)
docx_destination_entry.pack()
docx_destination_button = tk.Button(root, text="Browse", command=choose_docx_destination)
docx_destination_button.pack()

convert_button = tk.Button(root, text="Convert", command=convert_pdf_to_docx)
convert_button.pack()

result_label = tk.Label(root, text="", fg="black")
result_label.pack()

exit_button = tk.Button(root, text="Exit", command=exit_program)
exit_button.pack()

attribution_label = tk.Label(root, text="This software has been developed by Samec Ugljesa")
attribution_label.pack()

root.mainloop()
