import fitz
from datetime import datetime
import os
import re
from zipfile import ZipFile
from collections import Counter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def split_certificates(input_pdf, output_dir, save_pdf=True, save_zip=True):
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(input_pdf)
    pdf_paths = []
    name_counts = Counter()
    cert_numbers = []
    pdf_filenames = []

    for i, page in enumerate(doc):
        text = page.get_text("text")
       
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        if len(lines) >= 2:
            cert_no = lines[0]
            student_name = lines[1]
        else:
            cert_no = "unknown_cert"
            student_name = "unknown_name"

        cert_numbers.append(cert_no)

        grade = output_grade_var.get()
        if grade == "Other":
            grade = other_grade_variable.get().strip() or "UnknownGrade"

        name_counts[(student_name, cert_no)] += 1
        page_number = i + 1

        
        base_name = f"{student_name}_{cert_no}_{page_number}_{grade}"
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", base_name)
        output_pdf_path = os.path.join(output_dir, f"{safe_name}.pdf")
        pdf_filenames.append(safe_name + ".pdf")

        new_pdf = fitz.open()
        new_pdf.insert_pdf(doc, from_page=i, to_page=i)
        if save_pdf:
            new_pdf.save(output_pdf_path)
        pdf_paths.append(output_pdf_path)

    doc.close()

    
    zip_path = None
    if save_zip:
        zip_path = os.path.join(output_dir, "certificates.zip")
        with ZipFile(zip_path, "w") as zipf:
            for pdf_file in pdf_paths:
                zipf.write(pdf_file, arcname=os.path.basename(pdf_file))

    
    if cert_numbers:
        date_str = datetime.now().strftime("%Y%m%d")
        cert_from = cert_numbers[0]
        cert_to = cert_numbers[-1]
        txt_filename = f"{grade}_{date_str}_{cert_from}-{cert_to}.txt"
        txt_path = os.path.join(output_dir, txt_filename)
        with open(txt_path, "w", encoding="utf-8") as f:
            for fname in pdf_filenames:
                f.write(fname + "\n")

    return pdf_paths, zip_path


def browse_pdf():
    filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if filepath:
        pdf_path_var.set(filepath)

def choose_destination():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)

def process_pdf():
    input_pdf = pdf_path_var.get()
    output_folder = output_folder_var.get()
    save_option = output_option_var.get()

    if not input_pdf:
        messagebox.showerror("Error", "Please select a PDF file.")
        return

    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Downloads")

    save_pdf = save_option in ["PDF only", "Both"]
    save_zip = save_option in ["ZIP only", "Both"]

    try:
        pdfs, zip_path = split_certificates(input_pdf, output_folder, save_pdf, save_zip)
        result_msg = ""
        if save_pdf:
            result_msg += f"PDF files saved in:\n{output_folder}\n\n"
        if save_zip:
            result_msg += f"ZIP created at:\n{zip_path}"
        messagebox.showinfo("Success", result_msg.strip())
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{str(e)}")

def on_grade_selected(event):
    if output_grade_var.get() == "Other":
        other_grade_entry.config(state="normal")
    else:
        other_grade_entry.delete(0, tk.END)
        other_grade_entry.config(state="disabled")


root = tk.Tk()
root.title("Certificate Splitter")
root.geometry("600x600")
root.resizable(False, False)
root.configure(bg="#f0f0f0")


pdf_path_var = tk.StringVar()
output_folder_var = tk.StringVar()
output_option_var = tk.StringVar(value="Both")
output_grade_var = tk.StringVar()
other_grade_variable = tk.StringVar()


main_frame = tk.Frame(root, bg="#f0f0f0", padx=20, pady=20)
main_frame.pack(fill="both", expand=True)

title = tk.Label(main_frame, text="American Tamil Academy", font=("Helvetica", 16, "bold"), bg="#f0f0f0")
title.pack(pady=(0, 15))


tk.Label(main_frame, text="1. Select PDF File:", font=("Arial", 12), bg="#f0f0f0").pack(anchor="w")
file_frame = tk.Frame(main_frame, bg="#f0f0f0")
file_frame.pack(fill="x", pady=5)
tk.Entry(file_frame, textvariable=pdf_path_var).pack(side="left", expand=True, fill="x", padx=(0, 5))
tk.Button(file_frame, text="Browse", width=10, command=browse_pdf).pack(side="right")


tk.Label(main_frame, text="2. Choose Destination Folder (optional):", font=("Arial", 12), bg="#f0f0f0").pack(anchor="w", pady=(10, 0))
folder_frame = tk.Frame(main_frame, bg="#f0f0f0")
folder_frame.pack(fill="x", pady=5)
tk.Entry(folder_frame, textvariable=output_folder_var).pack(side="left", expand=True, fill="x", padx=(0, 5))
tk.Button(folder_frame, text="Select", width=10, command=choose_destination).pack(side="right")


tk.Label(main_frame, text="3. Grade:", font=("Arial", 12), bg="#f0f0f0").pack(anchor="w", pady=(10, 0))
menu = ttk.Combobox(main_frame, textvariable=output_grade_var,
                    values=["MUN MAZHALAI", "MAZHALAI", "Nilai 1", "Nilai 2", "Nilai 3",
                            "Nilai 4", "Nilai 5", "Nilai 6", "Nilai 7", "Nilai 8", "Other"],
                    state="readonly", width=57)
menu.pack(pady=5)
menu.bind("<<ComboboxSelected>>", on_grade_selected)


other_grade_label = tk.Label(main_frame, text="If 'Other', specify grade:", font=("Arial", 12), bg="#f0f0f0")
other_grade_label.pack(anchor="w", pady=(5, 0))
other_grade_entry = tk.Entry(main_frame, textvariable=other_grade_variable, width=60, state="disabled")
other_grade_entry.pack(pady=(0, 10))


tk.Label(main_frame, text="4. Output Format:", font=("Arial", 12), bg="#f0f0f0").pack(anchor="w", pady=(10, 0))
format_menu = ttk.Combobox(main_frame, textvariable=output_option_var,
                           values=["PDF only", "ZIP only", "Both"],
                           state="readonly", width=20)
format_menu.pack(pady=5)


tk.Button(main_frame, text="Split Certificates", command=process_pdf,
          bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
          padx=20, pady=10).pack(pady=25)

tk.Label(main_frame, text="© American Tamil Academy", font=("Arial", 9), bg="#f0f0f0", fg="gray").pack(side="bottom", pady=10)

root.mainloop()
