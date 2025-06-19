import fitz 
import os
import re
from zipfile import ZipFile
from collections import Counter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def extract_name(text):
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if len(lines) >= 2:
        return lines[1]  
    return "unknown_name"

def split_certificates(input_pdf, output_dir, save_pdf=True, save_zip=True):
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(input_pdf)
    pdf_paths = []
    name_counts = Counter()

    for i, page in enumerate(doc):
        text = page.get_text("text")
        name = extract_name(text)
        name_counts[name] += 1
        final_name = f"{name}_{name_counts[name]}" if name_counts[name] > 1 else name
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", final_name)
        output_pdf_path = os.path.join(output_dir, f"{safe_name}.pdf")

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

    # Determine what to save
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


root = tk.Tk()
root.title("Certificate Splitter")
root.geometry("500x350")

pdf_path_var = tk.StringVar()
output_folder_var = tk.StringVar()
output_option_var = tk.StringVar(value="Both")  

tk.Label(root, text="1. Select PDF File:", font=("Arial", 12)).pack(pady=5)
tk.Entry(root, textvariable=pdf_path_var, width=60).pack(padx=10)
tk.Button(root, text="Browse PDF", command=browse_pdf).pack(pady=5)

tk.Label(root, text="2. Choose Destination Folder (optional):", font=("Arial", 12)).pack(pady=5)
tk.Entry(root, textvariable=output_folder_var, width=60).pack(padx=10)
tk.Button(root, text="Select Folder", command=choose_destination).pack(pady=5)

tk.Label(root, text="3. Output Format:", font=("Arial", 12)).pack(pady=5)
format_menu = ttk.Combobox(root, textvariable=output_option_var, values=["PDF only", "ZIP only", "Both"], state="readonly", width=20)
format_menu.pack()

tk.Button(root, text="Split Certificates", command=process_pdf, bg="green", fg="white", font=("Arial", 12)).pack(pady=20)

root.mainloop()
