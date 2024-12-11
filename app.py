import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import os
import zipfile

def generate_certificates(file_path, output_folder, achievement):
    try:
        data = pd.read_excel(file_path)

        if "Name" not in data.columns or "Surname" not in data.columns:
            raise ValueError("Excel file must contain 'Name' and 'Surname' columns.")

        os.makedirs(output_folder, exist_ok=True)

        pdfs = []
        for _, row in data.iterrows():
            full_name = f"{row['Name']} {row['Surname']}"
            pdf = FPDF()
            pdf.add_page()

            pdf.set_fill_color(230, 230, 250) 
            pdf.rect(5, 5, 200, 287, 'F') 
            pdf.set_line_width(1.5)
            pdf.set_draw_color(100, 149, 237)
            pdf.rect(10, 10, 190, 277)

            # Add title
            pdf.set_font("Arial", style="B", size=30)
            pdf.set_text_color(65, 105, 225)
            pdf.ln(40)
            pdf.cell(0, 20, "Certificate of Achievement", ln=True, align="C")

            pdf.set_font("Times", style="B", size=24)
            pdf.set_text_color(0, 0, 0) 
            pdf.ln(20)
            pdf.cell(0, 10, f"Presented to: {full_name}", ln=True, align="C")

            pdf.set_font("Arial", size=18)
            pdf.ln(15)
            pdf.cell(0, 10, f"For: {achievement}", ln=True, align="C")

            pdf.set_draw_color(192, 192, 192) 
            pdf.set_line_width(0.5)
            pdf.line(60, 220, 150, 220)

            pdf.set_font("Arial", size=12)
            pdf.set_text_color(105, 105, 105) 
            pdf.ln(60)
            pdf.cell(0, 10, "Congratulations on your achievement!", ln=True, align="C")

            pdf_file = os.path.join(output_folder, f"{full_name}.pdf")
            pdf.output(pdf_file)
            pdfs.append(pdf_file)

        zip_file = os.path.join(output_folder, "Certificates.zip")
        with zipfile.ZipFile(zip_file, "w") as zipf:
            for pdf in pdfs:
                zipf.write(pdf, os.path.basename(pdf))

        return zip_file

    except Exception as e:
        messagebox.showerror("Error", str(e))
        return None

class CertificateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Generator")
        self.root.geometry("400x300")

        tk.Label(root, text="Upload Excel File", font=("Arial", 12)).pack(pady=10)
        self.file_label = tk.Label(root, text="No file selected", fg="grey")
        self.file_label.pack(pady=5)

        tk.Button(root, text="Browse", command=self.browse_file).pack(pady=5)

        tk.Label(root, text="Achievement Name", font=("Arial", 12)).pack(pady=10)
        self.achievement_entry = tk.Entry(root, width=30)
        self.achievement_entry.pack(pady=5)

        tk.Button(root, text="Generate Certificates", command=self.generate).pack(pady=20)

        self.file_path = ""

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.file_path:
            self.file_label.config(text=os.path.basename(self.file_path), fg="black")

    def generate(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        achievement = self.achievement_entry.get().strip()
        if not achievement:
            messagebox.showerror("Error", "Please enter an achievement name.")
            return

        output_folder = filedialog.askdirectory()
        if not output_folder:
            return

        zip_file = generate_certificates(self.file_path, output_folder, achievement)

        if zip_file:
            messagebox.showinfo("Success", f"Certificates generated successfully!\nSaved at: {zip_file}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateApp(root)
    root.mainloop()
