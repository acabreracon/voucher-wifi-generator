import pdfplumber
import re
from tkinter import filedialog
import tkinter as tk

tk.Tk().withdraw()
ruta = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

total = 0
with pdfplumber.open(ruta) as pdf:
    for i, page in enumerate(pdf.pages):
        texto = page.extract_text()
        codigos = re.findall(r'\b\d{5}-\d{5}\b', texto) if texto else []
        total += len(codigos)
        print(f"Página {i+1}: {len(codigos)} códigos")

print(f"\nTotal general: {total} códigos")