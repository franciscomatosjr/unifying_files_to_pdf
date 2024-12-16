# pip install PyPDF2 fpdf python-docx pandas openpyxl pillow


import os
from PyPDF2 import PdfMerger
from fpdf import FPDF
from docx import Document
import pandas as pd
from PIL import Image


def convert_to_pdf(file_path, output_folder):
    """
    Converte arquivos de diferentes formatos para PDF.
    Suporta: .docx, .txt, .xls, .xlsx
    """
    ext = os.path.splitext(file_path)[1].lower()
    output_file = os.path.join(output_folder, os.path.basename(file_path) + ".pdf")

    if ext == '.docx':
        # Converte documentos Word (.docx) para PDF
        document = Document(file_path)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for paragraph in document.paragraphs:
            pdf.multi_cell(0, 10, paragraph.text)
        pdf.output(output_file)

    elif ext in ['.xls', '.xlsx']:
        # Converte planilhas Excel para PDF
        df = pd.read_excel(file_path)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        for index, row in df.iterrows():
            pdf.cell(0, 10, txt=", ".join(map(str, row.values)), ln=True)
        pdf.output(output_file)

    elif ext == '.txt':
        # Converte arquivos de texto para PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        with open(file_path, "r", encoding="utf-8") as file:
            for line in file:
                pdf.multi_cell(0, 10, line)
        pdf.output(output_file)

    elif ext in ['.jpg', '.jpeg', '.png']:
        # Converte imagens para PDF
        image = Image.open(file_path)
        if image.mode in ("RGBA", "LA"):
            image = image.convert("RGB")
        output_file = os.path.join(output_folder, os.path.splitext(os.path.basename(file_path))[0] + ".pdf")
        image.save(output_file, "PDF")
    
    else:
        raise ValueError(f"Formato de arquivo {ext} não suportado.")

    return output_file


def unify_to_pdf(files, output_pdf):
    """
    Converte vários arquivos para PDF e unifica em um único arquivo PDF.
    
    :param files: Lista de arquivos para unificar.
    :param output_pdf: Caminho do PDF unificado.
    """
    output_folder = "temp_pdfs"
    os.makedirs(output_folder, exist_ok=True)
    merger = PdfMerger()

    for file in files:
        try:
            pdf_path = convert_to_pdf(file, output_folder)
            merger.append(pdf_path)
        except ValueError as e:
            print(f"Erro ao processar arquivo {file}: {e}")

    # Salva o PDF unificado
    merger.write(output_pdf)
    merger.close()

    # Limpa arquivos temporários
    for temp_file in os.listdir(output_folder):
        os.remove(os.path.join(output_folder, temp_file))
    os.rmdir(output_folder)

    print(f"PDF unificado salvo em: {output_pdf}")


# Exemplo de uso
files = [
    "documento.docx",
    "planilha.xlsx",
    "notas.txt",
    "imagem.jpg"
]  # Lista de arquivos para converter e unificar
output_pdf = "unificado.pdf"

unify_to_pdf(files, output_pdf)
