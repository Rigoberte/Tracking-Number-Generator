from pypdf import PdfReader, PdfWriter
import os

def merge_PDFs(list_of_pdfs: list, folder_path_to_download: str, output_file_name: str):
    """
    Combina una lista de archivos PDF en un solo archivo PDF.

    Args:
        list_of_pdfs (list): Lista de rutas de archivos PDF a combinar.
        folder_path_to_download (str): Carpeta donde se guardará el PDF combinado.
        file_name (str): Nombre del archivo combinado (sin extensión).
    """

    if not list_of_pdfs:
        raise ValueError("La lista de archivos PDF no puede estar vacía.")

    writer = PdfWriter()

    for pdf_path in list_of_pdfs:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)

    output_path = os.path.join(folder_path_to_download, f"{output_file_name}.pdf")
    with open(output_path, "wb") as f_out:
        writer.write(f_out)