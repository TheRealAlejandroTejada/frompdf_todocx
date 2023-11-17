from pdf2docx import Converter
from tkinter import Tk, filedialog

def välj_fil():
    root = Tk()
    root.withdraw()  # Dölj huvudfönstret för tkinter

    # Be användaren att välja en PDF-fil
    filväg = filedialog.askopenfilename(title="Välj PDF-fil",
                                         filetypes=[("PDF-filer", "*.pdf")])

    return filväg

def välj_mål():
    root = Tk()
    root.withdraw()  # Dölj huvudfönstret för tkinter

    # Be användaren att välja plats och namn på utgående DOCX-fil
    filväg = filedialog.asksaveasfilename(title="Spara som",
                                           defaultextension=".docx",
                                           filetypes=[("DOCX-filer", "*.docx")])

    return filväg

def pdf_till_docx(pdf_fil, docx_fil):
    try:
        # Skapa ett Converter-objekt
        cv = Converter(pdf_fil)

        # Konvertera PDF till DOCX
        cv.convert(docx_fil, start=0, end=None)

        # Stäng Converter-objektet
        cv.close()

        print(f'Konvertering lyckades. Filen {docx_fil} har skapats.')
    except Exception as e:
        print(f'Fel under konverteringen: {str(e)}')

if __name__ == "__main__":
    # Be användaren att välja en PDF-fil
    inmatning_pdf = välj_fil()

    # Be användaren att välja plats och namn på utgående DOCX-fil
    utmatning_docx = välj_mål()

    # Anropa konverteringsfunktionen
    pdf_till_docx(inmatning_pdf, utmatning_docx)
