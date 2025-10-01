#!/usr/bin/env python3
"""
Script per processare documenti: copia e converte in PDF
Supporta file .doc, .docx, .odt e .pdf

Autore: Assistant
Data: 30 Settembre 2025
"""

import os
import shutil
import sys
from pathlib import Path
from typing import List, Tuple

# Importazioni per la conversione
try:
    from docx2pdf import convert as docx_to_pdf
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import comtypes.client
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False

try:
    from odf import text, teletype
    from odf.opendocument import load
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    ODT_AVAILABLE = True
except ImportError:
    ODT_AVAILABLE = False

try:
    import PyPDF2
    PDF_COMPRESSION_AVAILABLE = True
except ImportError:
    PDF_COMPRESSION_AVAILABLE = False

try:
    from PIL import Image
    import fitz  # PyMuPDF per compressione avanzata
    ADVANCED_PDF_COMPRESSION = True
except ImportError:
    ADVANCED_PDF_COMPRESSION = False

# Cartella di destinazione predefinita (dinamica sul desktop dell'utente)
import os
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
DEFAULT_OUTPUT_FOLDER = os.path.join(desktop_path, "files")

# Dimensione massima file prima della compressione (40MB in bytes)
MAX_FILE_SIZE = 40 * 1024 * 1024  # 40MB

def check_dependencies():
    """Verifica le dipendenze necessarie"""
    missing = []
    
    if not DOCX_AVAILABLE:
        missing.append("docx2pdf (pip install docx2pdf)")
    
    if not WORD_AVAILABLE:
        missing.append("comtypes (pip install comtypes)")
    
    if not ODT_AVAILABLE:
        missing.append("odfpy e reportlab (pip install odfpy reportlab)")
    
    if not PDF_COMPRESSION_AVAILABLE:
        missing.append("PyPDF2 (pip install PyPDF2)")
    
    if not ADVANCED_PDF_COMPRESSION:
        missing.append("PyMuPDF e Pillow per compressione avanzata (pip install PyMuPDF Pillow)")
    
    if missing:
        print("⚠️  Dipendenze mancanti:")
        for dep in missing:
            print(f"   - {dep}")
        print("\nInstalla le dipendenze mancanti per utilizzare tutte le funzionalità.")
        print("Alcune conversioni potrebbero non funzionare.")
        return False
    
    return True

def create_output_folder(output_path: str) -> bool:
    """Crea la cartella di output se non esiste"""
    try:
        Path(output_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        print(f"❌ Errore nella creazione della cartella {output_path}: {e}")
        return False

def get_file_size_mb(file_path: str) -> float:
    """Restituisce la dimensione del file in MB"""
    try:
        size_bytes = os.path.getsize(file_path)
        size_mb = size_bytes / (1024 * 1024)
        return size_mb
    except Exception:
        return 0

def compress_pdf_advanced(pdf_path: str) -> Tuple[bool, str]:
    """Comprime un PDF usando PyMuPDF per riduzione più aggressiva"""
    try:
        if not ADVANCED_PDF_COMPRESSION:
            return False, pdf_path
        
        original_size = get_file_size_mb(pdf_path)
        temp_path = pdf_path.replace('.pdf', '_temp_advanced.pdf')
        
        # Apri il PDF con PyMuPDF
        doc = fitz.open(pdf_path)
        
        # Crea un nuovo PDF con compressione
        new_doc = fitz.open()
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Riduci la qualità delle immagini
            pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8))  # Riduci risoluzione all'80%
            img_data = pix.tobytes("jpeg", jpg_quality=70)  # Qualità JPEG al 70%
            
            # Crea una nuova pagina con l'immagine compressa
            new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img_data)
        
        # Salva il PDF compresso
        new_doc.save(temp_path, garbage=4, deflate=True, clean=True)
        new_doc.close()
        doc.close()
        
        if os.path.exists(temp_path):
            compressed_size = get_file_size_mb(temp_path)
            
            if compressed_size < original_size * 0.9:  # Almeno 10% di riduzione
                compression_ratio = (1 - compressed_size / original_size) * 100
                
                # Sostituisci il file originale
                os.remove(pdf_path)
                os.rename(temp_path, pdf_path)
                
                print(f"📉 PDF compresso (avanzato): {Path(pdf_path).name}")
                print(f"   📏 Dimensione originale: {original_size:.1f} MB")
                print(f"   📦 Dimensione compressa: {compressed_size:.1f} MB")
                print(f"   💾 Riduzione: {compression_ratio:.1f}%")
                print(f"   ⚠️  Nota: Qualità immagini ridotta per compressione")
                
                return True, pdf_path
            else:
                os.remove(temp_path)
                return False, pdf_path
        
        return False, pdf_path
    
    except Exception as e:
        print(f"❌ Errore nella compressione avanzata di {pdf_path}: {e}")
        # Pulisci file temporaneo
        temp_path = pdf_path.replace('.pdf', '_temp_advanced.pdf')
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False, pdf_path

def compress_pdf(pdf_path: str) -> Tuple[bool, str]:
    """Ottimizza un file PDF riducendo la sua dimensione quando possibile
    Nota: Con PyPDF2 3.0.1 la compressione è limitata, principalmente riscrittura ottimizzata"""
    try:
        if not PDF_COMPRESSION_AVAILABLE:
            print("⚠️  PyPDF2 non disponibile, compressione PDF saltata")
            return False, pdf_path
        
        if not pdf_path.lower().endswith('.pdf'):
            return False, pdf_path  # Non è un PDF
        
        original_size = get_file_size_mb(pdf_path)
        
        # Crea nome file temporaneo per il PDF compresso
        temp_path = pdf_path.replace('.pdf', '_temp_compressed.pdf')
        
        # Leggi il PDF originale
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            writer = PyPDF2.PdfWriter()
            
            # Copia tutte le pagine
            for page in reader.pages:
                # Comprimi il contenuto della pagina se possibile
                try:
                    page.compress_content_streams()
                except AttributeError:
                    # Il metodo non esiste in questa versione di PyPDF2
                    pass
                
                writer.add_page(page)
            
            # Scrivi il PDF compresso
            with open(temp_path, 'wb') as output_file:
                writer.write(output_file)
        
        # Verifica che la compressione sia riuscita
        if os.path.exists(temp_path):
            compressed_size = get_file_size_mb(temp_path)
            
            # Solo se la compressione ha effettivamente ridotto la dimensione
            if compressed_size < original_size * 0.999:  # Anche solo 0.1% di riduzione
                compression_ratio = (1 - compressed_size / original_size) * 100
                
                # Sostituisci il file originale
                os.remove(pdf_path)
                os.rename(temp_path, pdf_path)
                
                print(f"📉 PDF ottimizzato: {Path(pdf_path).name}")
                print(f"   📏 Dimensione originale: {original_size:.1f} MB")
                print(f"   📦 Dimensione ottimizzata: {compressed_size:.1f} MB")
                print(f"   💾 Riduzione: {compression_ratio:.2f}%")
                
                return True, pdf_path
            else:
                # La compressione non ha portato benefici
                os.remove(temp_path)
                print(f"⚠️  PDF non ottimizzabile con metodi base: {Path(pdf_path).name} ({original_size:.1f} MB)")
                print(f"   💡 Suggerimento: Usa software specializzato per PDF con molte immagini")
                return False, pdf_path
        
        return False, pdf_path
    
    except Exception as e:
        print(f"❌ Errore nella compressione PDF di {pdf_path}: {e}")
        # Pulisci file temporaneo se esiste
        temp_path = pdf_path.replace('.pdf', '_temp_compressed.pdf')
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False, pdf_path

def should_skip_file(file_path: Path) -> bool:
    """Controlla se un file deve essere scartato"""
    filename_lower = file_path.stem.lower()
    
    # File da scartare (nome senza estensione)
    skip_names = [
        "quarta di copertina",
        "quartadicopertina"
    ]
    
    return filename_lower in skip_names

def find_documents(folder_path: str) -> List[str]:
    """Trova tutti i documenti con estensioni supportate"""
    supported_extensions = {'.doc', '.docx', '.odt', '.pdf'}
    documents = []
    skipped_files = []
    
    try:
        folder = Path(folder_path)
        if not folder.exists():
            print(f"❌ La cartella {folder_path} non esiste!")
            return []
        
        for file_path in folder.rglob('*'):
            if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
                if should_skip_file(file_path):
                    skipped_files.append(file_path.name)
                    print(f"⏭️  Scartato: {file_path.name}")
                else:
                    documents.append(str(file_path))
        
        if skipped_files:
            print(f"📋 File scartati: {len(skipped_files)}")
        
        return documents
    
    except Exception as e:
        print(f"❌ Errore nella ricerca dei documenti: {e}")
        return []

def copy_file_to_destination(file_path: str, output_folder: str) -> Tuple[bool, str, bool]:
    """Copia un file nella cartella di destinazione e lo comprime se necessario
    Restituisce: (successo, percorso_finale, è_stato_compresso)"""
    try:
        source_path = Path(file_path)
        dest_path = Path(output_folder) / source_path.name
        
        # Se il file esiste già, aggiungi un numero
        counter = 1
        original_dest = dest_path
        while dest_path.exists():
            name_parts = original_dest.stem, counter, original_dest.suffix
            dest_path = original_dest.parent / f"{name_parts[0]}_{name_parts[1]}{name_parts[2]}"
            counter += 1
        
        # Copia il file
        shutil.copy2(file_path, dest_path)
        print(f"✅ Copiato: {source_path.name} -> {dest_path.name}")
        
        # Controlla se serve compressione PDF
        file_size = get_file_size_mb(str(dest_path))
        if file_size > (MAX_FILE_SIZE / (1024 * 1024)) and str(dest_path).lower().endswith('.pdf'):
            print(f"⚠️  PDF grande ({file_size:.1f} MB), ottimizzazione in corso...")
            
            # Prova prima compressione base
            compressed, final_path = compress_pdf(str(dest_path))
            
            # Se non ha funzionato, prova compressione avanzata
            if not compressed and ADVANCED_PDF_COMPRESSION:
                print(f"🔄 Tentativo compressione avanzata...")
                compressed, final_path = compress_pdf_advanced(str(dest_path))
            
            return True, str(dest_path), compressed
        
        return True, str(dest_path), False
    
    except Exception as e:
        print(f"❌ Errore nella copia di {file_path}: {e}")
        return False, "", False

def convert_doc_to_pdf_word(doc_path: str, output_folder: str) -> bool:
    """Converte file .doc/.docx in PDF usando Microsoft Word"""
    try:
        if not WORD_AVAILABLE:
            return False
        
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        
        doc_path = os.path.abspath(doc_path)
        doc = word.Documents.Open(doc_path)
        
        pdf_path = Path(output_folder) / f"{Path(doc_path).stem}.pdf"
        doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 = PDF format
        
        doc.Close()
        word.Quit()
        
        print(f"✅ Convertito in PDF: {Path(doc_path).name} -> {pdf_path.name}")
        return True
    
    except Exception as e:
        print(f"❌ Errore nella conversione di {doc_path} con Word: {e}")
        return False

def convert_docx_to_pdf_library(docx_path: str, output_folder: str) -> bool:
    """Converte file .docx in PDF usando docx2pdf"""
    try:
        if not DOCX_AVAILABLE:
            return False
        
        pdf_path = Path(output_folder) / f"{Path(docx_path).stem}.pdf"
        docx_to_pdf(docx_path, str(pdf_path))
        
        print(f"✅ Convertito in PDF: {Path(docx_path).name} -> {pdf_path.name}")
        return True
    
    except Exception as e:
        print(f"❌ Errore nella conversione di {docx_path} con docx2pdf: {e}")
        return False

def convert_odt_to_pdf(odt_path: str, output_folder: str) -> bool:
    """Converte file .odt in PDF (implementazione semplificata)"""
    try:
        if not ODT_AVAILABLE:
            return False
        
        # Questa è una conversione semplificata
        # Per una conversione completa, considera l'uso di LibreOffice via subprocess
        doc = load(odt_path)
        
        pdf_path = Path(output_folder) / f"{Path(odt_path).stem}.pdf"
        
        # Estrai il testo e crealo come PDF semplice
        c = canvas.Canvas(str(pdf_path), pagesize=letter)
        y_position = 750
        
        for paragraph in doc.getElementsByType(text.P):
            text_content = teletype.extractText(paragraph)
            if text_content.strip():
                c.drawString(50, y_position, text_content[:80])  # Limita la lunghezza
                y_position -= 20
                if y_position < 50:
                    c.showPage()
                    y_position = 750
        
        c.save()
        print(f"✅ Convertito in PDF: {Path(odt_path).name} -> {pdf_path.name}")
        return True
    
    except Exception as e:
        print(f"❌ Errore nella conversione di {odt_path}: {e}")
        return False

def process_documents(source_folder: str, output_folder: str = DEFAULT_OUTPUT_FOLDER):
    """Processo principale per gestire tutti i documenti"""
    print(f"🔍 Cercando documenti in: {source_folder}")
    print(f"📁 Cartella di destinazione: {output_folder}")
    print("-" * 60)
    
    # Verifica dipendenze
    check_dependencies()
    print()
    
    # Crea cartella di output
    if not create_output_folder(output_folder):
        return
    
    # Trova documenti
    documents = find_documents(source_folder)
    if not documents:
        print("❌ Nessun documento trovato!")
        return
    
    print(f"📄 Trovati {len(documents)} documenti")
    print("-" * 60)
    
    stats = {"copiati": 0, "convertiti": 0, "pdf_ottimizzati": 0, "errori": 0}
    
    for doc_path in documents:
        file_path = Path(doc_path)
        extension = file_path.suffix.lower()
        
        print(f"\n📋 Processando: {file_path.name}")
        
        # Copia il file originale
        success, copied_path, was_compressed = copy_file_to_destination(doc_path, output_folder)
        if success:
            stats["copiati"] += 1
            if was_compressed:
                stats["pdf_ottimizzati"] += 1
        else:
            stats["errori"] += 1
            continue
        
        # Se non è già PDF, prova a convertire
        if extension != '.pdf':
            converted = False
            
            if extension in ['.doc', '.docx']:
                # Prova prima con Word, poi con docx2pdf per .docx
                if convert_doc_to_pdf_word(copied_path, output_folder):
                    converted = True
                elif extension == '.docx' and convert_docx_to_pdf_library(copied_path, output_folder):
                    converted = True
            
            elif extension == '.odt':
                if convert_odt_to_pdf(copied_path, output_folder):
                    converted = True
            
            if converted:
                stats["convertiti"] += 1
                # Rimuovi il file originale copiato se la conversione è riuscita
                try:
                    os.remove(copied_path)
                    print(f"🗑️  Rimosso file originale: {Path(copied_path).name}")
                except:
                    pass
            else:
                print(f"⚠️  Impossibile convertire {file_path.name}, mantenuto file originale")
    
    # Statistiche finali
    print("\n" + "=" * 60)
    print("📊 RIEPILOGO:")
    print(f"   📄 File processati: {len(documents)}")
    print(f"   ✅ File copiati: {stats['copiati']}")
    print(f"   📉 PDF ottimizzati: {stats['pdf_ottimizzati']}")
    print(f"   🔄 File convertiti in PDF: {stats['convertiti']}")
    print(f"   ❌ Errori: {stats['errori']}")
    print(f"   📁 Cartella destinazione: {output_folder}")

def main():
    """Funzione principale"""
    print("=" * 60)
    print("📄 PROCESSATORE DI DOCUMENTI")
    print("   Supporta: .doc, .docx, .odt, .pdf")
    print("=" * 60)
    
    # Richiedi il percorso della cartella
    while True:
        folder_path = input("\n📂 Incolla il percorso della cartella con i documenti: ").strip()
        folder_path = folder_path.strip('"')  # Rimuovi virgolette se presenti
        
        if not folder_path:
            print("❌ Percorso vuoto!")
            continue
        
        if not os.path.exists(folder_path):
            print(f"❌ Il percorso '{folder_path}' non esiste!")
            retry = input("Vuoi riprovare? (s/n): ").lower()
            if retry != 's':
                return
            continue
        
        break
    
    print(f"\n📁 Cartella di destinazione: {DEFAULT_OUTPUT_FOLDER}")
    print("\n🚀 Avvio processamento...")
    process_documents(folder_path, DEFAULT_OUTPUT_FOLDER)
    
    input("\n✨ Pressione Enter per chiudere...")

if __name__ == "__main__":
    main()