# 📄 Document Processor

Un potente script Python per processar📋 Processando: manuale_   📉 PDF ottimizzati: 1rande.pdf
✅ Copiato: manuale_grande.pdf -> manuale_grande.pdf
⚠️  PDF grande (62.0 MB), ottimizzazione in corso...
⚠️  PDF non ottimizzabile con metodi base: manuale_grande.pdf (62.0 MB)
🔄 Tentativo compressione avanzata...
📉 PDF compresso (avanzato): manuale_grande.pdf
   📏 Dimensione originale: 62.0 MB
   📦 Dimensione compressa: 35.4 MB
   💾 Riduzione: 42.9%
   ⚠️  Nota: Qualità immagini ridotta per compressionevertire documenti in formato PDF.

## 🎯 Funzionalità

- **Scansione automatica**: Trova tutti i documenti supportati in una cartella
- **Filtro intelligente**: Scarta automaticamente i file "Quarta di copertina"
- **Compressione automatica**: Comprime automaticamente i PDF oltre 40MB
- **Copia intelligente**: Copia i file in una cartella di destinazione con gestione duplicati
- **Conversione automatica**: Converte automaticamente file non-PDF in formato PDF
- **Formati supportati**: `.doc`, `.docx`, `.odt`, `.pdf`
- **Interfaccia user-friendly**: Interfaccia a riga di comando semplice e intuitiva

## 🚀 Installazione Rapida

### Metodo 1: Setup Automatico (Consigliato)
```cmd
install_dependencies.bat
```

### Metodo 2: Setup Manuale
```cmd
pip install -r requirements.txt
```

## 🎮 Come Usare

1. **Esegui lo script**:
   ```cmd
   python main.py
   ```

2. **Incolla il percorso** della cartella contenente i documenti

3. **Lascia che lo script faccia tutto!** 🎉
   - La destinazione è sempre: `C:\Users\Youcanprint1\Desktop\files`

## 📋 Esempio di Utilizzo

```
📄 PROCESSATORE DI DOCUMENTI
   Supporta: .doc, .docx, .odt, .pdf
============================================================

📂 Incolla il percorso della cartella con i documenti: C:\MieiDocumenti

📁 Cartella di destinazione: C:\Users\Youcanprint1\Desktop\files

🚀 Avvio processamento...
🔍 Cercando documenti in: C:\MieiDocumenti
📁 Cartella di destinazione: C:\Users\Youcanprint1\Desktop\files
------------------------------------------------------------
✅ Tutte le dipendenze sono disponibili
⏭️  Scartato: Quarta di copertina.pdf

📄 Trovati 4 documenti
📋 File scartati: 1
------------------------------------------------------------

📋 Processando: relazione.docx
✅ Copiato: relazione.docx -> relazione.docx
✅ Convertito in PDF: relazione.docx -> relazione.pdf
🗑️  Rimosso file originale: relazione.docx

📋 Processando: manuale_grande.pdf
✅ Copiato: manuale_grande.pdf -> manuale_grande.pdf
⚠️  PDF grande (45.2 MB), compressione in corso...
� PDF compresso: manuale_grande.pdf
   📏 Dimensione originale: 45.2 MB
   📦 Dimensione compressa: 28.1 MB
   💾 Riduzione: 37.8%

📋 Processando: presentazione.pdf
✅ Copiato: presentazione.pdf -> presentazione.pdf

============================================================
📊 RIEPILOGO:
   📄 File processati: 4
   ✅ File copiati: 4
   � PDF compressi: 1
   🔄 File convertiti in PDF: 1
   ❌ Errori: 0
   📁 Cartella destinazione: C:\Users\Youcanprint1\Desktop\files
```

## 🔧 Requisiti di Sistema

- **Python**: 3.8 o superiore
- **Sistema Operativo**: Windows (ottimizzato per Windows)
- **Microsoft Word**: Opzionale (per migliore conversione di file .doc/.docx)

## 📦 Dipendenze

Lo script usa le seguenti librerie Python:

- `docx2pdf` - Conversione DOCX → PDF
- `comtypes` - Integrazione con Microsoft Word
- `odfpy` - Gestione file ODT (OpenDocument)
- `reportlab` - Generazione PDF
- `python-docx` - Manipolazione documenti Word
- `PyPDF2` - Compressione PDF base
- `PyMuPDF` - Compressione PDF avanzata
- `Pillow` - Elaborazione immagini per compressione

## 🎛️ Configurazione

### Cartella di Destinazione Predefinita
La cartella di destinazione è sempre:

```
C:\Users\Youcanprint1\Desktop\files
```

Se vuoi cambiarla, modifica nel file `main.py`:

```python
DEFAULT_OUTPUT_FOLDER = r"C:\TuaCartella"
```

### Formati Supportati
Per aggiungere nuovi formati, modifica la lista nel file:

```python
supported_extensions = {'.doc', '.docx', '.odt', '.pdf', '.nuovo_formato'}
```

## 🔍 Funzionalità Avanzate

### Gestione Duplicati
- I file con nomi duplicati vengono automaticamente rinominati
- Esempio: `documento.pdf` → `documento_1.pdf`

### Compressione Automatica PDF
- **Soglia**: PDF oltre 40MB vengono automaticamente compressi
- **Livello 1**: PyPDF2 - Ottimizzazione struttura (riduzione 0.1-5%)
- **Livello 2**: PyMuPDF - Compressione immagini avanzata (riduzione 10-50%)
- **Strategia**: Prova prima ottimizzazione base, poi compressione avanzata
- **Trade-off**: Livello 2 riduce qualità immagini (80% risoluzione, 70% qualità JPEG)
- **Feedback**: Mostra dimensioni originali, compresse e percentuale di riduzione
- **Conservazione**: Il file rimane sempre un PDF

### Filtri Automatici
- **File esclusi**: I file chiamati "Quarta di copertina" (in qualsiasi combinazione di maiuscole/minuscole) vengono automaticamente scartati
- **Varianti supportate**: 
  - `Quarta di copertina.pdf` ❌
  - `QUARTA DI COPERTINA.docx` ❌  
  - `quartadicopertina.odt` ❌
  - `QuartaDiCopertina.doc` ❌

### Conversione Multipla
- **File .doc/.docx**: Usa Microsoft Word (se disponibile) o libreria docx2pdf
- **File .odt**: Conversione tramite estrazione testo e generazione PDF
- **File .pdf**: Copia diretta (nessuna conversione)

### Gestione Errori
- Continua il processamento anche in caso di errori su singoli file
- Fornisce statistiche dettagliate al termine
- Log completi di tutte le operazioni

## 🛠️ Risoluzione Problemi

### "Dipendenze mancanti"
```cmd
pip install -r requirements.txt
```

### "Microsoft Word non trovato"
Lo script funziona anche senza Word, usando librerie alternative per la conversione.

### "PyMuPDF non disponibile"
La compressione avanzata richiede PyMuPDF e Pillow:
```cmd
pip install PyMuPDF Pillow
```

### "PDF non comprimibile"
Alcuni PDF già ottimizzati potrebbero non essere comprimibili ulteriormente.

### "Errore di permessi"
Esegui il prompt dei comandi come amministratore se necessario.

## 📁 Struttura del Progetto

```
document-processor/
├── main.py                   # Script principale
├── requirements.txt          # Dipendenze Python
├── install_dependencies.bat  # Setup automatico
├── README.md                # Questa documentazione
├── .gitignore               # File da escludere da Git
└── .github/
    └── copilot-instructions.md
```

## 🤝 Contributi

Per miglioramenti o bug report, modifica direttamente il codice o crea una issue.

## 📄 Licenza

Questo progetto è rilasciato sotto licenza MIT.

---

**Creato con ❤️ per semplificare la gestione dei documenti**