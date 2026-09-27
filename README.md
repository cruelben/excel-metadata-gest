# Excel Metadata Gest (`excel-metadata-gest`)

Un set di macro VBA per Microsoft Excel che si interfacciano con il potente eseguibile esterno **ExifTool** per leggere, analizzare e rimuovere in modo permanente i metadati reali (EXIF, IPTC, XMP come modello fotocamera, coordinate GPS, data di scatto, autore, software, ecc.) da singoli file o intere cartelle.

---

## 🚀 Caratteristiche Principali

Il modulo VBA mette a disposizione tre macro principali:

1. **`LeggiMetadatiConExifTool()`** *(Sola lettura)*
   * Chiede di selezionare un file tramite una finestra di dialogo.
   * Estrae tutti i metadati incorporati e li elenca in un foglio Excel dedicato chiamato **`Metadati_ExifTool`** (Colonna A: Nome Proprietà, Colonna B: Valore).
   * **Non modifica** in alcun modo il file originale.

2. **`CancellaMetadatiConExifTool()`** *(Modifica permanente)*
   * Chiede di selezionare un singolo file.
   * Rimuove **tutti** i metadati presenti (`-all=`) sovrascrivendo l'originale senza generare file di backup (`-overwrite_original`).

3. **`CancellaMetadatiCartella()`** *(Elaborazione massiva)*
   * Chiede di selezionare una cartella sul disco, con la possibilità opzionale di includere ricorsivamente tutte le sottocartelle.
   * Pulisce i metadati da ogni file trovato e genera un report dettagliato in un foglio chiamato **`Risultati_Pulizia_Cartella`**, mostrando l'esito (`OK` / `Errore`) e il codice di uscita di ciascun file.

---

## 📂 Requisiti e Struttura delle Cartelle

Per poter funzionare, il modulo richiede l'eseguibile ufficiale di **ExifTool**, scaricabile dal sito ufficiale [exiftool.org](https://exiftool.org).

La struttura dei file sul computer deve prevedere una sottocartella denominata `ExifTool` posizionata esattamente nella **stessa cartella** in cui è salvato il file Excel (`.xlsm`):

```text
Cartella del tuo file Excel\
├── Il_Tuo_File_Excel.xlsm
└── ExifTool\
    └── exiftool.exe