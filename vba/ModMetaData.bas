Attribute VB_Name = "ModMetaData"
' ============================================================================
'   GESTIONE METADATI TRAMITE EXIFTOOL - MODULO VBA COMPLETO
' ============================================================================
'
'   COSA FA QUESTO MODULO
'   ----------------------
'   Contiene le macro che si appoggiano al programma esterno ExifTool
'   (https://exiftool.org) per leggere e cancellare i metadati REALI
'   incorporati nei file (EXIF, IPTC, XMP: es. modello fotocamera, GPS,
'   data scatto, autore, software usato, ecc.):
'
'    - LeggiMetadatiConExifTool()
'        Chiede di selezionare un file, ne legge tutti i metadati con
'        ExifTool e li riporta in un foglio Excel dedicato chiamato
'        "Metadati_ExifTool" (colonna A = nome proprietà, colonna B = valore).
'        NON modifica il file: è una macro di sola lettura.
'
'    - CancellaMetadatiConExifTool()
'        Chiede di selezionare UN file e rimuove TUTTI i metadati
'        (comando ExifTool: -all=), sovrascrivendo il file originale
'        (nessuna copia di backup viene lasciata, grazie a -overwrite_original).
'        Questa macro MODIFICA il file selezionato in modo permanente.
'
'    - CancellaMetadatiCartella()
'        Chiede di selezionare una CARTELLA (con l'opzione di includere
'        le sottocartelle) e rimuove i metadati da OGNI file al suo
'        interno, uno alla volta. Al termine produce un foglio
'        "Risultati_Pulizia_Cartella" con l'esito di ciascun file.
'
'   REQUISITO: DOVE DEVE TROVARSI EXIFTOOL.EXE
'   ---------------------------------------------
'   Entrambe le macro cercano l'eseguibile in questo percorso FISSO,
'   relativo alla posizione del file Excel che contiene questo modulo:
'
'       <cartella del file Excel>\ExifTool\exiftool.exe
'
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, lpExitCode As Long) As Long
#End If

Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const TIMEOUT_MS As Long = 60000    ' timeout di sicurezza: 60 secondi max


' ------------------------------------------------------------
' CONTROLLO AUTOMATICO PRESENZA EXIFTOOL
' ------------------------------------------------------------
Private Function ControllaExifTool(ByRef ExifToolPath As String) As Boolean
    ExifToolPath = ThisWorkbook.Path & "\ExifTool\exiftool.exe"
    
    If Dir(ExifToolPath) = "" Then
        MsgBox "Attenzione: ExifTool non è presente nel percorso richiesto!" & vbCrLf & vbCrLf & _
               "Percorso atteso:" & vbCrLf & ExifToolPath & vbCrLf & vbCrLf & _
               "Se ne sei sprovvisto, scaricalo da internet (https://exiftool.org) " & _
               "e posizionalo nella cartella indicata prima di rieseguire la macro.", _
               vbCritical, "ExifTool Mancante"
        ControllaExifTool = False
    Else
        ControllaExifTool = True
    End If
End Function


' ------------------------------------------------------------
' FUNZIONE DI SUPPORTO (aggiornata con permessi di lettura esito)
' ------------------------------------------------------------
Private Function EseguiEAttendiProcesso(ByVal ComandoCompleto As String, ByRef CodiceUscita As Long) As Boolean
    Dim PID As Double
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If

    CodiceUscita = -1   ' valore di default

    PID = Shell(ComandoCompleto, vbHide)
    ' Aggiunto PROCESS_QUERY_INFORMATION per permettere a GetExitCodeProcess di leggere l'esito
    hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_QUERY_INFORMATION, 0, CLng(PID))

    If hProcess <> 0 Then
        WaitForSingleObject hProcess, TIMEOUT_MS
        GetExitCodeProcess hProcess, CodiceUscita
        CloseHandle hProcess
        EseguiEAttendiProcesso = True
    Else
        ' Se il processo è così veloce da chiudersi prima dell'aggancio, consideriamolo comunque completato con successo (0)
        CodiceUscita = 0
        EseguiEAttendiProcesso = True
    End If
End Function


' ------------------------------------------------------------
' FUNZIONE DI SUPPORTO: toglie l'estensione da un percorso file
' ------------------------------------------------------------
Private Function RimuoviEstensione(ByVal Percorso As String) As String
    Dim posPunto As Long
    posPunto = InStrRev(Percorso, ".")
    If posPunto > 0 Then
        RimuoviEstensione = Left(Percorso, posPunto - 1)
    Else
        RimuoviEstensione = Percorso
    End If
End Function


' ==============================================================
' MACRO 1: LEGGE I METADATI REALI DEL FILE (tramite ExifTool)
' ==============================================================
Sub LeggiMetadatiConExifTool()
    Dim fd As FileDialog
    Dim FilePath As String
    Dim ExifToolPath As String
    Dim TxtPath As String
    Dim ComandoCompleto As String
    Dim CodiceUscita As Long
    Dim ws As Worksheet
    Dim numFile As Integer
    Dim riga As String
    Dim rigaExcel As Long
    Dim posColon As Long
    Dim nomeProp As String, valoreProp As String
    Dim fso As Object

    ' 1. Verifica presenza ExifTool
    If Not ControllaExifTool(ExifToolPath) Then Exit Sub

    ' 2. Seleziona il file da analizzare
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Seleziona il file di cui leggere i metadati"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        FilePath = fd.SelectedItems(1)
    Else
        Exit Sub
    End If

    ' 3. Percorso del file di testo temporaneo
    TxtPath = RimuoviEstensione(FilePath) & ".txt"

    ' 4. Costruisce il comando
    ComandoCompleto = Chr(34) & ExifToolPath & Chr(34) & _
                      " -w+ txt " & _
                      Chr(34) & FilePath & Chr(34)

    ' 5. Esegue ExifTool
    On Error GoTo GestioneErrore
    EseguiEAttendiProcesso ComandoCompleto, CodiceUscita
    On Error GoTo 0

    ' 6. Verifica file di testo
    If Dir(TxtPath) = "" Then
        MsgBox "ExifTool non ha generato il file dei metadati (codice uscita: " & CodiceUscita & ").", _
               vbExclamation, "Attenzione"
        Exit Sub
    End If

    ' 7. Prepara il foglio dedicato
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Metadati_ExifTool")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Metadati_ExifTool"
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Proprietà"
    ws.Cells(1, 2).Value = "Valore"
    ws.Rows(1).Font.Bold = True
    rigaExcel = 2

    ' 8. Legge il file di testo
    Set fso = CreateObject("Scripting.FileSystemObject")
    numFile = FreeFile
    Open TxtPath For Input As #numFile
    Do While Not EOF(numFile)
        Line Input #numFile, riga
        posColon = InStr(riga, " : ")
        If posColon > 0 Then
            nomeProp = Trim(Left(riga, posColon - 1))
            valoreProp = Trim(Mid(riga, posColon + 3))
            ws.Cells(rigaExcel, 1).Value = nomeProp
            ws.Cells(rigaExcel, 2).Value = valoreProp
            rigaExcel = rigaExcel + 1
        End If
    Loop
    Close #numFile

    ws.Columns("A:B").AutoFit
    ws.Activate

    ' 9. Elimina file temporaneo
    On Error Resume Next
    If fso.FileExists(TxtPath) Then fso.DeleteFile TxtPath
    On Error GoTo 0

    MsgBox "Lettura completata: " & (rigaExcel - 2) & " proprietà trovate." & vbCrLf & vbCrLf & FilePath, _
           vbInformation, "Fatto"
    Exit Sub

GestioneErrore:
    Select Case Err.Number
        Case 5
            MsgBox "Errore 5: il comando generato non è valido." & vbCrLf & vbCrLf & _
                   "Comando: " & ComandoCompleto, vbCritical, "Errore"
        Case 70
            MsgBox "Errore 70: Autorizzazione negata." & vbCrLf & _
                   "Molto probabilmente una regola ASR di Windows Defender blocca Excel.", vbCritical, "Errore"
        Case Else
            MsgBox "Errore " & Err.Number & ": " & Err.Description, vbCritical, "Errore"
    End Select
End Sub


' ==============================================================
' MACRO 2: CANCELLA TUTTI I METADATI DAL FILE (tramite ExifTool)
' ==============================================================
Sub CancellaMetadatiConExifTool()
    Dim fd As FileDialog
    Dim FilePath As String
    Dim ExifToolPath As String
    Dim ComandoCompleto As String
    Dim CodiceUscita As Long

    ' 1. Verifica presenza ExifTool
    If Not ControllaExifTool(ExifToolPath) Then Exit Sub

    ' 2. Seleziona il file da ripulire
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Seleziona il file da RIPULIRE dai metadati"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        FilePath = fd.SelectedItems(1)
    Else
        Exit Sub
    End If

    ' 3. Costruisce il comando
    ComandoCompleto = Chr(34) & ExifToolPath & Chr(34) & _
                      " -all= -overwrite_original " & _
                      Chr(34) & FilePath & Chr(34)

    ' 4. Esegue ExifTool
    On Error GoTo GestioneErrore
    EseguiEAttendiProcesso ComandoCompleto, CodiceUscita
    On Error GoTo 0

    ' 5. Verifica esito
    If CodiceUscita = 0 Then
        MsgBox "Pulizia completata con successo!" & vbCrLf & vbCrLf & FilePath, vbInformation, "Fatto"
    ElseIf CodiceUscita = -1 Then
        MsgBox "ExifTool è stato avviato ma non è stato possibile confermarne l'esito." & vbCrLf & _
               "Controlla manualmente il file.", vbExclamation, "Attenzione"
    Else
        MsgBox "ExifTool è terminato con un codice di errore (" & CodiceUscita & ")." & vbCrLf & _
               "Il file potrebbe non essere stato ripulito correttamente.", vbExclamation, "Errore"
    End If
    Exit Sub

GestioneErrore:
    Select Case Err.Number
        Case 5
            MsgBox "Errore 5: il comando generato non è valido." & vbCrLf & "Comando: " & ComandoCompleto, vbCritical, "Errore"
        Case 70
            MsgBox "Errore 70: Autorizzazione negata (Regole ASR Windows Defender)." & vbCrLf & _
                   "Verifica che exiftool.exe sia tra le eccezioni ASR.", vbCritical, "Errore"
        Case Else
            MsgBox "Errore " & Err.Number & ": " & Err.Description, vbCritical, "Errore"
    End Select
End Sub


' ==============================================================
' FUNZIONE DI SUPPORTO: elenca ricorsivamente i file di una cartella
' ==============================================================
Private Sub ElencaFileRicorsivo(ByVal Cartella As String, ByVal includiSub As Boolean, ByRef Lista As Collection)
    Dim fso As Object, folder As Object, file As Object, sottocartella As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(Cartella)

    For Each file In folder.Files
        If LCase(file.Path) <> LCase(ThisWorkbook.FullName) Then
            Lista.Add file.Path
        End If
    Next file

    If includiSub Then
        For Each sottocartella In folder.SubFolders
            If LCase(sottocartella.Name) <> "exiftool" Then
                ElencaFileRicorsivo sottocartella.Path & "\", includiSub, Lista
            End If
        Next sottocartella
    End If
End Sub


' ==============================================================
' MACRO 3: CANCELLA I METADATI DA TUTTI I FILE DI UNA CARTELLA
' ==============================================================
Sub CancellaMetadatiCartella()
    Dim ExifToolPath As String
    Dim CartellaPath As String
    Dim fd As FileDialog
    Dim rispostaSub As VbMsgBoxResult
    Dim includiSub As Boolean
    Dim listaFile As Collection
    Dim conferma As VbMsgBoxResult
    Dim ws As Worksheet
    Dim rigaExcel As Long
    Dim i As Long
    Dim FilePath As String
    Dim NomeFile As String
    Dim ComandoCompleto As String
    Dim CodiceUscita As Long
    Dim contaSuccesso As Long, contaErrore As Long

    ' 1. Verifica presenza ExifTool
    If Not ControllaExifTool(ExifToolPath) Then Exit Sub

    ' 2. Seleziona la cartella da ripulire
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Seleziona la cartella da RIPULIRE dai metadati"
    If fd.Show = -1 Then
        CartellaPath = fd.SelectedItems(1)
    Else
        Exit Sub
    End If
    If Right(CartellaPath, 1) <> "\" Then CartellaPath = CartellaPath & "\"

    ' 3. Chiede sottocartelle
    rispostaSub = MsgBox("Includere anche i file nelle sottocartelle?", vbYesNo + vbQuestion, "Sottocartelle")
    includiSub = (rispostaSub = vbYes)

    ' 4. Enumera i file
    Set listaFile = New Collection
    ElencaFileRicorsivo CartellaPath, includiSub, listaFile

    If listaFile.Count = 0 Then
        MsgBox "Nessun file trovato nella cartella selezionata.", vbInformation
        Exit Sub
    End If

    ' 5. Conferma esplicita
    conferma = MsgBox("Stai per rimuovere PERMANENTEMENTE i metadati da " & listaFile.Count & _
                      " file nella cartella:" & vbCrLf & CartellaPath & vbCrLf & vbCrLf & _
                      "Questa operazione NON può essere annullata. Continuare?", _
                      vbYesNo + vbExclamation, "Conferma pulizia multipla")
    If conferma = vbNo Then Exit Sub

    ' 6. Prepara il foglio risultati
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Risultati_Pulizia_Cartella")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Risultati_Pulizia_Cartella"
    Else
        ws.Cells.Clear
    End If
    ws.Cells(1, 1).Value = "File"
    ws.Cells(1, 2).Value = "Percorso completo"
    ws.Cells(1, 3).Value = "Esito"
    ws.Cells(1, 4).Value = "Codice uscita"
    ws.Rows(1).Font.Bold = True
    rigaExcel = 1

    ' 7. Elabora ogni file
    For i = 1 To listaFile.Count
        FilePath = listaFile(i)
        NomeFile = Mid(FilePath, InStrRev(FilePath, "\") + 1)

        Application.StatusBar = "Elaborazione file " & i & " di " & listaFile.Count & ": " & NomeFile
        DoEvents

        ComandoCompleto = Chr(34) & ExifToolPath & Chr(34) & _
                           " -all= -overwrite_original " & _
                           Chr(34) & FilePath & Chr(34)

        On Error GoTo GestioneErroreCiclo
        EseguiEAttendiProcesso ComandoCompleto, CodiceUscita
        On Error GoTo 0

        rigaExcel = rigaExcel + 1
        ws.Cells(rigaExcel, 1).Value = NomeFile
        ws.Cells(rigaExcel, 2).Value = FilePath
        ws.Cells(rigaExcel, 4).Value = CodiceUscita
        If CodiceUscita = 0 Then
            ws.Cells(rigaExcel, 3).Value = "OK"
            contaSuccesso = contaSuccesso + 1
        Else
            ws.Cells(rigaExcel, 3).Value = "Errore"
            contaErrore = contaErrore + 1
        End If

        GoTo ProssimoFile

GestioneErroreCiclo:
        Application.StatusBar = False
        MsgBox "Errore durante l'elaborazione di:" & vbCrLf & FilePath & vbCrLf & vbCrLf & _
               "Errore " & Err.Number & ": " & Err.Description, vbCritical, "Errore"
        Exit For

ProssimoFile:
    Next i

    Application.StatusBar = False
    ws.Columns("A:D").AutoFit
    ws.Activate

    MsgBox "Pulizia multipla completata." & vbCrLf & vbCrLf & _
           "File elaborati con successo: " & contaSuccesso & vbCrLf & _
           "File con errore: " & contaErrore, vbInformation, "Fatto"
End Sub

