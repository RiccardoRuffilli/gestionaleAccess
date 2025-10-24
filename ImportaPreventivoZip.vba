Option Compare Database
Option Explicit

' ==============================================================================
' MODULO: ImportaPreventivoZip
' DESCRIZIONE: Importa preventivi completi da file ZIP contenenti CSV
' AUTORE: Generato da Claude Code
' DATA: 2025-10-24
' ==============================================================================

' Costanti per i percorsi e configurazione
Private Const CARTELLA_LAVORO As String = "_importazione_preventivi"

' ==============================================================================
' FUNZIONE PRINCIPALE
' ==============================================================================

Public Sub ImportaPreventivoCompleto()
    On Error GoTo ErrorHandler

    Dim zipFilePath As String
    Dim tempFolder As String
    Dim csvMainPath As String
    Dim csvPersonalePath As String
    Dim csvPreventivoPath As String
    Dim eventoID As String
    Dim nuovoIDPreventivo As Long
    Dim cartellaLavoro As String

    ' Ottieni percorso cartella di lavoro
    cartellaLavoro = GetCartellaLavoro()
    If cartellaLavoro = "" Then
        MsgBox "Impossibile trovare o creare la cartella di lavoro.", vbCritical, "Errore"
        Exit Sub
    End If

    ' Messaggio iniziale
    MsgBox "Selezionare il file ZIP contenente i file CSV del preventivo.", vbInformation, "Importazione Preventivo"

    ' 1. Selezione file ZIP
    zipFilePath = SelezionaFileZip(cartellaLavoro)
    If zipFilePath = "" Then
        MsgBox "Operazione annullata dall'utente.", vbExclamation, "Annullato"
        Exit Sub
    End If

    ' 2. Creazione cartella temporanea e estrazione ZIP
    tempFolder = cartellaLavoro & "temp\" & Format(Now(), "yyyymmddhhnnss") & "\"
    If Not CreaCartellaSeNonEsiste(tempFolder) Then
        MsgBox "Impossibile creare la cartella temporanea: " & tempFolder, vbCritical, "Errore"
        Exit Sub
    End If

    If Not EstraiZip(zipFilePath, tempFolder) Then
        MsgBox "Errore durante l'estrazione del file ZIP.", vbCritical, "Errore"
        EliminaCartella tempFolder
        Exit Sub
    End If

    ' 3. Identificazione file CSV
    eventoID = EstraiEventoID(Dir(tempFolder & "*.csv"))
    If eventoID = "" Then
        MsgBox "Impossibile identificare l'evento_ID dai file CSV.", vbCritical, "Errore"
        EliminaCartella tempFolder
        Exit Sub
    End If

    csvMainPath = tempFolder & eventoID & "_main.csv"
    csvPersonalePath = tempFolder & eventoID & "_personale.csv"
    csvPreventivoPath = tempFolder & eventoID & "_preventivo.csv"

    ' 4. Verifica esistenza file richiesti
    If Not FileExists(csvMainPath) Then
        MsgBox "File main non trovato: " & csvMainPath, vbCritical, "Errore"
        EliminaCartella tempFolder
        Exit Sub
    End If

    ' 5. Inizio transazione e importazione
    Dim db As DAO.Database
    Dim ws As DAO.Workspace

    Set ws = DBEngine(0)
    Set db = CurrentDb

    ' Inizia transazione per garantire atomicità (usando Workspace)
    ws.BeginTrans

    On Error GoTo RollbackHandler

    ' 5a. Importa CSV main e crea preventivo
    nuovoIDPreventivo = ImportaCSVMain(csvMainPath, db)

    If nuovoIDPreventivo = 0 Then
        Err.Raise vbObjectError + 1, , "Errore durante l'importazione del preventivo principale"
    End If

    ' 5b. Importa CSV personale (se esiste)
    If FileExists(csvPersonalePath) Then
        If Not ImportaCSVPersonale(csvPersonalePath, nuovoIDPreventivo, db) Then
            Err.Raise vbObjectError + 2, , "Errore durante l'importazione del personale"
        End If
    End If

    ' 5c. Importa CSV preventivo (se esiste)
    If FileExists(csvPreventivoPath) Then
        If Not ImportaCSVPreventivo(csvPreventivoPath, nuovoIDPreventivo, db) Then
            Err.Raise vbObjectError + 3, , "Errore durante l'importazione dei servizi"
        End If
    End If

    ' Commit transazione (usando Workspace)
    ws.CommitTrans

    ' 6. Pulizia file temporanei
    EliminaCartella tempFolder

    ' Messaggio successo
    MsgBox "Preventivo importato con successo!" & vbCrLf & _
           "ID Preventivo: " & nuovoIDPreventivo, vbInformation, "Successo"

    Exit Sub

RollbackHandler:
    ws.Rollback
    MsgBox "Errore durante l'importazione. Tutte le modifiche sono state annullate." & vbCrLf & _
           "Dettaglio: " & Err.Description, vbCritical, "Errore"
    EliminaCartella tempFolder
    Exit Sub

ErrorHandler:
    MsgBox "Errore: " & Err.Description & " (Codice: " & Err.Number & ")", vbCritical, "Errore"
    EliminaCartella tempFolder
End Sub

' ==============================================================================
' FUNZIONI DI UTILITÀ PER FILE E CARTELLE
' ==============================================================================

Private Function SelezionaFileZip(cartellaIniziale As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker

    With fd
        .Title = "Seleziona file ZIP da importare"
        .Filters.Clear
        .Filters.Add "File ZIP", "*.zip"
        .AllowMultiSelect = False

        ' Imposta la cartella iniziale
        .InitialFileName = cartellaIniziale

        If .Show = -1 Then
            SelezionaFileZip = .SelectedItems(1)
        Else
            SelezionaFileZip = ""
        End If
    End With

    Set fd = Nothing
End Function

Private Function GetCartellaLavoro() As String
    On Error GoTo ErrorHandler

    Dim downloadsPath As String
    Dim cartellaLavoroPath As String

    ' Ottieni percorso Downloads
    downloadsPath = Environ("USERPROFILE") & "\Downloads\"

    ' Crea percorso cartella di lavoro
    cartellaLavoroPath = downloadsPath & CARTELLA_LAVORO & "\"

    ' Crea cartella se non esiste
    If Not CreaCartellaSeNonEsiste(cartellaLavoroPath) Then
        GetCartellaLavoro = ""
        Exit Function
    End If

    GetCartellaLavoro = cartellaLavoroPath
    Exit Function

ErrorHandler:
    GetCartellaLavoro = ""
End Function

Private Function CreaCartellaSeNonEsiste(percorso As String) As Boolean
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verifica se la cartella esiste
    If Not fso.FolderExists(percorso) Then
        ' Crea la cartella (ricorsivo)
        fso.CreateFolder percorso
    End If

    CreaCartellaSeNonEsiste = (Err.Number = 0)

    Set fso = Nothing
    On Error GoTo 0
End Function

Private Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Private Function EstraiZip(zipPath As String, destFolder As String) As Boolean
    On Error GoTo ErrorHandler

    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")

    Dim zipFile As Object
    Set zipFile = shellApp.Namespace(zipPath)

    If zipFile Is Nothing Then
        EstraiZip = False
        Exit Function
    End If

    Dim destFld As Object
    Set destFld = shellApp.Namespace(destFolder)

    ' Estrai tutti i file (16 = senza dialoghi)
    destFld.CopyHere zipFile.Items, 16

    ' Attendi completamento
    Dim i As Integer
    For i = 1 To 100
        DoEvents
        If Dir(destFolder & "*.csv") <> "" Then Exit For
        Application.Wait Now + TimeValue("00:00:01")
    Next i

    EstraiZip = True
    Exit Function

ErrorHandler:
    EstraiZip = False
End Function

Private Sub EliminaCartella(percorso As String)
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Elimina la cartella e tutto il suo contenuto
    If fso.FolderExists(percorso) Then
        fso.DeleteFolder percorso, True
    End If

    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Function EstraiEventoID(nomeFile As String) As String
    ' Estrae l'evento_ID dal nome file formato: evento_41732_main.csv
    Dim parti() As String
    Dim nomeBase As String

    If nomeFile = "" Then
        EstraiEventoID = ""
        Exit Function
    End If

    ' Rimuovi estensione
    nomeBase = Left(nomeFile, InStrRev(nomeFile, ".") - 1)

    ' Trova l'ultimo underscore
    Dim posUltimo As Integer
    posUltimo = InStrRev(nomeBase, "_")

    If posUltimo > 0 Then
        EstraiEventoID = Left(nomeBase, posUltimo - 1)
    Else
        EstraiEventoID = ""
    End If
End Function

' ==============================================================================
' FUNZIONI DI LETTURA CSV
' ==============================================================================

Private Function LeggiCSV(filePath As String) As Collection
    On Error GoTo ErrorHandler

    Dim righe As New Collection
    Dim fileNum As Integer
    Dim riga As String
    Dim headers() As String
    Dim primaRiga As Boolean

    fileNum = FreeFile()
    Open filePath For Input As #fileNum

    primaRiga = True

    Do While Not EOF(fileNum)
        Line Input #fileNum, riga

        If primaRiga Then
            headers = ParseCSVLine(riga)
            primaRiga = False
        Else
            Dim valori() As String
            valori = ParseCSVLine(riga)

            ' Crea dizionario per la riga
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")

            Dim i As Integer
            For i = LBound(headers) To UBound(headers)
                If i <= UBound(valori) Then
                    dict(headers(i)) = valori(i)
                Else
                    dict(headers(i)) = ""
                End If
            Next i

            righe.Add dict
        End If
    Loop

    Close #fileNum

    Set LeggiCSV = righe
    Exit Function

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Errore durante la lettura del CSV: " & filePath & vbCrLf & Err.Description, vbCritical
    Set LeggiCSV = Nothing
End Function

Private Function ParseCSVLine(riga As String) As String()
    ' Parser CSV che gestisce virgolette e separatori
    Dim risultato() As String
    Dim temp As String
    Dim inQuote As Boolean
    Dim i As Integer
    Dim char As String
    Dim campoCorrente As String
    Dim numCampi As Integer

    ReDim risultato(0 To 100) ' Array temporaneo
    numCampi = 0
    campoCorrente = ""
    inQuote = False

    For i = 1 To Len(riga)
        char = Mid(riga, i, 1)

        If char = """" Then
            inQuote = Not inQuote
        ElseIf char = "," And Not inQuote Then
            risultato(numCampi) = campoCorrente
            numCampi = numCampi + 1
            campoCorrente = ""
        Else
            campoCorrente = campoCorrente & char
        End If
    Next i

    ' Aggiungi ultimo campo
    risultato(numCampi) = campoCorrente

    ' Ridimensiona array al numero effettivo di campi
    ReDim Preserve risultato(0 To numCampi)

    ParseCSVLine = risultato
End Function

' ==============================================================================
' FUNZIONI DI IMPORTAZIONE DATI
' ==============================================================================

Private Function ImportaCSVMain(csvPath As String, db As DAO.Database) As Long
    On Error GoTo ErrorHandler

    Dim righe As Collection
    Set righe = LeggiCSV(csvPath)

    If righe Is Nothing Or righe.Count = 0 Then
        ImportaCSVMain = 0
        Exit Function
    End If

    ' Prendi la prima riga (il main ha una sola riga)
    Dim riga As Object
    Set riga = righe(1)

    ' Verifica duplicati usando l'ID originale salvato nel campo "riferimento"
    Dim idOriginale As String
    idOriginale = "EVENTO_" & CStr(riga("id"))

    If VerificaDuplicatoPreventivo(idOriginale, db) Then
        MsgBox "Preventivo già esistente nel database (ID: " & idOriginale & ")", vbExclamation, "Duplicato"
        ImportaCSVMain = 0
        Exit Function
    End If

    ' Crea nuovo record preventivo
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("preventivi", dbOpenDynaset)

    rs.AddNew

    ' Mapping campi secondo conversione_main-preventivi.xlsx

    ' ID_cliente
    If Not IsNull(riga("cliente_id")) And riga("cliente_id") <> "" Then
        rs("ID_cliente") = CLng(riga("cliente_id"))
    End If

    ' id_referente_videorent (responsabile_id)
    If Not IsNull(riga("responsabile_id")) And riga("responsabile_id") <> "" Then
        rs("id_referente_videorent") = CLng(riga("responsabile_id"))
    End If

    ' Riferimento (salva ID originale + referente_id)
    Dim riferimentoText As String
    riferimentoText = idOriginale
    If Not IsNull(riga("referente_id")) And riga("referente_id") <> "" Then
        riferimentoText = riferimentoText & " - REF:" & riga("referente_id")
    End If
    rs("Riferimento") = Left(riferimentoText, 255)

    ' Date e ore - separare i campi datetime
    Dim dataOra As Variant

    ' Data/Ora allestimento
    dataOra = ParseDateTime(riga("data_ora_allestimento"))
    If Not IsNull(dataOra) Then
        rs("Data allestimento") = CDate(dataOra)
        rs("Ora allestimento") = CDate(dataOra)
    End If

    ' Data/Ora inizio
    dataOra = ParseDateTime(riga("data_ora_inizio"))
    If Not IsNull(dataOra) Then
        rs("Data inizio") = CDate(dataOra)
        rs("Ora inizio") = CDate(dataOra)
    End If

    ' Data/Ora fine
    dataOra = ParseDateTime(riga("data_ora_fine"))
    If Not IsNull(dataOra) Then
        rs("Data fine") = CDate(dataOra)
        rs("Ora fine") = CDate(dataOra)
    End If

    ' Data/Ora disallestimento
    dataOra = ParseDateTime(riga("data_ora_disallestimento"))
    If Not IsNull(dataOra) Then
        rs("Data disallestimento") = CDate(dataOra)
        rs("Ora disallestimento") = CDate(dataOra)
    End If

    ' Flag booleani
    rs("confermato") = CBool(riga("flag_confermato"))
    rs("annullato") = CBool(riga("flag_annullato"))
    rs("planner") = CBool(riga("flag_planner"))
    rs("Fatturato") = CBool(riga("flag_fatturazione"))

    ' Sconto cliente
    If Not IsNull(riga("sconto_cliente")) And riga("sconto_cliente") <> "" Then
        rs("sconto cliente") = CDbl(riga("sconto_cliente"))
    End If

    ' Pagamento
    If Not IsNull(riga("pagamento")) And riga("pagamento") <> "" Then
        rs("pagamento") = riga("pagamento")
    End If

    ' Gruppo
    If Not IsNull(riga("gruppo")) And riga("gruppo") <> "" Then
        rs("gruppo") = riga("gruppo")
    End If

    ' Note location (aggregare location_id + note_location)
    Dim noteLocation As String
    noteLocation = ""
    If Not IsNull(riga("location_id")) And riga("location_id") <> "" Then
        noteLocation = "Location ID: " & riga("location_id")
    End If
    If Not IsNull(riga("note_location")) And riga("note_location") <> "" Then
        If noteLocation <> "" Then noteLocation = noteLocation & vbCrLf
        noteLocation = noteLocation & riga("note_location")
    End If
    If noteLocation <> "" Then
        rs("Note location") = noteLocation
    End If

    ' Note (note_cliente)
    If Not IsNull(riga("note_cliente")) And riga("note_cliente") <> "" Then
        rs("Note") = riga("note_cliente")
    End If

    ' Note_fatturazione
    If Not IsNull(riga("note_fatturazione")) And riga("note_fatturazione") <> "" Then
        rs("Note_fatturazione") = riga("note_fatturazione")
    End If

    ' Accessori vari (aggregare: nome_evento, note_interne, note_runner_arrivo, note_runner_disallestimento, note_scheda_lavoro)
    Dim accessoriVari As String
    accessoriVari = ""

    If Not IsNull(riga("nome_evento")) And riga("nome_evento") <> "" Then
        accessoriVari = "EVENTO: " & riga("nome_evento")
    End If

    If Not IsNull(riga("note_interne")) And riga("note_interne") <> "" Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Note interne: " & riga("note_interne")
    End If

    If Not IsNull(riga("note_runner_arrivo")) And riga("note_runner_arrivo") <> "" Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Runner arrivo: " & riga("note_runner_arrivo")
    End If

    If Not IsNull(riga("note_runner_disallestimento")) And riga("note_runner_disallestimento") <> "" Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Runner disallestimento: " & riga("note_runner_disallestimento")
    End If

    If Not IsNull(riga("note_scheda_lavoro")) And riga("note_scheda_lavoro") <> "" Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Scheda lavoro: " & riga("note_scheda_lavoro")
    End If

    If accessoriVari <> "" Then
        rs("Accessori vari") = accessoriVari
    End If

    rs.Update

    ' Recupera l'ID autonumerico appena creato
    rs.Bookmark = rs.LastModified
    ImportaCSVMain = rs("ID_preventivo")

    rs.Close
    Set rs = Nothing

    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
        rs.Close
    End If
    MsgBox "Errore ImportaCSVMain: " & Err.Description, vbCritical
    ImportaCSVMain = 0
End Function

Private Function ImportaCSVPersonale(csvPath As String, idPreventivo As Long, db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    Dim righe As Collection
    Set righe = LeggiCSV(csvPath)

    If righe Is Nothing Or righe.Count = 0 Then
        ImportaCSVPersonale = True ' Nessun personale da importare, non è un errore
        Exit Function
    End If

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("Tecnici preventivati", dbOpenDynaset)

    Dim riga As Object
    For Each riga In righe
        rs.AddNew

        ' ID_preventivo (collegamento al preventivo appena creato)
        rs("ID_preventivo") = idPreventivo

        ' ID_Tecnico (user_id)
        If Not IsNull(riga("user_id")) And riga("user_id") <> "" Then
            rs("ID_Tecnico") = CLng(riga("user_id"))
        End If

        ' Date e ore - duplicare i valori come da mapping
        Dim dataInizio As Variant
        Dim oraInizio As Variant
        Dim dataFine As Variant
        Dim oraFine As Variant

        ' Data inizio -> data_allestimento_tecnico e data_inizio_tecnico
        If Not IsNull(riga("data_inizio")) And riga("data_inizio") <> "" Then
            dataInizio = CDate(riga("data_inizio"))
            rs("data_allestimento_tecnico") = dataInizio
            rs("data_inizio_tecnico") = dataInizio
        End If

        ' Ora inizio -> ora_allestimento_tecnico e ora_inizio_tecnico
        If Not IsNull(riga("ora_inizio")) And riga("ora_inizio") <> "" Then
            oraInizio = CDate(riga("ora_inizio"))
            rs("ora_allestimento_tecnico") = oraInizio
            rs("ora_inizio_tecnico") = oraInizio
        End If

        ' Data fine -> data_fine_tecnico e data_disallestimento_tecnico
        If Not IsNull(riga("data_fine")) And riga("data_fine") <> "" Then
            dataFine = CDate(riga("data_fine"))
            rs("data_fine_tecnico") = dataFine
            rs("data_disallestimento_tecnico") = dataFine
        End If

        ' Ora fine -> ora_fine_tecnico e ora_disallestimento_tecnico
        If Not IsNull(riga("ora_fine")) And riga("ora_fine") <> "" Then
            oraFine = CDate(riga("ora_fine"))
            rs("ora_fine_tecnico") = oraFine
            rs("ora_disallestimento_tecnico") = oraFine
        End If

        ' Conferma tecnico
        If Not IsNull(riga("confirmed")) And riga("confirmed") <> "" Then
            rs("Conferma tecnico") = CBool(riga("confirmed"))
        End If

        rs.Update
    Next riga

    rs.Close
    Set rs = Nothing

    ImportaCSVPersonale = True
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
        rs.Close
    End If
    MsgBox "Errore ImportaCSVPersonale: " & Err.Description, vbCritical
    ImportaCSVPersonale = False
End Function

Private Function ImportaCSVPreventivo(csvPath As String, idPreventivo As Long, db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    Dim righe As Collection
    Set righe = LeggiCSV(csvPath)

    If righe Is Nothing Or righe.Count = 0 Then
        ImportaCSVPreventivo = True ' Nessun servizio da importare, non è un errore
        Exit Function
    End If

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("Servizi preventivati", dbOpenDynaset)

    Dim riga As Object
    For Each riga In righe
        rs.AddNew

        ' ID_preventivo (collegamento al preventivo appena creato)
        rs("ID_preventivo") = idPreventivo

        ' ID_servizio (item_id)
        If Not IsNull(riga("item_id")) And riga("item_id") <> "" Then
            rs("ID_servizio") = CLng(riga("item_id"))
        End If

        ' Ordine
        If Not IsNull(riga("ord")) And riga("ord") <> "" Then
            rs("ordine") = CLng(riga("ord"))
        End If

        ' Quantità
        If Not IsNull(riga("qty")) And riga("qty") <> "" Then
            rs("quantità") = CDbl(riga("qty"))
        End If

        ' Giorni
        If Not IsNull(riga("giorni")) And riga("giorni") <> "" Then
            rs("giorni") = CLng(riga("giorni"))
        End If

        ' Listino (unit_price)
        If Not IsNull(riga("unit_price")) And riga("unit_price") <> "" Then
            rs("Listino") = CCur(riga("unit_price"))
        End If

        ' Importo (unit_price_net)
        If Not IsNull(riga("unit_price_net")) And riga("unit_price_net") <> "" Then
            rs("Importo") = CCur(riga("unit_price_net"))
        End If

        ' Sconto (discount_pct)
        If Not IsNull(riga("discount_pct")) And riga("discount_pct") <> "" Then
            rs("Sconto") = CDbl(riga("discount_pct"))
        End If

        ' Note articolo
        If Not IsNull(riga("note")) And riga("note") <> "" Then
            rs("note_articolo") = riga("note")
        End If

        rs.Update
    Next riga

    rs.Close
    Set rs = Nothing

    ImportaCSVPreventivo = True
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
        rs.Close
    End If
    MsgBox "Errore ImportaCSVPreventivo: " & Err.Description, vbCritical
    ImportaCSVPreventivo = False
End Function

' ==============================================================================
' FUNZIONI DI SUPPORTO
' ==============================================================================

Private Function VerificaDuplicatoPreventivo(riferimento As String, db As DAO.Database) As Boolean
    On Error Resume Next

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT COUNT(*) AS Totale FROM preventivi WHERE Riferimento LIKE '" & riferimento & "%'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        VerificaDuplicatoPreventivo = (rs("Totale") > 0)
    Else
        VerificaDuplicatoPreventivo = False
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Function ParseDateTime(dateTimeString As String) As Variant
    On Error Resume Next

    If IsNull(dateTimeString) Or dateTimeString = "" Then
        ParseDateTime = Null
        Exit Function
    End If

    ' Prova a convertire in data
    Dim dt As Date
    dt = CDate(dateTimeString)

    If Err.Number = 0 Then
        ParseDateTime = dt
    Else
        ParseDateTime = Null
    End If

    On Error GoTo 0
End Function
