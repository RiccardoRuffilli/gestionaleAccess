Option Compare Database
Option Explicit

' ==============================================================================
' MODULO: ImportaPreventivoJSON
' DESCRIZIONE: Importa preventivi da file JSON
' AUTORE: Generato da Claude Code
' DATA: 2025-10-24
' ==============================================================================

' Dichiarazione API Windows per Sleep
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Costanti per i percorsi e configurazione
Private Const CARTELLA_LAVORO As String = "_importazione_preventivi"

' ==============================================================================
' FUNZIONE PRINCIPALE
' ==============================================================================

Public Sub ImportaPreventivoJSON()
    ' Importa un preventivo da file JSON

    Dim jsonFilePath As String
    Dim cartellaLavoro As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim nuovoIDPreventivo As Long

    ' Inizializza database e workspace
    Set ws = DBEngine(0)
    Set db = CurrentDb

    ' Ottieni percorso cartella di lavoro
    cartellaLavoro = GetCartellaLavoro()
    If cartellaLavoro = "" Then
        MsgBox "Impossibile trovare o creare la cartella di lavoro.", vbCritical, "Errore"
        Exit Sub
    End If

    ' Messaggio iniziale
    MsgBox "Selezionare il file JSON contenente il preventivo.", vbInformation, "Importazione Preventivo"

    ' Selezione file JSON
    jsonFilePath = SelezionaFileJSON(cartellaLavoro)
    If jsonFilePath = "" Then
        MsgBox "Operazione annullata dall'utente.", vbExclamation, "Annullato"
        Exit Sub
    End If

    ' Leggi e parsifica il JSON
    Dim jsonData As Object
    Set jsonData = LeggiFileJSON(jsonFilePath)

    If jsonData Is Nothing Then
        MsgBox "Errore durante la lettura del file JSON.", vbCritical, "Errore"
        Exit Sub
    End If

    ' Verifica struttura JSON - accesso diretto all'oggetto JavaScript
    On Error Resume Next
    Dim eventoObj As Object
    Set eventoObj = CallByName(jsonData, "evento", VbGet)

    If Err.Number <> 0 Or eventoObj Is Nothing Then
        MsgBox "File JSON non valido: manca la sezione 'evento'.", vbCritical, "Errore"
        Exit Sub
    End If
    On Error GoTo 0

    ' Estrai ID preventivo
    Dim idPreventivoOriginale As String
    idPreventivoOriginale = CStr(GetJSONValue(eventoObj, "id"))

    If idPreventivoOriginale = "" Then
        MsgBox "Impossibile estrarre l'ID del preventivo dal JSON.", vbCritical, "Errore"
        Exit Sub
    End If

    ' Verifica se esiste già un preventivo con questo ID
    If VerificaEsistenzaPreventivo(idPreventivoOriginale, db) Then
        Dim risposta As VbMsgBoxResult
        risposta = MsgBox("Esiste già un preventivo con ID: " & idPreventivoOriginale & vbCrLf & vbCrLf & _
                          "Vuoi sostituirlo con i nuovi dati?" & vbCrLf & vbCrLf & _
                          "ATTENZIONE: Questa operazione eliminerà il preventivo esistente e tutti i dati collegati.", _
                          vbYesNo + vbQuestion, "Preventivo Esistente")

        If risposta = vbNo Then
            MsgBox "Operazione annullata dall'utente.", vbInformation, "Annullato"
            Exit Sub
        Else
            ' Elimina il preventivo esistente e tutti i record collegati
            EliminaPreventivo idPreventivoOriginale, db
        End If
    End If

    ' Inizia transazione
    ws.BeginTrans

    ' Importa dati evento (preventivo)
    nuovoIDPreventivo = ImportaEvento(eventoObj, db)

    If nuovoIDPreventivo = 0 Then
        ws.Rollback
        MsgBox "Errore durante l'importazione del preventivo.", vbCritical, "Errore"
        Exit Sub
    End If

    ' Importa personale (se presente)
    On Error Resume Next
    Dim personaleArray As Object
    Set personaleArray = CallByName(jsonData, "personale", VbGet)

    If Err.Number = 0 And Not personaleArray Is Nothing Then
        On Error GoTo 0
        If Not ImportaPersonale(personaleArray, nuovoIDPreventivo, db) Then
            ws.Rollback
            MsgBox "Errore durante l'importazione del personale.", vbCritical, "Errore"
            Exit Sub
        End If
    End If
    On Error GoTo 0

    ' Importa servizi (se presente)
    On Error Resume Next
    Dim serviziArray As Object
    Set serviziArray = CallByName(jsonData, "servizi", VbGet)

    If Err.Number = 0 And Not serviziArray Is Nothing Then
        On Error GoTo 0
        If Not ImportaServizi(serviziArray, nuovoIDPreventivo, db) Then
            ws.Rollback
            MsgBox "Errore durante l'importazione dei servizi.", vbCritical, "Errore"
            Exit Sub
        End If
    End If
    On Error GoTo 0

    ' Commit transazione
    ws.CommitTrans

    ' Messaggio successo
    MsgBox "Preventivo importato con successo!" & vbCrLf & _
           "ID Preventivo: " & nuovoIDPreventivo & vbCrLf & _
           "ID Originale: " & idPreventivoOriginale, vbInformation, "Successo"

End Sub

' ==============================================================================
' FUNZIONI DI UTILITÀ
' ==============================================================================

Private Function GetCartellaLavoro() As String
    Dim downloadsPath As String
    Dim cartellaLavoroPath As String

    ' Ottieni percorso Downloads
    downloadsPath = Environ("USERPROFILE")
    If Right(downloadsPath, 1) <> "\" Then downloadsPath = downloadsPath & "\"
    downloadsPath = downloadsPath & "Downloads\"

    ' Crea percorso cartella di lavoro
    cartellaLavoroPath = downloadsPath & CARTELLA_LAVORO
    If Right(cartellaLavoroPath, 1) <> "\" Then cartellaLavoroPath = cartellaLavoroPath & "\"

    ' Crea cartella se non esiste
    If Not CreaCartellaSeNonEsiste(cartellaLavoroPath) Then
        GetCartellaLavoro = ""
        Exit Function
    End If

    GetCartellaLavoro = cartellaLavoroPath
End Function

Private Function CreaCartellaSeNonEsiste(percorso As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(percorso) Then
        fso.CreateFolder percorso
    End If

    CreaCartellaSeNonEsiste = True
    Set fso = Nothing
End Function

Private Function SelezionaFileJSON(cartellaIniziale As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker

    With fd
        .Title = "Seleziona file JSON da importare"
        .Filters.Clear
        .Filters.Add "File JSON", "*.json"
        .AllowMultiSelect = False
        .InitialFileName = cartellaIniziale

        If .Show = -1 Then
            SelezionaFileJSON = .SelectedItems(1)
        Else
            SelezionaFileJSON = ""
        End If
    End With

    Set fd = Nothing
End Function

Private Function LeggiFileJSON(filePath As String) As Object
    ' Legge il file JSON con encoding UTF-8 corretto

    ' Usa ADODB.Stream per leggere UTF-8
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath

    Dim jsonText As String
    jsonText = stream.ReadText
    stream.Close

    ' Parse JSON usando ScriptControl
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"

    ' Aggiungi funzione helper per convertire JSON in oggetto VB
    sc.AddCode "function parseJSON(json) { return eval('(' + json + ')'); }"

    Dim jsObject As Object
    Set jsObject = sc.Run("parseJSON", jsonText)

    ' Restituisci direttamente l'oggetto JavaScript
    ' Non serve conversione - accederemo alle proprietà con notazione punto
    Set LeggiFileJSON = jsObject

    Set stream = Nothing
    ' NON liberare sc - l'oggetto JavaScript ne ha bisogno
End Function

Private Function GetJSONValue(jsObj As Object, key As String) As Variant
    ' Ottiene un valore dall'oggetto JavaScript gestendo valori null

    On Error Resume Next
    Dim val As Variant

    ' Accesso diretto alla proprietà JavaScript
    If IsObject(jsObj) Then
        val = CallByName(jsObj, key, VbGet)
    End If

    If Err.Number <> 0 Or IsNull(val) Or IsEmpty(val) Then
        GetJSONValue = Null
    ElseIf VarType(val) = vbString Then
        If val = "null" Or val = "" Then
            GetJSONValue = Null
        Else
            GetJSONValue = val
        End If
    Else
        GetJSONValue = val
    End If

    On Error GoTo 0
End Function

' ==============================================================================
' FUNZIONI DI IMPORTAZIONE
' ==============================================================================

Private Function ImportaEvento(evento As Object, db As DAO.Database) As Long
    ' Importa i dati dell'evento nella tabella preventivi

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("preventivi", dbOpenDynaset, dbSeeChanges)

    rs.AddNew

    ' ID_cliente
    If Not IsNull(GetJSONValue(evento, "cliente_id")) Then
        rs("ID_cliente") = CLng(GetJSONValue(evento, "cliente_id"))
    End If

    ' id_referente_videorent
    If Not IsNull(GetJSONValue(evento, "responsabile_id")) Then
        rs("id_referente_videorent") = CLng(GetJSONValue(evento, "responsabile_id"))
    End If

    ' Riferimento (salva ID originale)
    Dim idOriginale As String
    idOriginale = "EVENTO_" & CStr(GetJSONValue(evento, "id"))

    Dim riferimentoText As String
    riferimentoText = idOriginale
    If Not IsNull(GetJSONValue(evento, "referente_id")) Then
        riferimentoText = riferimentoText & " - REF:" & GetJSONValue(evento, "referente_id")
    End If
    rs("Riferimento") = Left(riferimentoText, 255)

    ' Date e ore
    If Not IsNull(GetJSONValue(evento, "data_ora_allestimento")) Then
        Dim dtAllest As Date
        dtAllest = CDate(GetJSONValue(evento, "data_ora_allestimento"))
        rs("Data allestimento") = dtAllest
        rs("Ora allestimento") = dtAllest
    End If

    If Not IsNull(GetJSONValue(evento, "data_ora_inizio")) Then
        Dim dtInizio As Date
        dtInizio = CDate(GetJSONValue(evento, "data_ora_inizio"))
        rs("Data inizio") = dtInizio
        rs("Ora inizio") = dtInizio
    End If

    If Not IsNull(GetJSONValue(evento, "data_ora_fine")) Then
        Dim dtFine As Date
        dtFine = CDate(GetJSONValue(evento, "data_ora_fine"))
        rs("Data fine") = dtFine
        rs("Ora fine") = dtFine
    End If

    If Not IsNull(GetJSONValue(evento, "data_ora_disallestimento")) Then
        Dim dtDisall As Date
        dtDisall = CDate(GetJSONValue(evento, "data_ora_disallestimento"))
        rs("Data disallestimento") = dtDisall
        rs("Ora disallestimento") = dtDisall
    End If

    ' Flag booleani
    rs("Confermato") = IIf(CLng(GetJSONValue(evento, "flag_confermato")) = 0, False, True)
    rs("annullato") = IIf(CLng(GetJSONValue(evento, "flag_annullato")) = 0, False, True)
    rs("planner") = IIf(CLng(GetJSONValue(evento, "flag_planner")) = 0, False, True)
    rs("Fatturato") = IIf(CLng(GetJSONValue(evento, "flag_fatturazione")) = 0, False, True)

    ' Sconto cliente
    If Not IsNull(GetJSONValue(evento, "sconto_cliente")) Then
        rs("sconto cliente") = CDbl(GetJSONValue(evento, "sconto_cliente"))
    End If

    ' Pagamento e gruppo
    If Not IsNull(GetJSONValue(evento, "pagamento")) Then
        rs("pagamento") = GetJSONValue(evento, "pagamento")
    End If

    If Not IsNull(GetJSONValue(evento, "gruppo")) Then
        rs("gruppo") = GetJSONValue(evento, "gruppo")
    End If

    ' Note location
    Dim noteLocation As String
    noteLocation = ""
    If Not IsNull(GetJSONValue(evento, "location_id")) Then
        noteLocation = "Location ID: " & GetJSONValue(evento, "location_id")
    End If
    If Not IsNull(GetJSONValue(evento, "note_location")) Then
        If noteLocation <> "" Then noteLocation = noteLocation & vbCrLf
        noteLocation = noteLocation & GetJSONValue(evento, "note_location")
    End If
    If noteLocation <> "" Then
        rs("Note location") = noteLocation
    End If

    ' Note
    If Not IsNull(GetJSONValue(evento, "note_cliente")) Then
        rs("Note") = GetJSONValue(evento, "note_cliente")
    End If

    ' Note_fatturazione
    If Not IsNull(GetJSONValue(evento, "note_fatturazione")) Then
        rs("Note_fatturazione") = GetJSONValue(evento, "note_fatturazione")
    End If

    ' Accessori vari (aggregare varie note)
    Dim accessoriVari As String
    accessoriVari = ""

    If Not IsNull(GetJSONValue(evento, "nome_evento")) Then
        accessoriVari = "EVENTO: " & GetJSONValue(evento, "nome_evento")
    End If

    If Not IsNull(GetJSONValue(evento, "note_interne")) Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Note interne: " & GetJSONValue(evento, "note_interne")
    End If

    If Not IsNull(GetJSONValue(evento, "note_runner_arrivo")) Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Runner arrivo: " & GetJSONValue(evento, "note_runner_arrivo")
    End If

    If Not IsNull(GetJSONValue(evento, "note_runner_disallestimento")) Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Runner disallestimento: " & GetJSONValue(evento, "note_runner_disallestimento")
    End If

    If Not IsNull(GetJSONValue(evento, "note_scheda_lavoro")) Then
        If accessoriVari <> "" Then accessoriVari = accessoriVari & vbCrLf
        accessoriVari = accessoriVari & "Scheda lavoro: " & GetJSONValue(evento, "note_scheda_lavoro")
    End If

    If accessoriVari <> "" Then
        rs("Accessori vari") = accessoriVari
    End If

    rs.Update

    ' Recupera l'ID autonumerico appena creato
    rs.Bookmark = rs.LastModified
    ImportaEvento = rs("ID_preventivo")

    rs.Close
    Set rs = Nothing
End Function

Private Function ImportaPersonale(personaleArray As Object, idPreventivo As Long, db As DAO.Database) As Boolean
    ' Importa l'array personale nella tabella Tecnici preventivati

    ' Ottieni lunghezza array
    On Error Resume Next
    Dim arrayLength As Long
    arrayLength = CallByName(personaleArray, "length", VbGet)

    If Err.Number <> 0 Or arrayLength = 0 Then
        ImportaPersonale = True
        Exit Function
    End If
    On Error GoTo 0

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("Tecnici preventivati", dbOpenDynaset, dbSeeChanges)

    Dim i As Long
    For i = 0 To arrayLength - 1
        Dim persona As Object
        Set persona = personaleArray(i)
        rs.AddNew

        ' ID_preventivo
        rs("ID_preventivo") = idPreventivo

        ' ID_Tecnico
        If Not IsNull(GetJSONValue(persona, "user_id")) Then
            rs("ID_Tecnico") = CLng(GetJSONValue(persona, "user_id"))
        End If

        ' Date e ore (duplicare nei campi allestimento e inizio)
        If Not IsNull(GetJSONValue(persona, "data_inizio")) Then
            Dim dataInizio As Date
            dataInizio = CDate(GetJSONValue(persona, "data_inizio"))
            rs("data_allestimento_tecnico") = dataInizio
            rs("data_inizio_tecnico") = dataInizio
        End If

        If Not IsNull(GetJSONValue(persona, "ora_inizio")) Then
            Dim oraInizio As Date
            oraInizio = CDate(GetJSONValue(persona, "ora_inizio"))
            rs("ora_allestimento_tecnico") = oraInizio
            rs("ora_inizio_tecnico") = oraInizio
        End If

        If Not IsNull(GetJSONValue(persona, "data_fine")) Then
            Dim dataFine As Date
            dataFine = CDate(GetJSONValue(persona, "data_fine"))
            rs("data_fine_tecnico") = dataFine
            rs("data_disallestimento_tecnico") = dataFine
        End If

        If Not IsNull(GetJSONValue(persona, "ora_fine")) Then
            Dim oraFine As Date
            oraFine = CDate(GetJSONValue(persona, "ora_fine"))
            rs("ora_fine_tecnico") = oraFine
            rs("ora_disallestimento_tecnico") = oraFine
        End If

        ' Conferma tecnico
        If Not IsNull(GetJSONValue(persona, "confirmed")) Then
            rs("Conferma tecnico") = IIf(CLng(GetJSONValue(persona, "confirmed")) = 0, False, True)
        Else
            rs("Conferma tecnico") = False
        End If

        rs.Update
    Next i

    rs.Close
    Set rs = Nothing

    ImportaPersonale = True
End Function

Private Function ImportaServizi(serviziArray As Object, idPreventivo As Long, db As DAO.Database) As Boolean
    ' Importa l'array servizi nella tabella Servizi preventivati

    ' Ottieni lunghezza array
    On Error Resume Next
    Dim arrayLength As Long
    arrayLength = CallByName(serviziArray, "length", VbGet)

    If Err.Number <> 0 Or arrayLength = 0 Then
        ImportaServizi = True
        Exit Function
    End If
    On Error GoTo 0

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("Servizi preventivati", dbOpenDynaset, dbSeeChanges)

    Dim i As Long
    For i = 0 To arrayLength - 1
        Dim servizio As Object
        Set servizio = serviziArray(i)
        rs.AddNew

        ' ID_preventivo
        rs("ID_preventivo") = idPreventivo

        ' ID_servizio
        If Not IsNull(GetJSONValue(servizio, "item_id")) Then
            rs("ID_servizio") = CLng(GetJSONValue(servizio, "item_id"))
        End If

        ' Ordine
        If Not IsNull(GetJSONValue(servizio, "ord")) Then
            rs("ordine") = CLng(GetJSONValue(servizio, "ord"))
        End If

        ' Quantità
        If Not IsNull(GetJSONValue(servizio, "qty")) Then
            rs("quantità") = CSng(GetJSONValue(servizio, "qty"))
        End If

        ' Giorni
        If Not IsNull(GetJSONValue(servizio, "giorni")) Then
            rs("giorni") = CLng(GetJSONValue(servizio, "giorni"))
        End If

        ' Listino
        If Not IsNull(GetJSONValue(servizio, "unit_price")) Then
            rs("Listino") = CCur(GetJSONValue(servizio, "unit_price"))
        End If

        ' Importo
        If Not IsNull(GetJSONValue(servizio, "unit_price_net")) Then
            rs("Importo") = CCur(GetJSONValue(servizio, "unit_price_net"))
        End If

        ' Sconto
        If Not IsNull(GetJSONValue(servizio, "discount_pct")) Then
            rs("Sconto") = CSng(GetJSONValue(servizio, "discount_pct"))
        End If

        ' Note articolo
        If Not IsNull(GetJSONValue(servizio, "note")) Then
            rs("note_articolo") = GetJSONValue(servizio, "note")
        End If

        rs.Update
    Next i

    rs.Close
    Set rs = Nothing

    ImportaServizi = True
End Function

' ==============================================================================
' FUNZIONI DI SUPPORTO
' ==============================================================================

Private Function VerificaEsistenzaPreventivo(idOriginale As String, db As DAO.Database) As Boolean
    ' Verifica se esiste già un preventivo con questo ID originale

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim riferimentoCerca As String

    riferimentoCerca = "EVENTO_" & idOriginale

    sql = "SELECT COUNT(*) AS Totale FROM preventivi WHERE Riferimento LIKE '" & riferimentoCerca & "%'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        VerificaEsistenzaPreventivo = (rs("Totale") > 0)
    Else
        VerificaEsistenzaPreventivo = False
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub EliminaPreventivo(idOriginale As String, db As DAO.Database)
    ' Elimina il preventivo esistente e tutti i record collegati

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim riferimentoCerca As String
    Dim idPreventivoDA As Long

    riferimentoCerca = "EVENTO_" & idOriginale

    ' Trova l'ID_preventivo del record da eliminare
    sql = "SELECT ID_preventivo FROM preventivi WHERE Riferimento LIKE '" & riferimentoCerca & "%'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        idPreventivoDA = rs("ID_preventivo")
        rs.Close

        ' Elimina prima i record collegati in Tecnici preventivati
        sql = "DELETE FROM [Tecnici preventivati] WHERE ID_preventivo = " & idPreventivoDA
        db.Execute sql, dbFailOnError

        ' Elimina record collegati in Servizi preventivati
        sql = "DELETE FROM [Servizi preventivati] WHERE ID_preventivo = " & idPreventivoDA
        db.Execute sql, dbFailOnError

        ' Infine elimina il preventivo stesso
        sql = "DELETE FROM preventivi WHERE ID_preventivo = " & idPreventivoDA
        db.Execute sql, dbFailOnError
    Else
        rs.Close
    End If

    Set rs = Nothing
End Sub
