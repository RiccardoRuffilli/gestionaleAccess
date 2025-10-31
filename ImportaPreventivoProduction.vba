Option Compare Database
Option Explicit

' ============================================================
' ImportaPreventivoProduction.vba
' Modulo VBA per importazione preventivi da JSON - PRODUZIONE
' ============================================================

' Costanti per la connessione al database
Private Const SQL_SERVER As String = "LENOVO-01\SQLEXPRESS"
Private Const SQL_DATABASE As String = "Videorent"
Private Const SQL_USER As String = "sa"
Private Const SQL_PASSWORD As String = "Video2009"

' ============================================================
' FUNZIONE PRINCIPALE PER PRODUZIONE
' ============================================================

Public Sub ImportaPreventivo()
    Dim jsonPath As String
    Dim preventivoID As Long
    Dim duplicateInfo As String
    Dim userResponse As VbMsgBoxResult

    On Error GoTo ErrorHandler

    ' Seleziona file JSON
    jsonPath = SelezionaFileJSON()
    If jsonPath = "" Then
        Exit Sub
    End If

    ' Leggi ID preventivo dal JSON
    preventivoID = LeggiIDDalJSON(jsonPath)
    If preventivoID = 0 Then
        MsgBox "Impossibile leggere l'ID del preventivo dal file JSON!", vbCritical, "Errore"
        Exit Sub
    End If

    ' Verifica se esiste già un preventivo con questo ID
    duplicateInfo = VerificaDuplicato(preventivoID)

    If duplicateInfo <> "" Then
        ' Preventivo esistente trovato - chiedi conferma sostituzione
        userResponse = MsgBox(duplicateInfo & vbCrLf & vbCrLf & "Desideri sostituirlo?", _
                             vbYesNo + vbQuestion, "Preventivo Esistente")

        If userResponse = vbNo Then
            MsgBox "Importazione annullata dall'utente.", vbInformation, "Annullato"
            Exit Sub
        End If

        ' Elimina preventivo esistente
        If Not EliminaPreventivo(preventivoID) Then
            MsgBox "Errore durante l'eliminazione del preventivo esistente!", vbCritical, "Errore"
            Exit Sub
        End If

        ' Procedi con importazione
        EseguiImportazionePowerShell jsonPath

    Else
        ' Nessun duplicato - chiedi conferma importazione
        userResponse = MsgBox("Importare preventivo n° " & preventivoID & "?", _
                             vbYesNo + vbQuestion, "Conferma Importazione")

        If userResponse = vbNo Then
            MsgBox "Importazione annullata dall'utente.", vbInformation, "Annullato"
            Exit Sub
        End If

        ' Procedi con importazione
        EseguiImportazionePowerShell jsonPath
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Errore"
End Sub

' ============================================================
' FUNZIONI DI CONNESSIONE DATABASE
' ============================================================

Private Function CreaConnessioneSQL() As Object
    Dim conn As Object
    Dim connString As String

    On Error GoTo ErrorHandler

    Set conn = CreateObject("ADODB.Connection")

    ' Connection string per SQL Server Authentication
    connString = "Provider=SQLOLEDB;Data Source=" & SQL_SERVER & ";" & _
                 "Initial Catalog=" & SQL_DATABASE & ";" & _
                 "User ID=" & SQL_USER & ";Password=" & SQL_PASSWORD & ";"

    conn.Open connString

    Set CreaConnessioneSQL = conn
    Exit Function

ErrorHandler:
    Set CreaConnessioneSQL = Nothing
    MsgBox "Errore connessione al database:" & vbCrLf & Err.Description, vbCritical, "Errore"
End Function

' ============================================================
' FUNZIONI DI VERIFICA E GESTIONE DUPLICATI
' ============================================================

Private Function VerificaDuplicato(preventivoID As Long) As String
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim nomeCliente As String
    Dim citta As String
    Dim dataAllest As String
    Dim risultato As String

    On Error GoTo ErrorHandler

    Set conn = CreaConnessioneSQL()
    If conn Is Nothing Then
        VerificaDuplicato = ""
        Exit Function
    End If

    ' Query per verificare esistenza e recuperare dettagli
    sql = "SELECT p.ID_preventivo, " & _
          "       c.Nome_azienda, " & _
          "       p.[Città], " & _
          "       p.[Data allestimento] " & _
          "FROM preventivi p " & _
          "LEFT JOIN [Anagrafica Clienti] c ON p.ID_cliente = c.ID_Cliente " & _
          "WHERE p.ID_preventivo = " & preventivoID

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    If Not rs.EOF Then
        ' Preventivo trovato - costruisci messaggio
        nomeCliente = IIf(IsNull(rs("Nome_azienda")), "Sconosciuto", rs("Nome_azienda"))
        citta = IIf(IsNull(rs("Città")), "", rs("Città"))

        If Not IsNull(rs("Data allestimento")) Then
            dataAllest = Format(rs("Data allestimento"), "dd/mm/yyyy")
        Else
            dataAllest = "Data non specificata"
        End If

        risultato = "Esiste già un preventivo n° " & preventivoID & _
                   " intestato a " & nomeCliente

        If citta <> "" Then
            risultato = risultato & " (" & citta & ")"
        End If

        risultato = risultato & " - allestimento " & dataAllest & "."
    Else
        risultato = ""
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    VerificaDuplicato = risultato
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    VerificaDuplicato = ""
End Function

Private Function EliminaPreventivo(preventivoID As Long) As Boolean
    Dim conn As Object
    Dim sql As String

    On Error GoTo ErrorHandler

    Set conn = CreaConnessioneSQL()
    If conn Is Nothing Then
        EliminaPreventivo = False
        Exit Function
    End If

    conn.BeginTrans

    ' Elimina tecnici preventivati
    sql = "DELETE FROM [Tecnici preventivati] WHERE ID_preventivo = " & preventivoID
    conn.Execute sql

    ' Elimina servizi preventivati
    sql = "DELETE FROM [Servizi preventivati] WHERE ID_preventivo = " & preventivoID
    conn.Execute sql

    ' Elimina preventivo
    sql = "DELETE FROM preventivi WHERE ID_preventivo = " & preventivoID
    conn.Execute sql

    conn.CommitTrans
    conn.Close
    Set conn = Nothing

    EliminaPreventivo = True
    Exit Function

ErrorHandler:
    If Not conn Is Nothing Then
        conn.RollbackTrans
        conn.Close
    End If
    EliminaPreventivo = False
    MsgBox "Errore durante l'eliminazione:" & vbCrLf & Err.Description, vbCritical, "Errore"
End Function

' ============================================================
' FUNZIONI DI LETTURA JSON
' ============================================================

Private Function LeggiIDDalJSON(jsonPath As String) As Long
    Dim fileNum As Integer
    Dim fileContent As String
    Dim startPos As Long
    Dim endPos As Long
    Dim idString As String
    Dim inEvento As Boolean

    On Error GoTo ErrorHandler

    ' Leggi tutto il contenuto del file
    fileNum = FreeFile
    Open jsonPath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum

    ' Cerca "evento": {
    startPos = InStr(1, fileContent, """evento""")
    If startPos = 0 Then
        LeggiIDDalJSON = 0
        Exit Function
    End If

    ' Cerca "id": dopo "evento"
    startPos = InStr(startPos, fileContent, """id""")
    If startPos = 0 Then
        LeggiIDDalJSON = 0
        Exit Function
    End If

    ' Trova il valore dopo i due punti
    startPos = InStr(startPos, fileContent, ":")
    If startPos = 0 Then
        LeggiIDDalJSON = 0
        Exit Function
    End If

    ' Salta spazi e cerca la prima cifra
    startPos = startPos + 1
    Do While Mid(fileContent, startPos, 1) = " " Or Mid(fileContent, startPos, 1) = vbTab
        startPos = startPos + 1
    Loop

    ' Leggi il numero fino alla virgola o al ritorno a capo
    endPos = startPos
    Do While IsNumeric(Mid(fileContent, endPos, 1))
        endPos = endPos + 1
    Loop

    idString = Mid(fileContent, startPos, endPos - startPos)
    LeggiIDDalJSON = CLng(idString)

    Exit Function

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    LeggiIDDalJSON = 0
End Function

' ============================================================
' FUNZIONI DI ESECUZIONE POWERSHELL
' ============================================================

Private Sub EseguiImportazionePowerShell(jsonPath As String)
    Dim cartellaDB As String
    Dim scriptPath As String
    Dim psCommand As String

    On Error GoTo ErrorHandler

    cartellaDB = CurrentProject.Path
    scriptPath = cartellaDB & "\ImportaPreventivo.ps1"

    ' Verifica esistenza script
    If Dir(scriptPath) = "" Then
        MsgBox "ERRORE: Script PowerShell non trovato!" & vbCrLf & vbCrLf & _
               "Percorso atteso: " & vbCrLf & scriptPath, _
               vbCritical, "Importa Preventivo"
        Exit Sub
    End If

    ' Crea comando PowerShell
    psCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command " & _
                """cd '" & cartellaDB & "'; " & _
                "& .\ImportaPreventivo.ps1 '" & jsonPath & "'; " & _
                "Write-Host ''; " & _
                "Write-Host 'Premi un tasto per chiudere...' -ForegroundColor Yellow; " & _
                "$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"" "

    ' Esegui in finestra visibile
    Shell psCommand, vbNormalFocus

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'esecuzione PowerShell:" & vbCrLf & Err.Description, _
           vbCritical, "Errore"
End Sub

' ============================================================
' FUNZIONI DI UTILITA'
' ============================================================

Private Function SelezionaFileJSON() As String
    Dim fileDialog As Object
    Dim selectedFile As String

    On Error GoTo ErrorHandler

    ' Crea finestra di selezione file
    Set fileDialog = Application.FileDialog(3) ' msoFileDialogFilePicker

    With fileDialog
        .Title = "Seleziona file JSON da importare"
        .Filters.Clear
        .Filters.Add "File JSON", "*.json"
        .AllowMultiSelect = False

        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            selectedFile = ""
        End If
    End With

    Set fileDialog = Nothing
    SelezionaFileJSON = selectedFile
    Exit Function

ErrorHandler:
    SelezionaFileJSON = ""
End Function
