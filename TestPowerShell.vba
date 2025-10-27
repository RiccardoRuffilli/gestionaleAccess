Option Compare Database
Option Explicit

' ============================================================
' TestPowerShell.vba
' Modulo VBA per testare gli script PowerShell passo-passo
' ============================================================

' API per visualizzare finestra DOS
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

Const SW_SHOWNORMAL = 1

' ============================================================
' PASSO 1: TEST CONNESSIONE DATABASE
' ============================================================

Public Sub Passo1_TestConnessione()
    Dim scriptPath As String
    Dim cartellaLavoro As String
    Dim psCommand As String

    On Error GoTo ErrorHandler

    ' Determina la cartella di lavoro
    cartellaLavoro = Environ("USERPROFILE") & "\Downloads\_importazione_preventivi"
    scriptPath = cartellaLavoro & "\Test-Connessione.ps1"

    ' Verifica esistenza script
    If Dir(scriptPath) = "" Then
        MsgBox "ERRORE: File non trovato!" & vbCrLf & vbCrLf & _
               "Percorso atteso: " & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Assicurati di aver copiato Test-Connessione.ps1 nella cartella _importazione_preventivi", _
               vbCritical, "Test Connessione"
        Exit Sub
    End If

    ' Mostra messaggio
    MsgBox "PASSO 1: Test Connessione Database" & vbCrLf & vbCrLf & _
           "Si aprirà una finestra PowerShell che mostrerà:" & vbCrLf & _
           "• Se riesce a connettersi al database" & vbCrLf & _
           "• Quanti record ci sono nelle tabelle" & vbCrLf & _
           "• Eventuali errori di connessione" & vbCrLf & vbCrLf & _
           "La finestra rimarrà aperta - premi un tasto per chiuderla.", _
           vbInformation, "Test Connessione"

    ' Crea comando PowerShell con pausa finale
    psCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command " & _
                """cd '" & cartellaLavoro & "'; " & _
                "& .\Test-Connessione.ps1; " & _
                "Write-Host ''; " & _
                "Write-Host 'Premi un tasto per chiudere...' -ForegroundColor Yellow; " & _
                "$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')" & """"

    ' Esegui in finestra visibile
    Shell psCommand, vbNormalFocus

    Exit Sub

ErrorHandler:
    MsgBox "Errore nell'esecuzione del test:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Errore"
End Sub

' ============================================================
' PASSO 2: TEST LETTURA JSON
' ============================================================

Public Sub Passo2_TestLetturaJSON()
    Dim scriptPath As String
    Dim jsonPath As String
    Dim cartellaLavoro As String
    Dim psCommand As String
    Dim fileDialog As Object

    On Error GoTo ErrorHandler

    ' Determina la cartella di lavoro
    cartellaLavoro = Environ("USERPROFILE") & "\Downloads\_importazione_preventivi"
    scriptPath = cartellaLavoro & "\Test-LetturaJSON.ps1"

    ' Verifica esistenza script
    If Dir(scriptPath) = "" Then
        MsgBox "ERRORE: File non trovato!" & vbCrLf & vbCrLf & _
               "Percorso atteso: " & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Assicurati di aver copiato Test-LetturaJSON.ps1 nella cartella _importazione_preventivi", _
               vbCritical, "Test Lettura JSON"
        Exit Sub
    End If

    ' Seleziona file JSON
    jsonPath = SelezionaFileJSON(cartellaLavoro)
    If jsonPath = "" Then
        Exit Sub
    End If

    ' Mostra messaggio
    MsgBox "PASSO 2: Test Lettura JSON" & vbCrLf & vbCrLf & _
           "Si aprirà una finestra PowerShell che mostrerà:" & vbCrLf & _
           "• I dati letti dal file JSON" & vbCrLf & _
           "• Dove verranno scritti nel database" & vbCrLf & _
           "• Un riepilogo completo dell'importazione" & vbCrLf & vbCrLf & _
           "La finestra rimarrà aperta - premi un tasto per chiuderla.", _
           vbInformation, "Test Lettura JSON"

    ' Crea comando PowerShell con pausa finale
    psCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command " & _
                """cd '" & cartellaLavoro & "'; " & _
                "& .\Test-LetturaJSON.ps1 '" & jsonPath & "'; " & _
                "Write-Host ''; " & _
                "Write-Host 'Premi un tasto per chiudere...' -ForegroundColor Yellow; " & _
                "$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')" & """"

    ' Esegui in finestra visibile
    Shell psCommand, vbNormalFocus

    Exit Sub

ErrorHandler:
    MsgBox "Errore nell'esecuzione del test:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Errore"
End Sub

' ============================================================
' PASSO 3: IMPORTAZIONE COMPLETA
' ============================================================

Public Sub Passo3_ImportazioneCompleta()
    Dim scriptPath As String
    Dim jsonPath As String
    Dim cartellaLavoro As String
    Dim psCommand As String

    On Error GoTo ErrorHandler

    ' Determina la cartella di lavoro
    cartellaLavoro = Environ("USERPROFILE") & "\Downloads\_importazione_preventivi"
    scriptPath = cartellaLavoro & "\ImportaPreventivo.ps1"

    ' Verifica esistenza script
    If Dir(scriptPath) = "" Then
        MsgBox "ERRORE: File non trovato!" & vbCrLf & vbCrLf & _
               "Percorso atteso: " & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Assicurati di aver copiato ImportaPreventivo.ps1 nella cartella _importazione_preventivi", _
               vbCritical, "Importazione"
        Exit Sub
    End If

    ' Conferma prima di procedere
    If MsgBox("Hai eseguito i test precedenti (Passo 1 e Passo 2)?" & vbCrLf & vbCrLf & _
              "È importante verificare la connessione e la lettura JSON" & vbCrLf & _
              "prima di procedere con l'importazione reale.", _
              vbYesNo + vbQuestion, "Conferma") = vbNo Then
        MsgBox "Esegui prima:" & vbCrLf & _
               "• Passo1_TestConnessione" & vbCrLf & _
               "• Passo2_TestLetturaJSON", vbInformation
        Exit Sub
    End If

    ' Seleziona file JSON
    jsonPath = SelezionaFileJSON(cartellaLavoro)
    If jsonPath = "" Then
        Exit Sub
    End If

    ' Mostra messaggio
    MsgBox "PASSO 3: Importazione Completa" & vbCrLf & vbCrLf & _
           "Si aprirà una finestra PowerShell che mostrerà:" & vbCrLf & _
           "• Tutte le operazioni di importazione" & vbCrLf & _
           "• Eventuali errori che si verificano" & vbCrLf & _
           "• Il risultato finale" & vbCrLf & vbCrLf & _
           "La finestra rimarrà aperta - premi un tasto per chiuderla.", _
           vbInformation, "Importazione"

    ' Crea comando PowerShell con pausa finale
    psCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command " & _
                """cd '" & cartellaLavoro & "'; " & _
                "& .\ImportaPreventivo.ps1 '" & jsonPath & "'; " & _
                "Write-Host ''; " & _
                "Write-Host 'Premi un tasto per chiudere...' -ForegroundColor Yellow; " & _
                "$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')" & """"

    ' Esegui in finestra visibile
    Shell psCommand, vbNormalFocus

    Exit Sub

ErrorHandler:
    MsgBox "Errore nell'esecuzione:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Errore"
End Sub

' ============================================================
' FUNZIONI DI UTILITA'
' ============================================================

Private Function SelezionaFileJSON(cartellaDefault As String) As String
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
        .InitialFileName = cartellaDefault & "\"

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
