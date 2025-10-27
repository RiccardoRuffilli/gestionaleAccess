Option Compare Database
Option Explicit

' ============================================================
' ImportaPreventivoProduction.vba
' Modulo VBA per importazione preventivi da JSON - PRODUZIONE
' ============================================================

' ============================================================
' FUNZIONE PRINCIPALE PER PRODUZIONE
' ============================================================

Public Sub ImportaPreventivo()
    Dim scriptPath As String
    Dim jsonPath As String
    Dim cartellaDB As String
    Dim psCommand As String

    On Error GoTo ErrorHandler

    ' Determina la cartella del database Access
    cartellaDB = CurrentProject.Path
    scriptPath = cartellaDB & "\ImportaPreventivo.ps1"

    ' Verifica esistenza script
    If Dir(scriptPath) = "" Then
        MsgBox "ERRORE: Script PowerShell non trovato!" & vbCrLf & vbCrLf & _
               "Percorso atteso: " & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Assicurati che ImportaPreventivo.ps1 sia nella stessa cartella del database Access.", _
               vbCritical, "Importa Preventivo"
        Exit Sub
    End If

    ' Seleziona file JSON
    jsonPath = SelezionaFileJSON()
    If jsonPath = "" Then
        Exit Sub
    End If

    ' Conferma importazione
    If MsgBox("Importare il preventivo da:" & vbCrLf & vbCrLf & _
              Dir(jsonPath) & vbCrLf & vbCrLf & _
              "Procedere?", _
              vbYesNo + vbQuestion, "Conferma Importazione") = vbNo Then
        Exit Sub
    End If

    ' Crea comando PowerShell con pausa finale
    psCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command " & _
                """cd '" & cartellaDB & "'; " & _
                "& .\ImportaPreventivo.ps1 '" & jsonPath & "'; " & _
                "Write-Host ''; " & _
                "Write-Host 'Premi un tasto per chiudere...' -ForegroundColor Yellow; " & _
                "$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')" & """"

    ' Esegui in finestra visibile
    Shell psCommand, vbNormalFocus

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante l'importazione:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Errore"
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
