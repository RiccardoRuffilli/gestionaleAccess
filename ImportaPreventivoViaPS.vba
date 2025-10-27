Option Compare Database
Option Explicit

' ==============================================================================
' MODULO: ImportaPreventivoViaPS
' DESCRIZIONE: Importa preventivi JSON usando PowerShell (semplice e affidabile)
' AUTORE: Generato da Claude Code
' DATA: 2025-10-27
' ==============================================================================

' Costanti
Private Const CARTELLA_LAVORO As String = "_importazione_preventivi"
Private Const SCRIPT_PS_NAME As String = "ImportaPreventivo.ps1"

' ==============================================================================
' FUNZIONE PRINCIPALE
' ==============================================================================

Public Sub ImportaPreventivoViaPS()
    ' Importa un preventivo JSON usando script PowerShell

    Dim jsonFilePath As String
    Dim cartellaLavoro As String
    Dim scriptPath As String

    ' Ottieni percorso cartella di lavoro
    cartellaLavoro = GetCartellaLavoro()
    If cartellaLavoro = "" Then
        MsgBox "Impossibile trovare o creare la cartella di lavoro.", vbCritical, "Errore"
        Exit Sub
    End If

    ' Verifica esistenza script PowerShell
    scriptPath = cartellaLavoro & SCRIPT_PS_NAME
    If Dir(scriptPath) = "" Then
        MsgBox "Script PowerShell non trovato!" & vbCrLf & vbCrLf & _
               "Assicurati che il file '" & SCRIPT_PS_NAME & "' sia presente in:" & vbCrLf & _
               cartellaLavoro, vbCritical, "Errore"
        Exit Sub
    End If

    ' Messaggio iniziale
    MsgBox "Seleziona il file JSON da importare.", vbInformation, "Importazione Preventivo"

    ' Selezione file JSON
    jsonFilePath = SelezionaFileJSON(cartellaLavoro)
    If jsonFilePath = "" Then
        MsgBox "Operazione annullata.", vbExclamation, "Annullato"
        Exit Sub
    End If

    ' Conferma
    Dim risposta As VbMsgBoxResult
    risposta = MsgBox("Importare il preventivo da:" & vbCrLf & vbCrLf & _
                      jsonFilePath & vbCrLf & vbCrLf & _
                      "Continuare?", vbYesNo + vbQuestion, "Conferma")

    If risposta = vbNo Then
        Exit Sub
    End If

    ' Mostra messaggio "In elaborazione..."
    DoCmd.Hourglass True

    ' Esegui script PowerShell
    Dim exitCode As Long
    Dim output As String

    exitCode = EseguiScriptPowerShell(scriptPath, jsonFilePath, output)

    DoCmd.Hourglass False

    ' Mostra risultato
    If exitCode = 0 Then
        MsgBox "Importazione completata con successo!" & vbCrLf & vbCrLf & _
               "Dettagli:" & vbCrLf & output, vbInformation, "Successo"
    Else
        MsgBox "Errore durante l'importazione." & vbCrLf & vbCrLf & _
               "Codice errore: " & exitCode & vbCrLf & vbCrLf & _
               "Output:" & vbCrLf & output, vbCritical, "Errore"
    End If

End Sub

' ==============================================================================
' FUNZIONI DI ESECUZIONE POWERSHELL
' ==============================================================================

Private Function EseguiScriptPowerShell(scriptPath As String, jsonPath As String, ByRef output As String) As Long
    ' Esegue lo script PowerShell e cattura l'output

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    ' Crea file temporaneo per l'output
    Dim tempOutputFile As String
    tempOutputFile = Environ("TEMP") & "\ImportaPreventivo_" & Format(Now(), "yyyymmddhhnnss") & ".txt"

    ' Costruisci comando PowerShell
    ' -NoProfile: Non caricare profilo utente (più veloce)
    ' -ExecutionPolicy Bypass: Permetti esecuzione script
    ' -File: Esegui file script
    Dim psCommand As String
    psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & _
                scriptPath & """ """ & jsonPath & """ > """ & tempOutputFile & """ 2>&1"

    ' Esegui e attendi completamento
    Dim exitCode As Long
    exitCode = wsh.Run(psCommand, 0, True) ' 0 = nascosto, True = attendi

    ' Leggi output
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempOutputFile) Then
        Dim file As Object
        Set file = fso.OpenTextFile(tempOutputFile, 1)
        output = file.ReadAll
        file.Close

        ' Elimina file temporaneo
        fso.DeleteFile tempOutputFile
    Else
        output = "Nessun output disponibile"
    End If
    On Error GoTo 0

    EseguiScriptPowerShell = exitCode

    Set wsh = Nothing
    Set fso = Nothing
End Function

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
        On Error Resume Next
        fso.CreateFolder percorso
        If Err.Number <> 0 Then
            CreaCartellaSeNonEsiste = False
            Exit Function
        End If
        On Error GoTo 0
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
