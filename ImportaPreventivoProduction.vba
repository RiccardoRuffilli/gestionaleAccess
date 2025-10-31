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
    Dim outputPath As String
    Dim exitCode As Integer
    Dim duplicateInfo As String
    Dim replaceExisting As Boolean

    On Error GoTo ErrorHandler

    ' Determina la cartella del database Access
    cartellaDB = CurrentProject.Path
    scriptPath = cartellaDB & "\ImportaPreventivo.ps1"
    outputPath = Environ("TEMP") & "\preventivo_import_output.txt"

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

    ' Prima chiamata: verifica duplicati
    replaceExisting = False
    exitCode = EseguiImportPowerShell(scriptPath, jsonPath, outputPath, replaceExisting)

    ' Controlla exit code
    If exitCode = 2 Then
        ' Duplicato trovato - leggi i dettagli e chiedi conferma
        duplicateInfo = LeggiDettagliDuplicato(outputPath)

        If duplicateInfo <> "" Then
            If MsgBox(duplicateInfo & vbCrLf & vbCrLf & "Desideri sostituirlo?", _
                      vbYesNo + vbQuestion, "Preventivo Esistente") = vbYes Then
                ' Utente ha confermato - riesegui con -ReplaceExisting
                replaceExisting = True
                exitCode = EseguiImportPowerShell(scriptPath, jsonPath, outputPath, replaceExisting)

                ' Mostra risultato finale
                If exitCode = 0 Then
                    MsgBox "Preventivo sostituito con successo!", vbInformation, "Importazione Completata"
                End If
            Else
                MsgBox "Importazione annullata dall'utente.", vbInformation, "Annullato"
            End If
        End If
    ElseIf exitCode = 0 Then
        MsgBox "Importazione completata con successo!", vbInformation, "Importazione Completata"
    End If

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

Private Function EseguiImportPowerShell(scriptPath As String, jsonPath As String, _
                                        outputPath As String, replaceExisting As Boolean) As Integer
    Dim wsh As Object
    Dim psCommand As String
    Dim replaceParam As String
    Dim exitCode As Integer

    On Error GoTo ErrorHandler

    Set wsh = CreateObject("WScript.Shell")

    ' Parametro opzionale per sostituzione
    If replaceExisting Then
        replaceParam = " -ReplaceExisting"
    Else
        replaceParam = ""
    End If

    ' Costruisci comando PowerShell che salva output su file
    psCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command " & _
                """cd '" & CurrentProject.Path & "'; " & _
                "& .\ImportaPreventivo.ps1 '" & jsonPath & "'" & replaceParam & " *>&1 | Tee-Object -FilePath '" & outputPath & "'; " & _
                "Write-Host ''; " & _
                "Write-Host 'Premi un tasto per chiudere...' -ForegroundColor Yellow; " & _
                "$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); " & _
                "exit $LASTEXITCODE" & """"

    ' Esegui e attendi completamento (usa Run con Wait=False per mostrare finestra)
    ' Nota: con Shell non possiamo catturare exit code, ma PowerShell salva su file
    wsh.Run psCommand, 1, True  ' 1 = finestra normale, True = aspetta

    ' Purtroppo con WScript.Shell.Run non possiamo catturare l'exit code facilmente
    ' Leggiamo l'exit code dall'output del PowerShell
    exitCode = LeggiExitCode(outputPath)

    Set wsh = Nothing
    EseguiImportPowerShell = exitCode
    Exit Function

ErrorHandler:
    EseguiImportPowerShell = -1
End Function

Private Function LeggiExitCode(outputPath As String) As Integer
    Dim fileNum As Integer
    Dim lineText As String

    On Error Resume Next

    ' Cerca nel file per determinare l'exit code
    If Dir(outputPath) <> "" Then
        fileNum = FreeFile
        Open outputPath For Input As #fileNum

        Do While Not EOF(fileNum)
            Line Input #fileNum, lineText

            ' Cerca marker di duplicato
            If InStr(lineText, "DUPLICATO_TROVATO") > 0 Then
                Close #fileNum
                LeggiExitCode = 2
                Exit Function
            End If

            ' Cerca marker di successo
            If InStr(lineText, "IMPORTAZIONE COMPLETATA") > 0 Then
                Close #fileNum
                LeggiExitCode = 0
                Exit Function
            End If
        Loop

        Close #fileNum
    End If

    ' Default: assume errore se non trovato marker
    LeggiExitCode = 1
End Function

Private Function LeggiDettagliDuplicato(outputPath As String) As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim prevId As String
    Dim clienteName As String
    Dim citta As String
    Dim dataAllest As String
    Dim inDuplicateSection As Boolean

    On Error Resume Next

    prevId = ""
    clienteName = ""
    citta = ""
    dataAllest = ""
    inDuplicateSection = False

    If Dir(outputPath) <> "" Then
        fileNum = FreeFile
        Open outputPath For Input As #fileNum

        Do While Not EOF(fileNum)
            Line Input #fileNum, lineText

            ' Cerca i marker di inizio sezione duplicato
            If InStr(lineText, "DUPLICATO_TROVATO") > 0 Then
                inDuplicateSection = True
            ElseIf inDuplicateSection Then
                ' Estrai i dettagli
                If InStr(lineText, "ID:") > 0 Then
                    prevId = Trim(Mid(lineText, InStr(lineText, "ID:") + 3))
                    prevId = Replace(prevId, "[WARN]", "")
                    prevId = Trim(prevId)
                ElseIf InStr(lineText, "CLIENTE:") > 0 Then
                    clienteName = Trim(Mid(lineText, InStr(lineText, "CLIENTE:") + 8))
                    clienteName = Replace(clienteName, "[WARN]", "")
                    clienteName = Trim(clienteName)
                ElseIf InStr(lineText, "CITTA:") > 0 Then
                    citta = Trim(Mid(lineText, InStr(lineText, "CITTA:") + 6))
                    citta = Replace(citta, "[WARN]", "")
                    citta = Trim(citta)
                ElseIf InStr(lineText, "DATA:") > 0 Then
                    dataAllest = Trim(Mid(lineText, InStr(lineText, "DATA:") + 5))
                    dataAllest = Replace(dataAllest, "[WARN]", "")
                    dataAllest = Trim(dataAllest)
                    ' Fine sezione
                    Exit Do
                End If
            End If
        Loop

        Close #fileNum

        ' Costruisci messaggio
        If prevId <> "" Then
            LeggiDettagliDuplicato = "Esiste già un preventivo n° " & prevId & _
                                     " intestato a " & clienteName
            If citta <> "" Then
                LeggiDettagliDuplicato = LeggiDettagliDuplicato & " (" & citta & ")"
            End If
            LeggiDettagliDuplicato = LeggiDettagliDuplicato & " - allestimento " & dataAllest & "."
        Else
            LeggiDettagliDuplicato = ""
        End If
    End If
End Function
