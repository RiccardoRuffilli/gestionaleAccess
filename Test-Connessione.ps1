# ============================================================
# Test-Connessione.ps1
# Script per testare la connessione al database SQL Server
# ============================================================

Write-Host "=== TEST CONNESSIONE DATABASE ===" -ForegroundColor Cyan
Write-Host ""

# Configurazione connessione database
$serverInstance = ".\SQLEXPRESS"
$databaseName = "Videorent-b"
$connectionString = "Server=$serverInstance;Database=$databaseName;Integrated Security=True;"

Write-Host "Server: $serverInstance" -ForegroundColor Yellow
Write-Host "Database: $databaseName" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "Tentativo di connessione..." -ForegroundColor Yellow

    # Carica assembly per SQL Server
    Add-Type -AssemblyName "System.Data"

    # Crea connessione
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString

    # Apri connessione
    $connection.Open()

    Write-Host "✓ CONNESSIONE RIUSCITA!" -ForegroundColor Green
    Write-Host ""

    # Test query per verificare accesso tabelle
    Write-Host "Verifica accesso alle tabelle..." -ForegroundColor Yellow

    $testCmd = $connection.CreateCommand()
    $testCmd.CommandText = "SELECT COUNT(*) FROM preventivi"
    $countPreventivi = $testCmd.ExecuteScalar()
    Write-Host "  - Tabella 'preventivi': $countPreventivi record trovati" -ForegroundColor Gray

    $testCmd.CommandText = "SELECT COUNT(*) FROM [Tecnici preventivati]"
    $countTecnici = $testCmd.ExecuteScalar()
    Write-Host "  - Tabella 'Tecnici preventivati': $countTecnici record trovati" -ForegroundColor Gray

    $testCmd.CommandText = "SELECT COUNT(*) FROM [Servizi preventivati]"
    $countServizi = $testCmd.ExecuteScalar()
    Write-Host "  - Tabella 'Servizi preventivati': $countServizi record trovati" -ForegroundColor Gray

    Write-Host ""
    Write-Host "✓ ACCESSO ALLE TABELLE VERIFICATO!" -ForegroundColor Green
    Write-Host ""

    # Chiudi connessione
    $connection.Close()

    Write-Host "=== TEST COMPLETATO CON SUCCESSO ===" -ForegroundColor Green
    exit 0

} catch {
    Write-Host ""
    Write-Host "✗ ERRORE DURANTE IL TEST!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Messaggio errore:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Dettagli completi:" -ForegroundColor Yellow
    Write-Host $_.Exception.ToString() -ForegroundColor Gray
    Write-Host ""

    # Suggerimenti
    Write-Host "POSSIBILI CAUSE:" -ForegroundColor Yellow
    Write-Host "1. SQL Server Express non è avviato" -ForegroundColor Gray
    Write-Host "2. Il nome dell'istanza non è '.\SQLEXPRESS'" -ForegroundColor Gray
    Write-Host "3. Il database non si chiama 'Videorent-b'" -ForegroundColor Gray
    Write-Host "4. L'utente Windows non ha permessi sul database" -ForegroundColor Gray
    Write-Host ""

    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
    }

    exit 1
}
