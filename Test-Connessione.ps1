# ============================================================
# Test-Connessione.ps1
# Script per testare la connessione al database SQL Server
# ============================================================

param(
    [string]$ServerInstance = "LENOVO-01\SQLEXPRESS",
    [string]$DatabaseName = "Videorent-b",
    [string]$SqlUser = "",
    [string]$SqlPassword = ""
)

Write-Host "=== TEST CONNESSIONE DATABASE ===" -ForegroundColor Cyan
Write-Host ""

# Costruisci connection string
if ($SqlUser -and $SqlPassword) {
    # SQL Server Authentication
    $connectionString = "Server=$ServerInstance;Database=$DatabaseName;User Id=$SqlUser;Password=$SqlPassword;"
    Write-Host "Modalita: SQL Server Authentication" -ForegroundColor Yellow
    Write-Host "Utente: $SqlUser" -ForegroundColor Yellow
} else {
    # Windows Authentication
    $connectionString = "Server=$ServerInstance;Database=$DatabaseName;Integrated Security=True;"
    Write-Host "Modalita: Windows Authentication" -ForegroundColor Yellow
}

Write-Host "Server: $ServerInstance" -ForegroundColor Yellow
Write-Host "Database: $DatabaseName" -ForegroundColor Yellow
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

    Write-Host "[OK] CONNESSIONE RIUSCITA!" -ForegroundColor Green
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
    Write-Host "[OK] ACCESSO ALLE TABELLE VERIFICATO!" -ForegroundColor Green
    Write-Host ""

    # Chiudi connessione
    $connection.Close()

    Write-Host "=== TEST COMPLETATO CON SUCCESSO ===" -ForegroundColor Green
    exit 0

} catch {
    Write-Host ""
    Write-Host "[ERRORE] ERRORE DURANTE IL TEST!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Messaggio errore:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Dettagli completi:" -ForegroundColor Yellow
    Write-Host $_.Exception.ToString() -ForegroundColor Gray
    Write-Host ""

    # Suggerimenti
    Write-Host "POSSIBILI CAUSE:" -ForegroundColor Yellow
    Write-Host "1. SQL Server Express non e' avviato" -ForegroundColor Gray
    Write-Host "2. Il nome dell'istanza e' errato (attuale: $ServerInstance)" -ForegroundColor Gray
    Write-Host "3. Il database non esiste o ha un nome diverso" -ForegroundColor Gray
    Write-Host "4. Credenziali errate (username/password)" -ForegroundColor Gray
    Write-Host "5. L'utente non ha permessi sul database" -ForegroundColor Gray
    Write-Host ""

    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
    }

    exit 1
}
