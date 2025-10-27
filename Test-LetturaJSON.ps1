# ============================================================
# Test-LetturaJSON.ps1
# Script per testare la lettura del file JSON
# Mostra un report dettagliato dei dati letti
# ============================================================

param(
    [string]$JsonFilePath
)

Write-Host "=== TEST LETTURA FILE JSON ===" -ForegroundColor Cyan
Write-Host ""

# Verifica parametro
if (-not $JsonFilePath) {
    Write-Host "[ERRORE] Specificare il percorso del file JSON" -ForegroundColor Red
    Write-Host ""
    Write-Host "Uso: .\Test-LetturaJSON.ps1 <percorso-file-json>" -ForegroundColor Yellow
    exit 1
}

Write-Host "File da leggere: $JsonFilePath" -ForegroundColor Yellow
Write-Host ""

# Verifica esistenza file
if (-not (Test-Path $JsonFilePath)) {
    Write-Host "[ERRORE] File non trovato!" -ForegroundColor Red
    exit 1
}

try {
    # Leggi file JSON
    Write-Host "Lettura file JSON..." -ForegroundColor Yellow
    $jsonContent = Get-Content -Path $JsonFilePath -Raw -Encoding UTF8

    Write-Host "[OK] File letto correttamente" -ForegroundColor Green
    Write-Host "  Dimensione: $($jsonContent.Length) caratteri" -ForegroundColor Gray
    Write-Host ""

    # Parse JSON
    Write-Host "Parsing JSON..." -ForegroundColor Yellow
    $data = $jsonContent | ConvertFrom-Json

    Write-Host "[OK] JSON parsato correttamente" -ForegroundColor Green
    Write-Host ""

    # ============================================================
    # REPORT DETTAGLIATO - SEZIONE EVENTO
    # ============================================================
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          DATI EVENTO (tabella: preventivi)               ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    $e = $data.evento

    Write-Host "ID originale:            $($e.id)" -ForegroundColor White
    Write-Host "Nome evento:             $($e.nome_evento)" -ForegroundColor White
    Write-Host "Cliente ID:              $($e.cliente_id)" -ForegroundColor White
    Write-Host "Responsabile ID:         $($e.responsabile_id)" -ForegroundColor White
    Write-Host "Data/ora inizio:         $($e.data_ora_inizio)" -ForegroundColor White
    Write-Host "Data/ora fine:           $($e.data_ora_fine)" -ForegroundColor White
    Write-Host "Luogo:                   $($e.luogo)" -ForegroundColor White
    Write-Host "Confermato:              $($e.flag_confermato)" -ForegroundColor White
    Write-Host "Data conferma:           $($e.data_conferma)" -ForegroundColor White
    Write-Host "Note generali:           $($e.note_generali)" -ForegroundColor White
    Write-Host "Note amministrazione:    $($e.note_amministrazione)" -ForegroundColor White
    Write-Host ""

    Write-Host "--- Dove verranno scritti questi dati: ---" -ForegroundColor Yellow
    Write-Host "Tabella: preventivi" -ForegroundColor Gray
    Write-Host "Campi mappati:" -ForegroundColor Gray
    Write-Host "  - ID_cliente           = $($e.cliente_id)" -ForegroundColor DarkGray
    Write-Host "  - id_referente_videorent = $($e.responsabile_id)" -ForegroundColor DarkGray
    Write-Host "  - Data inizio          = $(($e.data_ora_inizio -split ' ')[0])" -ForegroundColor DarkGray
    Write-Host "  - Ora inizio           = $(($e.data_ora_inizio -split ' ')[1])" -ForegroundColor DarkGray
    Write-Host "  - Data fine            = $(($e.data_ora_fine -split ' ')[0])" -ForegroundColor DarkGray
    Write-Host "  - Ora fine             = $(($e.data_ora_fine -split ' ')[1])" -ForegroundColor DarkGray
    Write-Host "  - luogo                = $($e.luogo)" -ForegroundColor DarkGray
    Write-Host "  - Confermato           = $(if ($e.flag_confermato -eq 1) { 'True' } else { 'False' })" -ForegroundColor DarkGray
    Write-Host "  - Data confermato      = $($e.data_conferma)" -ForegroundColor DarkGray
    Write-Host "  - Riferimento          = EVENTO_$($e.id)" -ForegroundColor DarkGray
    Write-Host "  - Note                 = $($e.note_generali)" -ForegroundColor DarkGray
    Write-Host "  - Note amministrazione = $($e.note_amministrazione)" -ForegroundColor DarkGray
    Write-Host ""

    # ============================================================
    # REPORT DETTAGLIATO - SEZIONE PERSONALE
    # ============================================================
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║      DATI PERSONALE (tabella: Tecnici preventivati)      ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    if ($data.personale -and $data.personale.Count -gt 0) {
        Write-Host "Numero di persone assegnate: $($data.personale.Count)" -ForegroundColor White
        Write-Host ""

        $counter = 1
        foreach ($persona in $data.personale) {
            Write-Host "--- PERSONA #$counter ---" -ForegroundColor Yellow
            Write-Host "ID originale:        $($persona.id)" -ForegroundColor White
            Write-Host "User ID (tecnico):   $($persona.user_id)" -ForegroundColor White
            Write-Host "Data inizio:         $($persona.data_inizio)" -ForegroundColor White
            Write-Host "Data fine:           $($persona.data_fine)" -ForegroundColor White
            Write-Host "Ora inizio:          $($persona.ora_inizio)" -ForegroundColor White
            Write-Host "Ora fine:            $($persona.ora_fine)" -ForegroundColor White
            Write-Host "Funzione:            $($persona.funzione)" -ForegroundColor White
            Write-Host "Note:                $($persona.note)" -ForegroundColor White
            Write-Host ""

            Write-Host "Verrà scritto in: Tecnici preventivati" -ForegroundColor Gray
            Write-Host "  - id_tecnico       = $($persona.user_id)" -ForegroundColor DarkGray
            Write-Host "  - Data inizio      = $($persona.data_inizio)" -ForegroundColor DarkGray
            Write-Host "  - Data fine        = $($persona.data_fine)" -ForegroundColor DarkGray
            Write-Host "  - Ora inizio       = $($persona.ora_inizio)" -ForegroundColor DarkGray
            Write-Host "  - Ora fine         = $($persona.ora_fine)" -ForegroundColor DarkGray
            Write-Host "  - Funzione         = $($persona.funzione)" -ForegroundColor DarkGray
            Write-Host "  - Note             = $($persona.note)" -ForegroundColor DarkGray
            Write-Host "  - Riferimento      = PERSONA_$($persona.id)" -ForegroundColor DarkGray
            Write-Host ""

            $counter++
        }
    } else {
        Write-Host "Nessun personale assegnato" -ForegroundColor Gray
        Write-Host ""
    }

    # ============================================================
    # REPORT DETTAGLIATO - SEZIONE SERVIZI
    # ============================================================
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║      DATI SERVIZI (tabella: Servizi preventivati)        ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    if ($data.servizi -and $data.servizi.Count -gt 0) {
        Write-Host "Numero di servizi/articoli: $($data.servizi.Count)" -ForegroundColor White
        Write-Host ""

        $counter = 1
        foreach ($servizio in $data.servizi) {
            Write-Host "--- SERVIZIO #$counter ---" -ForegroundColor Yellow
            Write-Host "ID originale:            $($servizio.id)" -ForegroundColor White
            Write-Host "Item ID (articolo):      $($servizio.item_id)" -ForegroundColor White
            Write-Host "Quantità:                $($servizio.qty)" -ForegroundColor White
            Write-Host "Prezzo unitario:         $($servizio.prezzo_unitario)" -ForegroundColor White
            Write-Host "Costo unitario:          $($servizio.costo_unitario)" -ForegroundColor White
            Write-Host "Sconto %:                $($servizio.sconto_percentuale)" -ForegroundColor White
            Write-Host "IVA %:                   $($servizio.iva_percentuale)" -ForegroundColor White
            Write-Host "Descrizione:             $($servizio.descrizione)" -ForegroundColor White
            Write-Host "Note:                    $($servizio.note)" -ForegroundColor White
            Write-Host ""

            # Calcola totale
            $subtotale = [decimal]$servizio.qty * [decimal]$servizio.prezzo_unitario
            $sconto = $subtotale * ([decimal]$servizio.sconto_percentuale / 100)
            $netto = $subtotale - $sconto
            $iva = $netto * ([decimal]$servizio.iva_percentuale / 100)
            $totale = $netto + $iva

            Write-Host "Calcoli:" -ForegroundColor Gray
            Write-Host "  Subtotale (qty × prezzo): $('{0:N2}' -f $subtotale) €" -ForegroundColor DarkGray
            Write-Host "  Sconto:                   $('{0:N2}' -f $sconto) €" -ForegroundColor DarkGray
            Write-Host "  Netto:                    $('{0:N2}' -f $netto) €" -ForegroundColor DarkGray
            Write-Host "  IVA:                      $('{0:N2}' -f $iva) €" -ForegroundColor DarkGray
            Write-Host "  Totale:                   $('{0:N2}' -f $totale) €" -ForegroundColor DarkGray
            Write-Host ""

            Write-Host "Verrà scritto in: Servizi preventivati" -ForegroundColor Gray
            Write-Host "  - id_servizio      = $($servizio.item_id)" -ForegroundColor DarkGray
            Write-Host "  - Quantita         = $($servizio.qty)" -ForegroundColor DarkGray
            Write-Host "  - Prezzo_unitario  = $($servizio.prezzo_unitario)" -ForegroundColor DarkGray
            Write-Host "  - Costo_unitario   = $($servizio.costo_unitario)" -ForegroundColor DarkGray
            Write-Host "  - Sconto_pct       = $($servizio.sconto_percentuale)" -ForegroundColor DarkGray
            Write-Host "  - Iva_pct          = $($servizio.iva_percentuale)" -ForegroundColor DarkGray
            Write-Host "  - Descrizione      = $($servizio.descrizione)" -ForegroundColor DarkGray
            Write-Host "  - Note             = $($servizio.note)" -ForegroundColor DarkGray
            Write-Host "  - Riferimento      = SERVIZIO_$($servizio.id)" -ForegroundColor DarkGray
            Write-Host ""

            $counter++
        }
    } else {
        Write-Host "Nessun servizio/articolo presente" -ForegroundColor Gray
        Write-Host ""
    }

    # ============================================================
    # RIEPILOGO FINALE
    # ============================================================
    Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                    RIEPILOGO FINALE                       ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "[OK] File JSON letto e parsato correttamente" -ForegroundColor Green
    Write-Host ""
    Write-Host "Record da creare:" -ForegroundColor White
    Write-Host "  - 1 preventivo nella tabella 'preventivi'" -ForegroundColor White
    Write-Host "  - $($data.personale.Count) record nella tabella 'Tecnici preventivati'" -ForegroundColor White
    Write-Host "  - $($data.servizi.Count) record nella tabella 'Servizi preventivati'" -ForegroundColor White
    Write-Host ""
    Write-Host "Il preventivo verrà identificato con: Riferimento = 'EVENTO_$($e.id)'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "=== TEST COMPLETATO CON SUCCESSO ===" -ForegroundColor Green

    exit 0

} catch {
    Write-Host ""
    Write-Host "[ERRORE] ERRORE DURANTE LA LETTURA!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Messaggio errore:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Dettagli completi:" -ForegroundColor Yellow
    Write-Host $_.Exception.ToString() -ForegroundColor Gray
    Write-Host ""

    exit 1
}
