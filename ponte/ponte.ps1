# ============================================================
# PONTE PC PER RAPPORTINO DI CANTIERE (GC-CodeLab)
# Piccolo servizio locale su 127.0.0.1:47821: legge e scrive i file
# Excel e l'archivio rapportini senza passare dai permessi di Chrome.
# Non richiede installazioni: usa solo Windows e .NET integrati.
# ============================================================
param(
  [switch]$Configura
)

$ErrorActionPreference = 'Stop'
$PORTA = 47821
$VERSIONE = '1.0'
$CARTELLA_SCRIPT = Split-Path -Parent $MyInvocation.MyCommand.Path
$FILE_SCRIPT = Join-Path $CARTELLA_SCRIPT 'ponte.ps1'
$FILE_CONFIG = Join-Path $CARTELLA_SCRIPT 'config.json'
$LNK_AVVIO = Join-Path ([Environment]::GetFolderPath('Startup')) 'Ponte Rapportino.lnk'

$ORIGINI_AMMESSE = @(
  'https://rapportino.gccodelab.it',
  'https://rapportino-vicla.web.app',
  'https://giannivicla74-jpg.github.io',
  'http://127.0.0.1:5173',
  'http://localhost:5173',
  'http://127.0.0.1:5174',
  'http://localhost:5174'
)

function Leggi-Config {
  if (-not (Test-Path -LiteralPath $FILE_CONFIG)) {
    return [pscustomobject]@{ excelDir = ''; archivioDir = '' }
  }
  try {
    $raw = [System.IO.File]::ReadAllText($FILE_CONFIG, [System.Text.Encoding]::UTF8)
    $obj = $raw | ConvertFrom-Json
    return [pscustomobject]@{
      excelDir    = [string]($obj.excelDir)
      archivioDir = [string]($obj.archivioDir)
    }
  } catch {
    return [pscustomobject]@{ excelDir = ''; archivioDir = '' }
  }
}

function Salva-Config([string]$excelDir, [string]$archivioDir) {
  $json = ([ordered]@{
    excelDir    = $excelDir.Trim()
    archivioDir = $archivioDir.Trim()
  }) | ConvertTo-Json
  $utf8 = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($FILE_CONFIG, $json, $utf8)
}

function Test-AvvioAutomatico {
  return [bool](Test-Path -LiteralPath $LNK_AVVIO)
}

function Imposta-AvvioAutomatico([bool]$attivo) {
  if ($attivo) {
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($LNK_AVVIO)
    $sc.TargetPath = 'powershell.exe'
    $sc.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$FILE_SCRIPT`""
    $sc.WorkingDirectory = $CARTELLA_SCRIPT
    $sc.WindowStyle = 7
    $sc.Description = 'Ponte PC per Rapportino di Cantiere'
    $sc.Save()
  } else {
    if (Test-Path -LiteralPath $LNK_AVVIO) {
      Remove-Item -LiteralPath $LNK_AVVIO -Force -ErrorAction SilentlyContinue
    }
  }
}

# Registra il collegamento ponte-rapportino:// solo per l'utente corrente (senza permessi amministratore)
# così il pulsante "Avvia Ponte PC" dentro l'app può riaccendere il Ponte con un clic
function Registra-Protocollo {
  try {
    $chiave = 'HKCU:\Software\Classes\ponte-rapportino'
    $cmdKey = "$chiave\shell\open\command"
    if (-not (Test-Path -LiteralPath $cmdKey)) {
      New-Item -Path $cmdKey -Force | Out-Null
    }
    Set-ItemProperty -LiteralPath $chiave -Name '(default)' -Value 'URL:Ponte Rapportino di Cantiere'
    Set-ItemProperty -LiteralPath $chiave -Name 'URL Protocol' -Value ''
    $cmd = "powershell.exe -WindowStyle Minimized -NoProfile -ExecutionPolicy Bypass -File `"$FILE_SCRIPT`""
    Set-ItemProperty -LiteralPath $cmdKey -Name '(default)' -Value $cmd
  } catch { }
}

function Scegli-CartellaDialogo([System.Windows.Forms.IWin32Window]$owner, [string]$titolo, [string]$iniziale) {
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = $titolo
  $dlg.ShowNewFolderButton = $true
  if ($iniziale -and (Test-Path -LiteralPath $iniziale)) {
    $dlg.SelectedPath = $iniziale
  }
  $esito = $dlg.ShowDialog($owner)
  $scelta = if ($esito -eq [System.Windows.Forms.DialogResult]::OK) { $dlg.SelectedPath } else { $null }
  $dlg.Dispose()
  return $scelta
}

function Mostra-Configurazione {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()

  $cfg = Leggi-Config

  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'Ponte Rapportino - Scegli le cartelle'
  $form.Size = New-Object System.Drawing.Size(640, 330)
  $form.StartPosition = 'CenterScreen'
  $form.FormBorderStyle = 'FixedDialog'
  $form.MaximizeBox = $false
  $form.MinimizeBox = $false
  $form.TopMost = $true
  $form.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)

  $lblIntro = New-Object System.Windows.Forms.Label
  $lblIntro.Text = 'Scegli la cartella dei file Excel e quella dell''archivio (anche su NAS, es. \\NAS\Presenze). Puoi usare Sfoglia oppure incollare il percorso.'
  $lblIntro.Location = New-Object System.Drawing.Point(16, 14)
  $lblIntro.Size = New-Object System.Drawing.Size(596, 38)
  $form.Controls.Add($lblIntro)

  $lblExcel = New-Object System.Windows.Forms.Label
  $lblExcel.Text = 'Cartella file Excel (Master_Operai e Presenze mensili):'
  $lblExcel.Location = New-Object System.Drawing.Point(16, 62)
  $lblExcel.Size = New-Object System.Drawing.Size(596, 20)
  $form.Controls.Add($lblExcel)

  $txtExcel = New-Object System.Windows.Forms.TextBox
  $txtExcel.Location = New-Object System.Drawing.Point(16, 84)
  $txtExcel.Size = New-Object System.Drawing.Size(496, 26)
  $txtExcel.Text = $cfg.excelDir
  $form.Controls.Add($txtExcel)

  $btnExcel = New-Object System.Windows.Forms.Button
  $btnExcel.Text = 'Sfoglia...'
  $btnExcel.Location = New-Object System.Drawing.Point(520, 82)
  $btnExcel.Size = New-Object System.Drawing.Size(92, 28)
  $btnExcel.Add_Click({
    $scelta = Scegli-CartellaDialogo $form 'Scegli la cartella dei file Excel' $txtExcel.Text
    if ($scelta) { $txtExcel.Text = $scelta }
  })
  $form.Controls.Add($btnExcel)

  $lblArch = New-Object System.Windows.Forms.Label
  $lblArch.Text = 'Cartella Archivio rapportini (copie dei rapportini convalidati):'
  $lblArch.Location = New-Object System.Drawing.Point(16, 126)
  $lblArch.Size = New-Object System.Drawing.Size(596, 20)
  $form.Controls.Add($lblArch)

  $txtArch = New-Object System.Windows.Forms.TextBox
  $txtArch.Location = New-Object System.Drawing.Point(16, 148)
  $txtArch.Size = New-Object System.Drawing.Size(496, 26)
  $txtArch.Text = $cfg.archivioDir
  $form.Controls.Add($txtArch)

  $btnArch = New-Object System.Windows.Forms.Button
  $btnArch.Text = 'Sfoglia...'
  $btnArch.Location = New-Object System.Drawing.Point(520, 146)
  $btnArch.Size = New-Object System.Drawing.Size(92, 28)
  $btnArch.Add_Click({
    $scelta = Scegli-CartellaDialogo $form 'Scegli la cartella dell''Archivio rapportini' $txtArch.Text
    if ($scelta) { $txtArch.Text = $scelta }
  })
  $form.Controls.Add($btnArch)

  $chkAvvio = New-Object System.Windows.Forms.CheckBox
  $chkAvvio.Text = 'Avvia automaticamente il Ponte all''accensione del PC (ridotto a icona)'
  $chkAvvio.Location = New-Object System.Drawing.Point(16, 192)
  $chkAvvio.Size = New-Object System.Drawing.Size(596, 26)
  $chkAvvio.Checked = Test-AvvioAutomatico
  $form.Controls.Add($chkAvvio)

  $salvato = $false

  $btnSalva = New-Object System.Windows.Forms.Button
  $btnSalva.Text = 'Salva'
  $btnSalva.Location = New-Object System.Drawing.Point(416, 240)
  $btnSalva.Size = New-Object System.Drawing.Size(94, 32)
  $btnSalva.Add_Click({
    $eDir = $txtExcel.Text.Trim()
    $aDir = $txtArch.Text.Trim()
    if (-not $eDir -and -not $aDir) {
      [System.Windows.Forms.MessageBox]::Show($form, 'Scegli almeno una cartella prima di salvare.', 'Ponte Rapportino', 'OK', 'Warning') | Out-Null
      return
    }
    if ($eDir -and -not (Test-Path -LiteralPath $eDir -PathType Container)) {
      [System.Windows.Forms.MessageBox]::Show($form, "La cartella Excel non esiste o non e raggiungibile:`r`n$eDir", 'Ponte Rapportino', 'OK', 'Error') | Out-Null
      return
    }
    if ($aDir -and -not (Test-Path -LiteralPath $aDir -PathType Container)) {
      [System.Windows.Forms.MessageBox]::Show($form, "La cartella Archivio non esiste o non e raggiungibile:`r`n$aDir", 'Ponte Rapportino', 'OK', 'Error') | Out-Null
      return
    }
    Salva-Config $eDir $aDir
    Imposta-AvvioAutomatico $chkAvvio.Checked
    Registra-Protocollo
    $script:configSalvata = $true
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
  })
  $form.Controls.Add($btnSalva)

  $btnAnnulla = New-Object System.Windows.Forms.Button
  $btnAnnulla.Text = 'Annulla'
  $btnAnnulla.Location = New-Object System.Drawing.Point(518, 240)
  $btnAnnulla.Size = New-Object System.Drawing.Size(94, 32)
  $btnAnnulla.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $form.Controls.Add($btnAnnulla)

  $form.AcceptButton = $btnSalva
  $form.CancelButton = $btnAnnulla
  $script:configSalvata = $false
  $form.ShowDialog() | Out-Null
  $form.Dispose()
  return $script:configSalvata
}

if ($Configura) {
  if (Mostra-Configurazione) {
    Write-Host 'Cartelle salvate correttamente.' -ForegroundColor Green
  } else {
    Write-Host 'Modifica annullata.' -ForegroundColor Yellow
  }
  exit 0
}

# Se manca la configurazione alla prima apertura, apre subito la scelta cartelle
$cfgIniziale = Leggi-Config
if (-not $cfgIniziale.excelDir -and -not $cfgIniziale.archivioDir) {
  Write-Host 'Prima configurazione: scegli le cartelle nella finestra che si e aperta...' -ForegroundColor Cyan
  if (-not (Mostra-Configurazione)) {
    Write-Host 'Configurazione annullata. Avvia di nuovo quando sei pronto.' -ForegroundColor Yellow
    exit 0
  }
}

# Registra il protocollo ponte-rapportino:// anche all'avvio normale (cosi funziona il tasto nell'app)
Registra-Protocollo

# Verifica se un altro Ponte e gia attivo sulla stessa porta
try {
  $reqTest = [System.Net.WebRequest]::Create("http://127.0.0.1:$PORTA/stato")
  $reqTest.Timeout = 400
  $resTest = $reqTest.GetResponse()
  $resTest.Close()
  Write-Host ''
  Write-Host '============================================================' -ForegroundColor Yellow
  Write-Host '  Il Ponte Rapportino e GIA ATTIVO in un''altra finestra!' -ForegroundColor Yellow
  Write-Host '============================================================' -ForegroundColor Yellow
  Write-Host '  Puoi chiudere questa finestra: l''app e gia collegata.'
  Write-Host ''
  Start-Sleep -Seconds 4
  exit 0
} catch { }

function Nome-Cartella([string]$percorso) {
  if (-not $percorso) { return '' }
  $pulito = $percorso.TrimEnd('\', '/')
  $nome = Split-Path -Leaf $pulito
  if (-not $nome) { return $pulito }
  return $nome
}

function Test-NomeExcelSicuro([string]$nome, [bool]$permettiLock = $false) {
  if ([string]::IsNullOrWhiteSpace($nome)) { return $false }
  if ($nome -match '[\\/:*?"<>|]' -or $nome.Contains('..')) { return $false }
  if ($permettiLock -and $nome.StartsWith('~$')) {
    return ($nome -match '\.(xlsm|xlsx)$')
  }
  if ($nome.StartsWith('~$')) { return $false }
  return ($nome -match '\.(xlsm|xlsx)$')
}

function Test-NomeArchivioSicuro([string]$nome) {
  if ([string]::IsNullOrWhiteSpace($nome)) { return $false }
  if ($nome -match '[\\/:*?"<>|]' -or $nome.Contains('..')) { return $false }
  return ($nome -match '\.html$')
}

function Risolvi-Sotto([string]$baseDir, [string[]]$parti) {
  $basePieno = [System.IO.Path]::GetFullPath($baseDir).TrimEnd('\')
  $corrente = $basePieno
  foreach ($p in $parti) {
    if ([string]::IsNullOrWhiteSpace($p) -or $p -eq '.' -or $p -eq '..' -or $p -match '[\\/:*?"<>|]') {
      throw 'Percorso non valido'
    }
    $corrente = Join-Path $corrente $p
  }
  $finale = [System.IO.Path]::GetFullPath($corrente)
  if (-not $finale.StartsWith($basePieno + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Percorso fuori dalla cartella configurata'
  }
  return $finale
}

function Leggi-CorpoBytes($req) {
  $ms = New-Object System.IO.MemoryStream
  try {
    $req.InputStream.CopyTo($ms)
    return $ms.ToArray()
  } finally {
    $ms.Dispose()
  }
}

function Invia-Json($res, [int]$codice, $dati) {
  $json = $dati | ConvertTo-Json -Compress -Depth 6
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $res.StatusCode = $codice
  $res.ContentType = 'application/json; charset=utf-8'
  $res.ContentLength64 = $bytes.Length
  $res.OutputStream.Write($bytes, 0, $bytes.Length)
  $res.OutputStream.Close()
}

function Invia-Bytes($res, [int]$codice, [byte[]]$bytes, [string]$tipo = 'application/octet-stream') {
  $res.StatusCode = $codice
  $res.ContentType = $tipo
  $res.ContentLength64 = $bytes.Length
  $res.OutputStream.Write($bytes, 0, $bytes.Length)
  $res.OutputStream.Close()
}

function Applica-Cors($req, $res) {
  $orig = $req.Headers['Origin']
  $res.Headers.Add('Cache-Control', 'no-store')
  if ($orig) {
    if ($ORIGINI_AMMESSE -notcontains $orig) { return $false }
    $res.Headers.Add('Access-Control-Allow-Origin', $orig)
    $res.Headers.Add('Vary', 'Origin')
    $res.Headers.Add('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    $res.Headers.Add('Access-Control-Allow-Headers', 'Content-Type')
    $res.Headers.Add('Access-Control-Allow-Private-Network', 'true')
  }
  return $true
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$PORTA/")
try {
  $listener.Start()
} catch {
  Write-Host "ERRORE: impossibile avviare il Ponte sulla porta $PORTA." -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
  exit 1
}

$cfg = Leggi-Config
$cfgStamp = if (Test-Path -LiteralPath $FILE_CONFIG) { (Get-Item -LiteralPath $FILE_CONFIG).LastWriteTimeUtc } else { [datetime]::MinValue }

Clear-Host
Write-Host '============================================================' -ForegroundColor Green
Write-Host "  PONTE RAPPORTINO DI CANTIERE ATTIVO (v$VERSIONE)" -ForegroundColor Green
Write-Host '============================================================' -ForegroundColor Green
Write-Host ("  Cartella Excel:    " + $(if ($cfg.excelDir) { $cfg.excelDir } else { '(non impostata)' }))
Write-Host ("  Cartella Archivio: " + $(if ($cfg.archivioDir) { $cfg.archivioDir } else { '(non impostata)' }))
Write-Host ''
Write-Host '  Puoi ridurre a icona questa finestra mentre lavori.' -ForegroundColor Cyan
Write-Host '  Per spegnere il Ponte, chiudi questa finestra con la X.' -ForegroundColor Yellow
Write-Host '============================================================' -ForegroundColor Green
Write-Host ''

try {
  while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response

    try {
      # Ricarica config.json automaticamente se modificato da Scegli_Cartelle_Ponte.bat
      if (Test-Path -LiteralPath $FILE_CONFIG) {
        $nuovoStamp = (Get-Item -LiteralPath $FILE_CONFIG).LastWriteTimeUtc
        if ($nuovoStamp -ne $cfgStamp) {
          $cfg = Leggi-Config
          $cfgStamp = $nuovoStamp
          Write-Host ("[" + (Get-Date -Format 'HH:mm:ss') + "] Configurazione aggiornata.") -ForegroundColor Cyan
        }
      }

      if (-not (Applica-Cors $req $res)) {
        Invia-Json $res 403 @{ errore = 'Origine non consentita' }
        continue
      }

      if ($req.HttpMethod -eq 'OPTIONS') {
        $res.StatusCode = 204
        $res.OutputStream.Close()
        continue
      }

      $percorsoUrl = $req.Url.AbsolutePath
      $metodo = $req.HttpMethod

      # Richieste senza header Origin consentite solo su GET /stato (controllo locale)
      if (-not $req.Headers['Origin'] -and -not ($metodo -eq 'GET' -and $percorsoUrl -eq '/stato')) {
        Invia-Json $res 403 @{ errore = 'Richiesta non consentita' }
        continue
      }

      if ($metodo -eq 'GET' -and $percorsoUrl -eq '/stato') {
        $excelOk = [bool]($cfg.excelDir -and (Test-Path -LiteralPath $cfg.excelDir -PathType Container))
        $archOk  = [bool]($cfg.archivioDir -and (Test-Path -LiteralPath $cfg.archivioDir -PathType Container))
        Invia-Json $res 200 @{
          ok             = $true
          versione       = $VERSIONE
          excelPronta    = $excelOk
          excelNome      = if ($excelOk) { Nome-Cartella $cfg.excelDir } else { '' }
          archivioPronta = $archOk
          archivioNome   = if ($archOk) { Nome-Cartella $cfg.archivioDir } else { '' }
        }
        continue
      }

      # -------- EXCEL --------
      if ($percorsoUrl.StartsWith('/excel/')) {
        if (-not $cfg.excelDir -or -not (Test-Path -LiteralPath $cfg.excelDir -PathType Container)) {
          Invia-Json $res 400 @{ errore = 'Cartella dei file Excel non impostata o non raggiungibile nel Ponte PC' }
          continue
        }

        if ($metodo -eq 'GET' -and $percorsoUrl -eq '/excel/lista') {
          $nomi = @(Get-ChildItem -LiteralPath $cfg.excelDir -File | ForEach-Object { $_.Name })
          Invia-Json $res 200 @{ files = $nomi }
          continue
        }

        if ($metodo -eq 'GET' -and $percorsoUrl -eq '/excel/esiste') {
          $nome = $req.QueryString['nome']
          if (-not (Test-NomeExcelSicuro $nome $true)) {
            Invia-Json $res 400 @{ errore = 'Nome file non valido' }
            continue
          }
          $full = Risolvi-Sotto $cfg.excelDir @($nome)
          Invia-Json $res 200 @{ esiste = [bool](Test-Path -LiteralPath $full -PathType Leaf) }
          continue
        }

        if ($metodo -eq 'GET' -and $percorsoUrl -eq '/excel/leggi') {
          $nome = $req.QueryString['nome']
          if (-not (Test-NomeExcelSicuro $nome $false)) {
            Invia-Json $res 400 @{ errore = 'Nome file non valido' }
            continue
          }
          $full = Risolvi-Sotto $cfg.excelDir @($nome)
          if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
            Invia-Json $res 404 @{ errore = "File non trovato: $nome" }
            continue
          }
          $bytes = [System.IO.File]::ReadAllBytes($full)
          Invia-Bytes $res 200 $bytes
          continue
        }

        if ($metodo -eq 'POST' -and $percorsoUrl -eq '/excel/scrivi') {
          $nome = $req.QueryString['nome']
          $crea = ($req.QueryString['crea'] -eq '1')
          if (-not (Test-NomeExcelSicuro $nome $false)) {
            Invia-Json $res 400 @{ errore = 'Nome file non valido' }
            continue
          }
          $full = Risolvi-Sotto $cfg.excelDir @($nome)
          if (-not $crea -and -not (Test-Path -LiteralPath $full -PathType Leaf)) {
            Invia-Json $res 404 @{ errore = "File non trovato: $nome" }
            continue
          }
          $bytes = Leggi-CorpoBytes $req
          if ($bytes.Length -eq 0) {
            Invia-Json $res 400 @{ errore = 'Contenuto vuoto' }
            continue
          }
          $tmp = Join-Path $cfg.excelDir ("__ponte_tmp_" + [guid]::NewGuid().ToString('N') + '.tmp')
          try {
            [System.IO.File]::WriteAllBytes($tmp, $bytes)
            [System.IO.File]::Copy($tmp, $full, $true)
          } catch [System.IO.IOException] {
            Invia-Json $res 409 @{ errore = 'il file e in uso (forse aperto in Excel): chiudilo e riprova' }
            continue
          } finally {
            if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
          }
          Write-Host ("[" + (Get-Date -Format 'HH:mm:ss') + "] Scritto Excel: $nome") -ForegroundColor Green
          Invia-Json $res 200 @{ ok = $true }
          continue
        }

        if ($metodo -eq 'GET' -and $percorsoUrl -eq '/excel/backup-oggi') {
          $nome = $req.QueryString['nome']
          $label = $req.QueryString['label']
          if (-not $label) { $label = 'BACKUP' }
          if (-not (Test-NomeExcelSicuro $nome $false) -or $label -match '[^A-Za-z0-9_-]') {
            Invia-Json $res 400 @{ errore = 'Parametri backup non validi' }
            continue
          }
          $backupDir = Join-Path $cfg.excelDir 'Backup'
          if (-not (Test-Path -LiteralPath $backupDir -PathType Container)) {
            Invia-Json $res 200 @{ presente = $false }
            continue
          }
          $base = [System.IO.Path]::GetFileNameWithoutExtension($nome)
          $oggi = Get-Date -Format 'yyyy-MM-dd'
          $prefisso = "${base}_${label}_${oggi}_"
          $trovato = $false
          foreach ($f in (Get-ChildItem -LiteralPath $backupDir -File)) {
            if ($f.Name.StartsWith($prefisso)) { $trovato = $true; break }
          }
          Invia-Json $res 200 @{ presente = $trovato }
          continue
        }

        if ($metodo -eq 'POST' -and $percorsoUrl -eq '/excel/backup') {
          $nome = $req.QueryString['nome']
          $label = $req.QueryString['label']
          if (-not $label) { $label = 'BACKUP' }
          $tieni = 20
          if ($req.QueryString['tieni'] -match '^\d+$') { $tieni = [Math]::Max(1, [int]$req.QueryString['tieni']) }
          if (-not (Test-NomeExcelSicuro $nome $false) -or $label -match '[^A-Za-z0-9_-]') {
            Invia-Json $res 400 @{ errore = 'Parametri backup non validi' }
            continue
          }
          $backupDir = Join-Path $cfg.excelDir 'Backup'
          if (-not (Test-Path -LiteralPath $backupDir -PathType Container)) {
            [System.IO.Directory]::CreateDirectory($backupDir) | Out-Null
          }
          $base = [System.IO.Path]::GetFileNameWithoutExtension($nome)
          $ext  = [System.IO.Path]::GetExtension($nome)
          $stamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
          $esistenti = @(Get-ChildItem -LiteralPath $backupDir -File | ForEach-Object { $_.Name })
          $backupName = "${base}_${label}_${stamp}${ext}"
          $n = 2
          while ($esistenti -contains $backupName) {
            $backupName = "${base}_${label}_${stamp}_${n}${ext}"
            $n++
          }
          $bytes = Leggi-CorpoBytes $req
          $dest = Risolvi-Sotto $backupDir @($backupName)
          [System.IO.File]::WriteAllBytes($dest, $bytes)

          $prefisso = "${base}_${label}_"
          $tutti = @($esistenti + $backupName | Where-Object { $_.StartsWith($prefisso) -and $_.EndsWith($ext) } | Sort-Object)
          $daTogliere = [Math]::Max(0, $tutti.Count - $tieni)
          for ($i = 0; $i -lt $daTogliere; $i++) {
            $vecchio = Join-Path $backupDir $tutti[$i]
            Remove-Item -LiteralPath $vecchio -Force -ErrorAction SilentlyContinue
          }
          Write-Host ("[" + (Get-Date -Format 'HH:mm:ss') + "] Backup creato: Backup\$backupName") -ForegroundColor DarkGreen
          Invia-Json $res 200 @{ backupName = $backupName }
          continue
        }
      }

      # -------- ARCHIVIO --------
      if ($percorsoUrl.StartsWith('/archivio/')) {
        if (-not $cfg.archivioDir -or -not (Test-Path -LiteralPath $cfg.archivioDir -PathType Container)) {
          Invia-Json $res 400 @{ errore = 'Cartella dell''archivio non impostata o non raggiungibile nel Ponte PC' }
          continue
        }

        if ($metodo -eq 'GET' -and $percorsoUrl -eq '/archivio/esiste') {
          $perc = $req.QueryString['percorso']
          $nome = $req.QueryString['nome']
          if (-not $perc -or -not (Test-NomeArchivioSicuro $nome)) {
            Invia-Json $res 400 @{ errore = 'Parametri archivio non validi' }
            continue
          }
          $parti = @($perc.Split('/') | Where-Object { $_ -ne '' }) + @($nome)
          $full = Risolvi-Sotto $cfg.archivioDir $parti
          $ok = (Test-Path -LiteralPath $full -PathType Leaf) -and ((Get-Item -LiteralPath $full).Length -gt 0)
          Invia-Json $res 200 @{ esiste = [bool]$ok }
          continue
        }

        if ($metodo -eq 'POST' -and $percorsoUrl -eq '/archivio/salva') {
          $perc = $req.QueryString['percorso']
          $nome = $req.QueryString['nome']
          if (-not $perc -or -not (Test-NomeArchivioSicuro $nome)) {
            Invia-Json $res 400 @{ errore = 'Parametri archivio non validi' }
            continue
          }
          $partiDir = @($perc.Split('/') | Where-Object { $_ -ne '' })
          $dirDest = Risolvi-Sotto $cfg.archivioDir $partiDir
          if (-not (Test-Path -LiteralPath $dirDest -PathType Container)) {
            [System.IO.Directory]::CreateDirectory($dirDest) | Out-Null
          }
          $full = Risolvi-Sotto $cfg.archivioDir ($partiDir + @($nome))
          $bytes = Leggi-CorpoBytes $req
          [System.IO.File]::WriteAllBytes($full, $bytes)
          $riletto = [System.IO.File]::ReadAllBytes($full)
          if ($riletto.Length -ne $bytes.Length) {
            Invia-Json $res 500 @{ errore = 'il file salvato non corrisponde: riprova' }
            continue
          }
          Write-Host ("[" + (Get-Date -Format 'HH:mm:ss') + "] Salvato archivio: $perc/$nome") -ForegroundColor Green
          Invia-Json $res 200 @{ ok = $true }
          continue
        }
      }

      Invia-Json $res 404 @{ errore = 'Comando non riconosciuto' }
    } catch {
      try {
        Invia-Json $res 500 @{ errore = $_.Exception.Message }
      } catch { }
    }
  }
} finally {
  $listener.Stop()
  $listener.Close()
}
