#Requires -Version 5.1
<#
.SYNOPSIS
    Liest die Filialnummern aus der Excel "Filial_Rufnummer", sucht sie im Telekom-Portal
    unter "Meine Vertraege" und traegt die gefundene Mobilrufnummer in die Excel ein.

.DESCRIPTION
    Treffer  (Kostenstelle == Filiale)  -> Rufnummer eintragen, Zeile GRUEN
    Kein Treffer                        -> Zeile ROT
    Mehrere Rufnummern / unklar         -> alle eintragen bzw. nichts eintragen, Zeile GELB
    Schon eine Rufnummer vorhanden      -> Zeile wird NICHT angefasst

    Sicherheit:
    - Vor jedem echten Lauf wird eine Sicherheitskopie der Excel angelegt.
    - Auf der Website wird NICHTS angeklickt. Das Skript tippt nur die Filialnummer
      ins Suchfeld, drueckt Enter und liest die Seite aus.
    - Kein Passwort im Skript: Du loggst dich selbst in Chrome ein.

.EXAMPLE
    .\Telekom-Rufnummern-Abgleich.ps1 -Testlauf
    Sucht nur die ersten 3 Filialen und zeigt das Ergebnis an. Die Excel wird NICHT veraendert.

.EXAMPLE
    .\Telekom-Rufnummern-Abgleich.ps1
    Echter Lauf ueber alle Filialen.

.EXAMPLE
    .\Telekom-Rufnummern-Abgleich.ps1 -Diagnose
    Schreibt eine Datei mit dem Aufbau der Portal-Seite (Suchfelder, Tabellen-Ueberschriften,
    keine Vertragsdaten). Hilfreich, wenn das Suchfeld oder die Tabelle nicht erkannt wird.
#>
[CmdletBinding()]
param(
    [string]$Ordner        = 'C:\Users\dario.mihaljevic\OneDrive - TEDi GmbH & Co. KG\IT-Einkauf\Mobilfunk\Excels',
    [string]$DateiName     = 'Filial_Rufnummer',
    [string]$SpalteFiliale = 'Filiale',
    [string]$SpalteRufnummer = 'Mobilrufnummer',

    # Nur die ersten X Filialen suchen, Excel wird nicht veraendert
    [switch]$Testlauf,
    [int]$TestAnzahl = 3,

    # Nur den Seitenaufbau des Portals in eine Datei schreiben
    [switch]$Diagnose,

    # Optional: CSS-Selektor fuer das Suchfeld, falls es nicht automatisch erkannt wird
    [string]$SuchfeldSelektor = '',

    # Optional: Adresse, die beim Start in Chrome geoeffnet wird
    [string]$PortalUrl = 'about:blank',

    [int]$Port = 9222,
    [int]$MaxWarteSekunden = 15
)

$ErrorActionPreference = 'Stop'
$script:Basis = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:Zeitstempel = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:LogDatei = $null
$script:Ws = $null
$script:CdpId = 0

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------
function Write-Log {
    param([string]$Text, [ConsoleColor]$Farbe = 'Gray')
    $zeile = '{0:HH:mm:ss}  {1}' -f (Get-Date), $Text
    Write-Host $zeile -ForegroundColor $Farbe
    if ($script:LogDatei) { Add-Content -LiteralPath $script:LogDatei -Value $zeile -Encoding UTF8 }
}

function Confirm-JaNein([string]$Frage) {
    while ($true) {
        $a = (Read-Host "$Frage (j/n)").Trim().ToLower()
        if ($a -eq 'j' -or $a -eq 'ja') { return $true }
        if ($a -eq 'n' -or $a -eq 'nein') { return $false }
    }
}

# Excel erwartet Farben als BGR-Zahl
function Get-ExcelFarbe([int]$R, [int]$G, [int]$B) { return $R + ($G * 256) + ($B * 65536) }
$FarbeGruen = Get-ExcelFarbe 198 239 206
$FarbeRot   = Get-ExcelFarbe 255 199 206
$FarbeGelb  = Get-ExcelFarbe 255 235 156

# ---------------------------------------------------------------------------
# Chrome steuern (DevTools-Protokoll, ohne Selenium / Treiber / Module)
# ---------------------------------------------------------------------------
function Find-Chrome {
    $pfade = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $pfade) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
    throw 'Chrome wurde nicht gefunden.'
}

function Test-ChromePort {
    try { $null = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json/version" -TimeoutSec 2; return $true }
    catch { return $false }
}

function Start-SteuerbaresChrome {
    if (Test-ChromePort) { Write-Log 'Steuerbares Chrome laeuft bereits.'; return }
    $chrome = Find-Chrome
    # Eigenes Profil: Chrome erlaubt die Steuerung nicht mit dem normalen Profil.
    # Dein normales Chrome bleibt davon unberuehrt.
    $profil = Join-Path $env:LOCALAPPDATA 'TelekomRufnummernSkript\ChromeProfil'
    New-Item -ItemType Directory -Path $profil -Force | Out-Null
    $argumente = @(
        "--remote-debugging-port=$Port",
        "--user-data-dir=`"$profil`"",
        '--no-first-run',
        '--no-default-browser-check',
        $PortalUrl
    )
    Start-Process -FilePath $chrome -ArgumentList $argumente | Out-Null
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Milliseconds 500
        if (Test-ChromePort) { return }
    }
    throw 'Chrome konnte nicht im Steuer-Modus gestartet werden.'
}

function Select-PortalTab {
    $tabs = @(Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json/list" -TimeoutSec 5 |
        Where-Object { $_.type -eq 'page' -and $_.url -notmatch '^(chrome|devtools|chrome-extension|about):' })
    if ($tabs.Count -eq 0) { throw 'Kein geoeffneter Tab mit einer Website gefunden.' }
    $tab = $tabs | Where-Object { ($_.title + ' ' + $_.url) -match 'vertr' } | Select-Object -First 1
    if (-not $tab) { $tab = $tabs | Where-Object { $_.url -match 'telekom' } | Select-Object -First 1 }
    if (-not $tab) { $tab = $tabs[0] }
    return $tab
}

function Connect-Cdp([string]$WsUrl) {
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
    $null = $ws.ConnectAsync([Uri]$WsUrl, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
    $script:Ws = $ws
}

function Invoke-Cdp([string]$Methode, [hashtable]$Parameter = @{}) {
    $script:CdpId++
    $id = $script:CdpId
    $json = @{ id = $id; method = $Methode; params = $Parameter } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $null = $script:Ws.SendAsync([ArraySegment[byte]]::new($bytes),
        [Net.WebSockets.WebSocketMessageType]::Text, $true,
        [Threading.CancellationToken]::None).GetAwaiter().GetResult()

    $puffer = New-Object byte[] 65536
    $cts = New-Object Threading.CancellationTokenSource ([TimeSpan]::FromSeconds(60))
    while ($true) {
        $ms = New-Object IO.MemoryStream
        do {
            $r = $script:Ws.ReceiveAsync([ArraySegment[byte]]::new($puffer), $cts.Token).GetAwaiter().GetResult()
            if ($r.MessageType -eq [Net.WebSockets.WebSocketMessageType]::Close) {
                throw 'Die Verbindung zu Chrome wurde geschlossen (Tab oder Chrome zu?).'
            }
            $ms.Write($puffer, 0, $r.Count)
        } while (-not $r.EndOfMessage)
        $nachricht = [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
        if ($nachricht.id -eq $id) {
            if ($nachricht.error) { throw "Chrome-Fehler bei ${Methode}: $($nachricht.error.message)" }
            return $nachricht.result
        }
    }
}

# Fuehrt JavaScript im Tab aus. Das JavaScript gibt immer JSON-Text zurueck.
function Invoke-Js([string]$Code) {
    $res = Invoke-Cdp 'Runtime.evaluate' @{ expression = $Code; returnByValue = $true }
    if ($res.exceptionDetails) {
        throw "JavaScript-Fehler auf der Seite: $($res.exceptionDetails.text) $($res.exceptionDetails.exception.description)"
    }
    if ($null -eq $res.result.value) { return $null }
    return ($res.result.value | ConvertFrom-Json)
}

function Send-Enter {
    $null = Invoke-Cdp 'Input.dispatchKeyEvent' @{ type = 'keyDown'; key = 'Enter'; code = 'Enter'; windowsVirtualKeyCode = 13; nativeVirtualKeyCode = 13; text = "`r" }
    $null = Invoke-Cdp 'Input.dispatchKeyEvent' @{ type = 'keyUp'; key = 'Enter'; code = 'Enter'; windowsVirtualKeyCode = 13; nativeVirtualKeyCode = 13 }
}

# ---------------------------------------------------------------------------
# JavaScript, das in der Seite laeuft (liest nur, klickt nie)
# ---------------------------------------------------------------------------
$JsGemeinsam = @'
function __docs(){var d=[document];document.querySelectorAll('iframe').forEach(function(f){try{if(f.contentDocument&&f.contentDocument.body)d.push(f.contentDocument);}catch(e){}});return d;}
function __vis(el){return !!(el.offsetWidth||el.offsetHeight||el.getClientRects().length);}
function __norm(s){return (s||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();}
function __txt(el){return __norm(el.innerText||el.textContent);}
function __cells(row){return Array.prototype.slice.call(row.cells?row.cells:row.querySelectorAll('[role=cell],[role=gridcell],[role=columnheader],[role=rowheader]'));}
'@

# Sucht das Suchfeld, setzt den Fokus darauf und markiert den Inhalt
$JsSuchfeld = $JsGemeinsam + @'
(function(sel){
  var best=null,bestScore=0;
  __docs().forEach(function(d){
    var list=sel?d.querySelectorAll(sel):d.querySelectorAll('input,textarea');
    Array.prototype.forEach.call(list,function(el){
      if(!__vis(el)||el.disabled||el.readOnly)return;
      var t=(el.type||'').toLowerCase();
      if(!sel&&['hidden','checkbox','radio','button','submit','reset','file','image','password','date','email'].indexOf(t)>=0)return;
      var s=sel?10:0;
      if(t==='search')s+=3;
      var meta=[el.placeholder,el.getAttribute('aria-label'),el.name,el.id,el.title].join(' ').toLowerCase();
      if(/such|search|filter|rufnummer|kostenstelle/.test(meta))s+=5;
      if(el.closest&&el.closest('header,nav,[role=banner],[role=navigation]'))s-=4;
      if(s>bestScore){best=el;bestScore=s;}
    });
  });
  if(!best)return JSON.stringify({ok:false});
  window.__tkSuchfeld=best;
  best.focus();
  if(best.value){
    var desc=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(best),'value');
    if(desc&&desc.set)desc.set.call(best,'');else best.value='';
    best.dispatchEvent(new Event('input',{bubbles:true}));
  }
  if(best.select)best.select();
  var info=[best.placeholder,best.getAttribute('aria-label'),best.name,best.id].filter(function(x){return x;}).join(' | ');
  return JSON.stringify({ok:best.ownerDocument.activeElement===best,info:info,pfad:location.origin+location.pathname});
})
'@

# Prueft vor dem Enter: Fokus im Suchfeld und genau die Filialnummer steht drin
$JsPruefen = @'
(function(v){
  var el=window.__tkSuchfeld;
  if(!el||!el.isConnected)return JSON.stringify({ok:false,wert:''});
  return JSON.stringify({ok:(el.ownerDocument.activeElement===el&&el.value===v),wert:el.value,pfad:location.origin+location.pathname});
})
'@

# Liest Treffer aus: Zeilen, deren Kostenstelle EXAKT der Filialnummer entspricht
$JsAuslesen = $JsGemeinsam + @'
(function(kst){
  var out={modus:'none',treffer:[],kstGefunden:false};
  var phoneRe=/(?:\+49|0049|\b0)[\s\-\/()]*1[5-7]\d(?:[\s\-\/]*\d){6,9}/g;
  function phones(t){var m=__norm(t).match(phoneRe);return m?m.map(__norm):[];}
  function kstWert(t){t=__norm(t);var m=t.match(/^kostenstelle\s*:?\s*(.*)$/i);return m?__norm(m[1]):t;}
  var res=[];
  __docs().forEach(function(d){
    d.querySelectorAll('table,[role=table],[role=grid],[role=treegrid]').forEach(function(t){
      if(!__vis(t))return;
      var rows=Array.prototype.slice.call(t.tagName==='TABLE'?t.rows:t.querySelectorAll('[role=row]'));
      var kIdx=-1,pIdx=-1,hdr=-1;
      for(var i=0;i<rows.length&&i<5&&kIdx<0;i++){
        __cells(rows[i]).forEach(function(c,j){
          var h=__txt(c).toLowerCase();
          if(kIdx<0&&h.indexOf('kostenstelle')>=0){kIdx=j;hdr=i;}
          if(pIdx<0&&/rufnummer|mobilfunknummer|msisdn|telefonnummer/.test(h))pIdx=j;
        });
      }
      if(kIdx<0)return;
      out.modus='tabelle';
      for(var r=hdr+1;r<rows.length;r++){
        var cells=__cells(rows[r]);
        if(cells.length<=kIdx)continue;
        if(kstWert(__txt(cells[kIdx]))!==kst)continue;
        out.kstGefunden=true;
        var p=(pIdx>=0&&cells[pIdx])?phones(__txt(cells[pIdx])):[];
        if(!p.length&&pIdx>=0&&cells[pIdx]){var raw=__txt(cells[pIdx]);if(raw.replace(/\D/g,'').length>=8)p=[raw];}
        if(!p.length)p=phones(__txt(rows[r]));
        res=res.concat(p);
      }
    });
  });
  if(out.modus==='none'){
    __docs().forEach(function(d){
      d.querySelectorAll('body *').forEach(function(el){
        if(el.children.length>2||!__vis(el))return;
        if(el.closest('th,thead,[role=columnheader]'))return;
        var t=__txt(el);
        if(t.length>60)return;
        var m=t.match(/^kostenstelle\s*:?\s*(.*)$/i);
        if(!m)return;
        out.modus='felder';
        var wert=__norm(m[1]);
        if(!wert&&el.nextElementSibling)wert=__txt(el.nextElementSibling);
        if(wert!==kst)return;
        out.kstGefunden=true;
        var box=el,beste=null;
        for(var k=0;k<8&&box.parentElement;k++){
          box=box.parentElement;
          var anzahl=(__txt(box).match(/kostenstelle/gi)||[]).length;
          if(anzahl>1)break;
          if(phones(__txt(box)).length){beste=box;break;}
        }
        if(beste)res=res.concat(phones(__txt(beste)));
      });
    });
  }
  out.treffer=res.filter(function(x,i){return res.indexOf(x)===i;});
  return JSON.stringify(out);
})
'@

# Seitenaufbau ohne Vertragsdaten (fuer -Diagnose)
$JsDiagnose = $JsGemeinsam + @'
(function(){
  var o={url:location.origin+location.pathname,titel:document.title,iframes:document.querySelectorAll('iframe').length,eingabefelder:[],tabellen:[],kostenstelleFelder:0};
  __docs().forEach(function(d){
    d.querySelectorAll('input,textarea').forEach(function(el){
      if((el.type||'')==='hidden')return;
      o.eingabefelder.push({typ:el.type,placeholder:el.placeholder,aria:el.getAttribute('aria-label'),name:el.name,id:el.id,sichtbar:__vis(el),imKopf:!!(el.closest&&el.closest('header,nav,[role=banner],[role=navigation]'))});
    });
    d.querySelectorAll('table,[role=table],[role=grid],[role=treegrid]').forEach(function(t){
      var kopf=[];
      t.querySelectorAll('th,[role=columnheader]').forEach(function(h){kopf.push(__txt(h).substring(0,40));});
      o.tabellen.push({tag:t.tagName,role:t.getAttribute('role'),sichtbar:__vis(t),ueberschriften:kopf});
    });
    d.querySelectorAll('body *').forEach(function(el){if(el.children.length<=2&&/^kostenstelle\s*:?/i.test(__txt(el))&&__txt(el).length<=60)o.kostenstelleFelder++;});
  });
  return JSON.stringify(o);
})()
'@

function ConvertTo-JsText([string]$s) { return ($s | ConvertTo-Json) }

# Sucht eine Filiale im Portal. Gibt das Ergebnis-Objekt zurueck.
function Search-Filiale([string]$Kst, [string]$ErlaubterPfad) {
    for ($versuch = 1; $versuch -le 2; $versuch++) {
        $feld = Invoke-Js ($JsSuchfeld + '(' + (ConvertTo-JsText $SuchfeldSelektor) + ')')
        if (-not $feld.ok) { if ($versuch -eq 2) { throw 'Suchfeld nicht mehr gefunden oder nicht fokussierbar.' }; Start-Sleep 2; continue }
        if ($feld.pfad -ne $ErlaubterPfad) { throw "Die Seite hat gewechselt ($($feld.pfad)). Abbruch zur Sicherheit." }

        $null = Invoke-Cdp 'Input.insertText' @{ text = $Kst }
        Start-Sleep -Milliseconds 300
        $pruef = Invoke-Js ($JsPruefen + '(' + (ConvertTo-JsText $Kst) + ')')
        if ($pruef.ok) { break }
        if ($versuch -eq 2) { throw "Im Suchfeld steht '$($pruef.wert)' statt '$Kst'. Abbruch zur Sicherheit, es wurde kein Enter gedrueckt." }
        Start-Sleep 1
    }

    Send-Enter

    # Warten, bis das Ergebnis stabil ist
    $start = Get-Date
    Start-Sleep -Milliseconds 1500
    $letztes = $null; $stabil = 0; $r = $null
    while ($true) {
        $r = Invoke-Js ($JsAuslesen + '(' + (ConvertTo-JsText $Kst) + ')')
        $sig = $r | ConvertTo-Json -Compress -Depth 5
        if ($sig -eq $letztes) { $stabil++ } else { $stabil = 0; $letztes = $sig }
        $vergangen = ((Get-Date) - $start).TotalSeconds
        if ($stabil -ge 1 -and (@($r.treffer).Count -gt 0 -or $vergangen -ge 4)) { break }
        if ($vergangen -ge $MaxWarteSekunden) { break }
        Start-Sleep -Milliseconds 700
    }
    # Bei keinem Treffer einmal nachpruefen (Seite evtl. langsam)
    if (@($r.treffer).Count -eq 0) {
        Start-Sleep -Seconds 2
        $r = Invoke-Js ($JsAuslesen + '(' + (ConvertTo-JsText $Kst) + ')')
    }

    $nachher = Invoke-Js ($JsPruefen + '(' + (ConvertTo-JsText $Kst) + ')')
    if ($nachher.pfad -and $nachher.pfad -ne $ErlaubterPfad) { throw "Die Seite hat nach der Suche gewechselt ($($nachher.pfad)). Abbruch zur Sicherheit." }
    return $r
}

# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
function Find-ExcelDatei {
    $direkt = Join-Path $Ordner $DateiName
    if (Test-Path -LiteralPath $direkt -PathType Leaf) { return (Get-Item -LiteralPath $direkt).FullName }
    if (-not (Test-Path -LiteralPath $Ordner)) { throw "Ordner nicht gefunden: $Ordner" }
    $treffer = @(Get-ChildItem -LiteralPath $Ordner -File |
        Where-Object { $_.BaseName -eq $DateiName -and $_.Extension -match '^\.xls[xm]?$' -and $_.Name -notlike '~$*' })
    if ($treffer.Count -eq 0) { throw "Keine Excel '$DateiName' in '$Ordner' gefunden." }
    if ($treffer.Count -gt 1) { throw "Mehrere Dateien '$DateiName.*' gefunden. Bitte -DateiName mit Endung angeben." }
    return $treffer[0].FullName
}

function Test-DateiFrei([string]$Pfad) {
    try { $s = [IO.File]::Open($Pfad, 'Open', 'ReadWrite', 'None'); $s.Close(); return $true }
    catch { return $false }
}

function Get-ZellText($Zelle) {
    $t = [string]$Zelle.Text
    if ($t -match '^#+$') {
        $v = $Zelle.Value2
        if ($v -is [double] -and $v -eq [math]::Floor($v)) { $t = [string][int64]$v } else { $t = [string]$v }
    }
    return $t.Trim()
}

function Find-Kopfzeile($Mappe) {
    foreach ($blatt in $Mappe.Worksheets) {
        $used = $blatt.UsedRange
        $maxR = [Math]::Min(20, $used.Row + $used.Rows.Count - 1)
        $maxC = [Math]::Min(100, $used.Column + $used.Columns.Count - 1)
        for ($r = 1; $r -le $maxR; $r++) {
            $cF = 0; $cR = 0
            for ($c = 1; $c -le $maxC; $c++) {
                $t = ([string]$blatt.Cells.Item($r, $c).Text).Trim()
                if ($cF -eq 0 -and $t -ieq $SpalteFiliale) { $cF = $c }
                if ($cR -eq 0 -and $t -ieq $SpalteRufnummer) { $cR = $c }
            }
            if ($cF -gt 0 -and $cR -gt 0) {
                $letzteSpalte = $blatt.Cells.Item($r, $blatt.Columns.Count).End(-4159).Column   # xlToLeft
                return [pscustomobject]@{ Blatt = $blatt; Zeile = $r; SpalteF = $cF; SpalteR = $cR; LetzteSpalte = [Math]::Max($letzteSpalte, [Math]::Max($cF, $cR)) }
            }
        }
    }
    throw "Keine Kopfzeile mit '$SpalteFiliale' und '$SpalteRufnummer' gefunden."
}

function Set-ZeilenFarbe($Kopf, [int]$Zeile, [int]$Farbe) {
    $b = $Kopf.Blatt
    $b.Range($b.Cells.Item($Zeile, 1), $b.Cells.Item($Zeile, $Kopf.LetzteSpalte)).Interior.Color = $Farbe
}

function Set-Rufnummer($Kopf, [int]$Zeile, [string]$Nummer) {
    $zelle = $Kopf.Blatt.Cells.Item($Zeile, $Kopf.SpalteR)
    $zelle.NumberFormat = '@'     # als Text, damit die fuehrende 0 bleibt
    $zelle.Value2 = $Nummer
}

# ---------------------------------------------------------------------------
# Ablauf
# ---------------------------------------------------------------------------
$logOrdner = Join-Path $script:Basis 'Protokolle'
New-Item -ItemType Directory -Path $logOrdner -Force | Out-Null
$script:LogDatei = Join-Path $logOrdner "Lauf_$($script:Zeitstempel).txt"

Write-Host ''
Write-Host '=== Telekom Rufnummern-Abgleich ===' -ForegroundColor Cyan
if ($Testlauf) { Write-Host "TESTLAUF: nur $TestAnzahl Filialen, die Excel wird NICHT veraendert." -ForegroundColor Yellow }
Write-Host ''

$excel = $null; $mappe = $null
$aenderungen = 0
try {
    # --- 1. Chrome vorbereiten -------------------------------------------------
    Start-SteuerbaresChrome
    Write-Host ''
    Write-Host 'Im Chrome-Fenster, das das Skript geoeffnet hat:' -ForegroundColor Cyan
    Write-Host '  1. Im Telekom-Portal einloggen'
    Write-Host '  2. "Meine Vertraege" oeffnen'
    Write-Host '  3. Waehrend das Skript laeuft, das Fenster nicht benutzen'
    Write-Host '(Dein normales Chrome zaehlt nicht, das kann das Skript nicht steuern.)' -ForegroundColor DarkGray
    $null = Read-Host 'Wenn "Meine Vertraege" offen ist: Enter druecken'

    $tab = Select-PortalTab
    Connect-Cdp $tab.webSocketDebuggerUrl
    $null = Invoke-Cdp 'Page.bringToFront'
    $null = Invoke-Cdp 'Emulation.setFocusEmulationEnabled' @{ enabled = $true }
    Write-Log "Verbunden mit Tab: $($tab.title)"

    if ($Diagnose) {
        $d = Invoke-Js $JsDiagnose
        $datei = Join-Path $script:Basis "Diagnose_$($script:Zeitstempel).txt"
        $d | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $datei -Encoding UTF8
        Write-Log "Diagnose gespeichert: $datei" Green
        Write-Log 'Die Datei enthaelt nur den Seitenaufbau (Felder, Ueberschriften), keine Vertragsdaten.'
        return
    }

    # --- 2. Suchfeld pruefen ---------------------------------------------------
    $feld = Invoke-Js ($JsSuchfeld + '(' + (ConvertTo-JsText $SuchfeldSelektor) + ')')
    if (-not $feld.ok) {
        throw 'Kein Suchfeld auf der Seite gefunden. Bitte mit -Diagnose starten und die Datei schicken.'
    }
    $erlaubterPfad = $feld.pfad
    Write-Host ''
    Write-Host "Gefundenes Suchfeld: $($feld.info)" -ForegroundColor Cyan
    Write-Host "Seite:              $erlaubterPfad" -ForegroundColor Cyan
    Write-Host 'Das Feld sollte jetzt in Chrome den Cursor haben.'
    if (-not (Confirm-JaNein 'Ist das das Suchfeld unter "Meine Vertraege"?')) {
        throw 'Abgebrochen. Tipp: mit -Diagnose starten oder -SuchfeldSelektor angeben.'
    }

    # --- 3. Excel oeffnen ------------------------------------------------------
    $pfad = Find-ExcelDatei
    Write-Log "Excel: $pfad"
    if (-not $Testlauf) {
        if (-not (Test-DateiFrei $pfad)) { throw 'Die Excel ist geoeffnet oder gesperrt. Bitte schliessen und neu starten.' }
        $backupOrdner = Join-Path (Split-Path $pfad) 'Backups'
        New-Item -ItemType Directory -Path $backupOrdner -Force | Out-Null
        $info = Get-Item -LiteralPath $pfad
        $backup = Join-Path $backupOrdner ("{0}_BACKUP_{1}{2}" -f $info.BaseName, $script:Zeitstempel, $info.Extension)
        Copy-Item -LiteralPath $pfad -Destination $backup
        if (-not (Test-Path -LiteralPath $backup)) { throw 'Sicherheitskopie konnte nicht erstellt werden. Abbruch.' }
        Write-Log "Sicherheitskopie: $backup" Green
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3   # Makros nie ausfuehren
    $mappe = $excel.Workbooks.Open($pfad, 0, [bool]$Testlauf)
    $kopf = Find-Kopfzeile $mappe
    Write-Log "Blatt '$($kopf.Blatt.Name)', Kopfzeile $($kopf.Zeile)"

    $b = $kopf.Blatt
    $letzteZeile = $b.Cells.Item($b.Rows.Count, $kopf.SpalteF).End(-4162).Row   # xlUp
    $zuSuchen = New-Object System.Collections.ArrayList
    $schonVorhanden = 0
    for ($z = $kopf.Zeile + 1; $z -le $letzteZeile; $z++) {
        $fil = Get-ZellText $b.Cells.Item($z, $kopf.SpalteF)
        if (-not $fil) { continue }
        $vorhanden = Get-ZellText $b.Cells.Item($z, $kopf.SpalteR)
        if ($vorhanden) { $schonVorhanden++; continue }
        [void]$zuSuchen.Add([pscustomobject]@{ Zeile = $z; Filiale = $fil })
    }
    Write-Log "Filialen ohne Rufnummer: $($zuSuchen.Count)  |  schon mit Rufnummer (bleiben unberuehrt): $schonVorhanden"
    if ($zuSuchen.Count -eq 0) { Write-Log 'Nichts zu tun.' Green; return }
    if ($Testlauf) { $zuSuchen = @($zuSuchen | Select-Object -First $TestAnzahl) }
    elseif (-not (Confirm-JaNein "$($zuSuchen.Count) Filialen jetzt suchen und Excel bearbeiten?")) { throw 'Abgebrochen.' }

    # --- 4. Suchen -------------------------------------------------------------
    $strukturErkannt = $false
    $offenOhneStruktur = New-Object System.Collections.ArrayList
    $zaehler = @{ gruen = 0; rot = 0; gelb = 0 }
    $nr = 0
    foreach ($eintrag in $zuSuchen) {
        $nr++
        $ergebnis = Search-Filiale $eintrag.Filiale $erlaubterPfad
        $nummern = @($ergebnis.treffer)
        $text = "[$nr/$($zuSuchen.Count)] Filiale $($eintrag.Filiale) (Zeile $($eintrag.Zeile)): "
        if ($ergebnis.modus -ne 'none') { $strukturErkannt = $true }

        if ($nummern.Count -eq 1) {
            Write-Log ($text + "Treffer $($nummern[0]) -> gruen") Green
            if (-not $Testlauf) { Set-Rufnummer $kopf $eintrag.Zeile $nummern[0]; Set-ZeilenFarbe $kopf $eintrag.Zeile $FarbeGruen; $aenderungen++ }
            $zaehler.gruen++
        }
        elseif ($nummern.Count -gt 1) {
            Write-Log ($text + "MEHRERE Nummern: $($nummern -join '; ') -> gelb") Yellow
            if (-not $Testlauf) { Set-Rufnummer $kopf $eintrag.Zeile ($nummern -join '; '); Set-ZeilenFarbe $kopf $eintrag.Zeile $FarbeGelb; $aenderungen++ }
            $zaehler.gelb++
        }
        elseif ($ergebnis.kstGefunden) {
            Write-Log ($text + 'Kostenstelle gefunden, aber keine Rufnummer erkannt -> gelb') Yellow
            if (-not $Testlauf) { Set-ZeilenFarbe $kopf $eintrag.Zeile $FarbeGelb; $aenderungen++ }
            $zaehler.gelb++
        }
        elseif ($strukturErkannt) {
            Write-Log ($text + 'kein Treffer -> rot') Red
            if (-not $Testlauf) { Set-ZeilenFarbe $kopf $eintrag.Zeile $FarbeRot; $aenderungen++ }
            $zaehler.rot++
        }
        else {
            # Tabelle noch nie erkannt: lieber noch nicht rot markieren
            Write-Log ($text + 'kein Treffer (Tabelle noch nicht erkannt, wird spaeter entschieden)') DarkYellow
            [void]$offenOhneStruktur.Add($eintrag)
        }

        # Zeilen, die vor dem Erkennen der Tabelle offen blieben, jetzt rot markieren
        if ($strukturErkannt -and $offenOhneStruktur.Count -gt 0) {
            foreach ($o in $offenOhneStruktur) {
                Write-Log "Filiale $($o.Filiale) (Zeile $($o.Zeile)): kein Treffer -> rot" Red
                if (-not $Testlauf) { Set-ZeilenFarbe $kopf $o.Zeile $FarbeRot; $aenderungen++ }
                $zaehler.rot++
            }
            $offenOhneStruktur.Clear()
        }

        if (-not $Testlauf -and $aenderungen -gt 0 -and ($nr % 10) -eq 0) { $mappe.Save(); Write-Log 'Zwischenstand gespeichert.' DarkGray }
        Start-Sleep -Milliseconds 800
    }

    if ($offenOhneStruktur.Count -gt 0) {
        Write-Log "$($offenOhneStruktur.Count) Filialen nicht markiert: Die Ergebnis-Tabelle wurde nie erkannt. Bitte mit -Diagnose pruefen." Yellow
    }
    Write-Host ''
    Write-Log "Fertig. Gruen: $($zaehler.gruen)  Rot: $($zaehler.rot)  Gelb: $($zaehler.gelb)" Cyan
    if ($Testlauf) { Write-Log 'Testlauf: Excel wurde nicht veraendert.' Yellow }
}
catch {
    Write-Log "FEHLER: $($_.Exception.Message)" Red
    if (-not $Testlauf -and $aenderungen -gt 0) {
        Write-Log 'Bereits gemachte Eintraege werden gespeichert. Die Sicherheitskopie liegt im Ordner "Backups".' Yellow
    }
}
finally {
    if ($mappe) {
        if (-not $Testlauf -and $aenderungen -gt 0) {
            try { $mappe.Save(); Write-Log "Excel gespeichert ($aenderungen Aenderungen)." Green }
            catch { Write-Log "Speichern fehlgeschlagen: $($_.Exception.Message)" Red }
        }
        $mappe.Close($false)
    }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if ($script:Ws) { try { $script:Ws.Dispose() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($script:LogDatei) { Write-Host "Protokoll: $($script:LogDatei)" -ForegroundColor DarkGray }
}
