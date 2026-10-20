<#
    TEST mit genau 1 Filiale.

    Sucht nur die erste Filiale ohne Rufnummer und traegt das Ergebnis in die Excel ein
    (gruen / rot / gelb). Vorher wird eine Sicherheitskopie im Ordner "Backups" angelegt.

    Starten:
        powershell -ExecutionPolicy Bypass -File .\Test-1-Filiale.ps1
#>
& (Join-Path $PSScriptRoot 'Telekom-Rufnummern-Abgleich.ps1') -MaxAnzahl 1 @args
