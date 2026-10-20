# Telekom Rufnummern-Abgleich

Liest die Filialen aus `Filial_Rufnummer.xlsx`, sucht sie im Telekom-Portal unter **Meine Verträge**
und trägt die Mobilrufnummer ein.

| Ergebnis | Excel |
|---|---|
| Kostenstelle = Filiale, 1 Nummer | Nummer in „Mobilrufnummer“, Zeile **grün** |
| Kein Treffer | Zeile **rot** |
| Mehrere Nummern / Nummer nicht lesbar | Zeile **gelb** (bitte von Hand prüfen) |
| Rufnummer schon vorhanden | **nichts** passiert |

## Sicherheit
- Vor jedem echten Lauf: Sicherheitskopie im Unterordner `Backups` neben der Excel.
- Auf der Website wird **nichts angeklickt**. Das Skript tippt nur die Filialnummer ins Suchfeld und drückt Enter.
  Vor jedem Enter prüft es, dass der Cursor im Suchfeld steht und genau die Filialnummer drinsteht.
- Wechselt die Seite (z. B. auf eine andere Unterseite), bricht das Skript sofort ab.
- Kein Passwort im Skript, du loggst dich selbst ein.
- Makros in der Excel werden nicht ausgeführt.
- Jede Änderung steht im Protokoll (`Protokolle\Lauf_...txt`).

## Benutzung
1. Excel schließen.
2. PowerShell im Ordner des Skripts öffnen und **zuerst den Testlauf** starten:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Telekom-Rufnummern-Abgleich.ps1 -Testlauf
   ```
3. Es öffnet sich ein **eigenes Chrome-Fenster** (getrennt von deinem normalen Chrome).
   Dort einloggen, „Meine Verträge“ öffnen, im PowerShell-Fenster Enter drücken.
4. Das Skript zeigt das gefundene Suchfeld und fragt nach. Danach sucht es 3 Filialen und zeigt das Ergebnis
   **ohne** die Excel zu ändern.
5. Stimmt alles: ohne `-Testlauf` starten.
6. Während des Laufs das Chrome-Fenster nicht benutzen.

## Wenn etwas nicht erkannt wird
```powershell
powershell -ExecutionPolicy Bypass -File .\Telekom-Rufnummern-Abgleich.ps1 -Diagnose
```
Erstellt `Diagnose_...txt` mit dem Aufbau der Seite (Suchfelder, Tabellen-Überschriften, **keine Vertragsdaten**).
Diese Datei schicken, dann wird das Skript angepasst.
