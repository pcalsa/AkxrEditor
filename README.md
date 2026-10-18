# AKXR Editor

Kostenloser Editor für Text, Tabellen, Zeichnungen, PDF-Ausdrucke, Bilder und Dateianhänge. Verarbeitung und Speicherung laufen im Browser. Kein Konto, Serverdienst, kostenpflichtiger API-Schlüssel oder Abonnement erforderlich.

## Starten und prüfen

```powershell
npm ci
npm run dev
```

```powershell
npm run lint
npm audit
npx playwright install chromium
npm test -- --workers=1
```

`npm test` baut die App und startet Browserprüfungen gegen Entwicklungsserver und Produktionsvorschau. `npm run test:browser` führt nur die Browserprüfungen aus. Für Offline-Nutzung ist der Produktionsbuild erforderlich:

```powershell
npm run build
npm run preview
```

Die Seite einmal online unter derselben Adresse öffnen und auf „Offline bereit“ warten. Danach sind Neuladen und lokale Import-/Exportfunktionen auch ohne Netzwerk verfügbar. Der Service Worker speichert auch die nachgeladenen Bibliotheken und den PDF-Worker. Auf einem anderen Gerät muss die App zunächst ebenfalls geladen werden. Hosting funktioniert über HTTPS beziehungsweise lokal über localhost.

## Bedienung und lokale Daten

- Unter **Seite finden und bearbeiten** führen vier farbige Bereiche durch Suche, Umbenennen/Einordnen, Dateianhänge und Wiederherstellung. Das Feld „Bereich“ verschiebt die aktive Seite in einen anderen oder neuen Bereich. Seitentabs lassen sich ziehen; die Tab-Leiste scrollt horizontal. Erklärtexte erscheinen nach etwa 650 ms Verweilen mit Maus oder Tastaturfokus. Statusmeldungen bestätigen Kopieren, Anhänge und Wiederherstellungen; Projektstände zeigen Zeitpunkt und Seitenzahl. Vor einer Versionswiederherstellung wird der aktuelle Stand gesichert.
- Die Suche berücksichtigt alle Seiten, Schlagwörter und frei platzierte Canvas-Texte. Treffer öffnen die betreffende Seite.
- Die Arbeitsfläche zeigt die aktive Seite. Damit gibt es unabhängig von der Seitenzahl nur einen Fabric-Canvas mit seiner Interaktionsfläche. Lange Texte vergrößern die Notizfläche; sie werden nicht unten abgeschnitten. Sie werden nicht automatisch auf neue A4-Seiten verteilt.
- **Rückgängig/Wiederholen** verwendet einen gemeinsamen Verlauf für Text, Zeichnungen, Bilder und Seiten. Zusammenhängende Texteingaben werden gruppiert. Bis zu 80 Aktionen bleiben während der Sitzung erhalten, auch beim Wechsel zum Konverter. Bei einer Aktion auf einer anderen Seite wird diese Seite geöffnet.
- Autosave verwendet IndexedDB. Die Anzeige bestätigt erst einen abgeschlossenen Schreibvorgang. Speicherfehler und Konflikte mit anderen Tabs werden angezeigt; ein fehlgeschlagener Schreibvorgang gilt weiter als ungespeichert. Projekt-Downloads bleiben möglich. Beim Verbergen oder Verlassen wird zusätzlich gespeichert; vor noch ungespeicherten Änderungen warnt der Browser.
- Mehrere Tabs überschreiben den Entwurf nicht ungeprüft. Eine atomare Revisionsprüfung stoppt konkurrierende Schreibvorgänge. Im Konfliktfall das eigene Projekt herunterladen oder ausdrücklich den Stand des anderen Tabs laden.
- Bis zu 50 lokale Versionsstände bleiben über Neuladen und Browser-Neustarts erhalten. Bei Bearbeitung wird ungefähr alle 30 Sekunden ein Stand angelegt, außerdem vor Projektwechseln und beim Abrufen des Versionsverlaufs. Der Papierkorb gehört zum Projekt und wird auch in `.akxr` gespeichert.
- Lokale Daten gehören zur Browseradresse und zum Browserprofil. Das Löschen der Website-Daten entfernt auch Entwurf und Versionen. `.akxr` ist das vollständige portable Projektformat; es enthält auch Anhänge.
- Der **Objektradierer** entfernt berührte Zeichenobjekte einschließlich ihrer gespeicherten Daten. Er übermalt nicht weiß. Einzelne Pixel eines Bildes beziehungsweise Teilstücke eines vorhandenen Pfades werden damit nicht herausgeschnitten.
- Der zusätzliche **Flächenradierer** radiert nur die überstrichenen Bereiche: aus Bildern und PDF-Ausdrucken werden die Pixel mit Transparenz entfernt, normale Schrift erhält gespeicherte Masken pro betroffenem Zeichen. Betroffene Canvas-Objekte werden in Pixelobjekte umgewandelt; unberührte Zeichenobjekte bleiben bearbeitbar. Der Durchmesser ist separat einstellbar. Eine Radierbewegung ist ein Rückgängig-Schritt; Rückgängig stellt auch die ursprünglichen Vektorobjekte wieder her. HTML-, PDF-, Bild- und DOCX-Export berücksichtigen die Ausschnitte. TXT und Markdown lassen angeschnittene Schriftzeichen weg, da diese Formate keine teilweise sichtbaren Buchstaben darstellen können.
- Zeichnungsobjekte werden anhand ihrer sichtbaren Pixel getroffen, ohne zusätzliche Innenabstände um den Auswahlrahmen. Beim Objektradieren gibt es weder Verschieben noch Aufziehen eines Auswahlkastens.
- Die kleinen **− / Prozent / +**-Schaltflächen ändern ausschließlich die Seitenansicht zwischen 35 % und 200 %. Der Prozentknopf setzt auf 100 % zurück. Werkzeugleisten, Dokumentmaße, Schriftgrößen und Exporte bleiben davon unabhängig; Pointer- und Bildkoordinaten werden an die Vergrößerung angepasst.
- **Stiftdruck** speichert einen variabel breiten Vektorpfad. **Nur Stift / Handballenerkennung** ignoriert Touch-Ereignisse auf der Zeichenfläche. Die Qualität der Druckwerte hängt vom Stift, Treiber und Browser ab; dies ist eine zuschaltbare Eingabefilterung, keine geräteunabhängige automatische Handerkennung.
- Englisch, Deutsch, Arabisch und Dänisch haben vollständige, explizite Sprachkataloge in `src/i18n.ts`. Oberfläche, Hinweise, Statusmeldungen und Datumsangaben wechseln direkt mit der Sprache; die Projektsprache wird gespeichert. Arabisch verwendet RTL. Fehlende Übersetzungsschlüssel werden beim Bauen und in den Sprachtests erkannt.

## Import und Export

Projektdateien werden vollständig geprüft, bevor sie das offene Dokument ersetzen. HTML wird mit DOMPurify bereinigt; ungültige Seiten oder Zeichnungen werden abgewiesen und nicht still ausgefiltert. Abgebrochene oder inzwischen überholte Importe dürfen neuere Arbeit nicht ersetzen.

Neue Tabellen verwenden den echten Texteditor mit mehreren Zeilen, Tab-Navigation und Zeilen-/Spaltenbefehlen. Alte Canvas-Tabellen können weiter geöffnet und bearbeitet werden.

| Format | Verhalten |
| --- | --- |
| AKXR | Vollständiges Projekt einschließlich Seitenformaten, Darstellung, Canvas, Metadaten, Papierkorb und Anhängen |
| TXT / Markdown | Seitentext und Canvas-Texte; TXT erhält Absatz-/Zeilenumbrüche. Markdown-Import verwendet einen GFM-Parser für Links, Bilder, Tabellen, Code und verschachtelte Listen. |
| HTML | Bereinigte Formatierung und Bilder; Sprache, Schreibrichtung, Zeilenabstand und Ebenenreihenfolge werden beim Export mitgegeben. |
| DOCX | Bearbeitbare Absätze, Formatierungen, Links und normale Tabellen; Bilder und Zeichnungen ergänzen diese Inhalte als Bilder. Frei positionierte Canvas-Inhalte werden in Word im Dokumentfluss ergänzt, nicht als identische frei positionierte Word-Objekte rekonstruiert. |
| PDF | Physische Maße aus 96-dpi-Pixelmaßen in Millimeter; PDF-Import behält Originalmaße und Orientierung bei. Export bleibt bildbasiert und enthält keine durchsuchbare Textebene. |
| PNG / JPG | Seiten werden nacheinander gerendert. Mehrseitige Konverter-Ergebnisse werden als ZIP geliefert. Der Editor bestätigt vor dem Download mehrerer einzelner Seitenbilder. |

Der Konverter kopiert gleiche Quell-/Zielformate unverändert. PDF → DOCX erhält die Seitenansicht als Bild und ergänzt extrahierbaren Text bearbeitbar. Es gibt keine OCR für gescannte Dokumente. Bild → TXT/Markdown wird deshalb nicht angeboten. DOCX-Import berücksichtigt getrennte Abschnitte und Nullränder; komplexe Word-Elemente wie schwebende Rahmen, Kopf-/Fußzeilen und spezielle Felder werden vom verwendeten HTML-Konverter nicht verlustfrei rekonstruiert. HTML-/Markdown-Rundreisen sind ebenfalls kein Ersatz für `.akxr`.

Zum Schutz des Browsers gelten Ressourcenlimits: 50 MB pro Dokumentimport, 10 MB pro Anhang, 150 PDF-Seiten, insgesamt 100 Millionen PDF-Renderpixel, 16 Millionen Pixel für den Fülleimer und 50 Millionen Pixel pro exportiertem Seitenbild. Größere Vorgänge werden mit Fehler abgebrochen. PDF-Seiten werden nacheinander gerendert, Ressourcen anschließend freigegeben; der Fülleimer arbeitet in einem Worker.

## Prüfstand

Die Browserprüfungen decken eingeschleustes JavaScript, beschädigte Projektdateien, seitenbezogenen Verlauf, Autosave einschließlich Quota-Fehler und Tab-Konflikten, Auswahlformatierung, editierbares DOCX, physische PDF-Maße und Importformate, HTML-Bilder/Formatierung, Tabellen, Papierkorb, lange Notizen, Touch/Stiftdruck, schmale Fenster und Offline-Neuladen mit Export ab. Synthetische Pointer-Tests ersetzen keinen Test mit einem realen Stiftgerät.

Die Sicherheitsaktualisierung verwendet unter anderem ProseMirror View 1.42.6 und Tiptap 3.32.0. Das [ProseMirror-Advisory](https://github.com/ProseMirror/prosemirror-view/security/advisories/GHSA-c8x8-7fp4-3x9w) beschreibt die behobene Einfüge-Lücke. Die Sperrdatei enthält die geprüften Versionen; `npm audit` sollte bei späteren Updates erneut ausgeführt werden.

Import-/Exportbibliotheken werden bei Bedarf nachgeladen. Das Hauptpaket liegt mit dem zusätzlichen Radierer und vier Sprachen bei etwa 1,20 MB (gzip etwa 377 kB), gegenüber dem ursprünglich gemeldeten Stand von etwa 2,71 MB. Vite meldet dafür weiterhin eine Größenwarnung. Ein vollständiger Test sämtlicher Kombinationen komplexer Fremddokumente oder eine unabhängige Sicherheitsprüfung ist damit nicht ersetzt.
