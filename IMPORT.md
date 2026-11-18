# Import neuer Bestandslisten (ab Saison 2026)

Gilt für die Listen im Format von 2026: „Magic Sterne Bestand.xlsx“ und „Mea_Sterne_Abstrichliste.xlsx“.
Das alte Format („Sterne 2025.xlsx“ mit den Blättern „Kalkulation“ und „Sterne Bestellung“) wird nicht mehr gelesen.
Den Code dafür gibt es noch in der Git-Historie und in der Sicherung unter `sternkasse-backups/`.

## Was der Import erkennt

- **Kopfzeile:** Es ist die Zeile mit einer Spalte „Nr.“. Steht in der Zeile darunter in der Nr.-Spalte nichts, gilt sie als zweite Kopfzeile, zum Beispiel „Lager / Fulda“.
- **Spalten:** Gesucht werden „Bezeichnung“ (fehlt sie, gilt die Spalte rechts von „Nr.“), „VK“ oder „Preis“, „Lager“ und je Marktstand eine Spalte mit dessen Namen aus dem Blatt „Orte“, zum Beispiel „Fulda“.
  Alles ab der zweiten „Nr.“-Spalte (der Strichliste) wird ignoriert.
- **Kategorie:** Sie kommt aus der Kopfzeile, sonst aus dem Blatt- oder Dateinamen. Enthält der Text „Blumen“, wird es **Blumen-Sterne**, enthält er „Magic“, wird es **Magic-Sterne**.
  Die Kategorie landet in der Spalte „Filter: Kategorie“ und erscheint in der Kasse als Filter.
- **Zubehör:** Alle Nummern, die mit „Z“ beginnen. Leerzeichen in Nummern werden entfernt, aus „ZB  01“ wird also „ZB01“.
- **Hinweise:** Zusätze wie „Keine Nachlieferung“, „Restbestand“ oder „Rest“ am Ende des Namens wandern in die Spalte „Hinweis“.
- **Bestände:** Sie werden je Ort als **Inventur** gebucht. Gebucht wird also die Differenz zum aktuellen Bestand, ein erneuter Import doppelt nichts.

## Doppelte Nummern

- Für den Abgleich zählen führende Nullen nicht: „9“, „09“ und „009“ gelten als dieselbe Nummer.
- **Gleiche Nummer bei Magic- und Blumen-Sternen:** Beide bekommen ein Präfix, also `M9` und `B009`. Das Importprotokoll nennt jeden Fall.
- **Gleiche Nummer zweimal in derselben Datei:** Das ist ein Fehler, der Import bricht ab. Bitte die Liste korrigieren.
- **Gleiches Zubehör in beiden Dateien:** Es wird einmal angelegt. Steht dieselbe Nummer für zwei verschiedene Artikel, ist das ein Fehler.
- **Vorhandene Artikel** werden über die Nummer gefunden. Preis, Name und Kategorie werden aktualisiert.
  Zubehör-Zuordnung, Mindestbestände, Bild und Kurzname bleiben erhalten.
- **Artikel, die in einer neuen Liste fehlen**, werden ausgeblendet (Aktiv = nein), aber nicht gelöscht. Das gilt nur für die Kategorien, die importiert wurden.

## Ablauf im Google Sheet

1. **Excel nach Google Drive hochladen.** Jede Datei öffnen und **Datei → Als Google Tabellen speichern** wählen. Die Links der neuen Google-Tabellen kopieren.
2. Das Magic-Sterne-Sheet öffnen und die Seite neu laden, damit das Menü aktuell ist.
3. **Menü Magic-Sterne → Sicherung erstellen.** Das legt eine vollständige Kopie mit Datum im selben Drive-Ordner an.
4. Nur wenn alles neu beginnen soll: **Menü → Testdaten entfernen …**
   Das löscht alle Artikel und Bewegungen. Orte, PINs und Bilderordner bleiben.
   Der Menüpunkt geht nur, wenn die letzte Sicherung jünger als eine Stunde ist.
5. **Menü → Import aus Bestandslisten …** öffnen, beide Links mit Leerzeichen getrennt einfügen und OK wählen.
   - Zuerst kommt nur eine **Prüfung**. Sie zeigt die Zusammenfassung und Hinweise, alles steht außerdem im Blatt „Importprotokoll“.
   - Bei Fehlern wird nichts geschrieben.
   - Mit „Ja“ wird importiert.
6. Im Blatt **Artikel** nachsehen:
   - **Zubehör je Stern:** Vorbelegt in fester Reihenfolge (Ständer nur bei 60 cm):
     - Magic-Sterne: LED E14 4 W → Kabel E14 → Ständer → LED E14 8 W
     - Blumen-Sterne: LED E14 8 W → Kabel E14 → LED E14 4 W → Ständer → LED E27 → Kabel E27

     Die Regel steht in `ZUB_REIHENFOLGE_` in `gas/Import.gs`. Mit **Menü → Zubehör-Vorschläge neu setzen …** wird sie auf alle Sterne angewendet; von Hand geänderte Einträge gehen dabei verloren.
   - **Mindestbestände:** Die Spalten „Min: …“ sind vorbelegt.
   - **Sterne ohne Größe:** Sie haben keinen Größenfilter. Das Importprotokoll nennt sie.
7. Auf den Geräten die App neu laden oder in den Einstellungen „Abgleich“ wählen.

## Vorher lokal prüfen (optional)

Ohne Google lässt sich sehen, was der Import machen würde (Node.js nötig):

```
cd tools
npm install
node test-import.js "C:\Pfad\Magic Sterne Bestand.xlsx" "C:\Pfad\Mea_Sterne_Abstrichliste.xlsx"
```
