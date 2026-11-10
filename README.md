# Magic-Sterne

Kasse und Bestandsverwaltung für den Verkauf von Papiersternen auf Weihnachtsmärkten.
Es ist eine Bestandsverwaltung mit Preisrechner, kein Kassensystem: Es werden keine Belege
erstellt und keine Beträge pro Verkauf gespeichert, nur Stückzahlen.

- `docs/`: die Web-App (PWA), ausgeliefert über GitHub Pages.
  Auf Tablet oder Handy öffnen und „Zum Startbildschirm hinzufügen“ wählen.
- `gas/`: Google Apps Script am Google Sheet (Datenhaltung und Schnittstelle).
  Hochladen mit `clasp push`, danach `clasp deploy -i <Deployment-ID>`.

Daten, PINs und die Adresse der Schnittstelle stehen **nicht** in diesem Repository.
Sie liegen im Google Sheet bzw. in den Script-Eigenschaften und werden pro Gerät über
einen Einrichtungslink (`…/#setup=<Adresse der Schnittstelle>`) und die PIN hinterlegt.

Neue App-Version: in `docs/sw.js` die `VERSION` erhöhen, sonst behalten installierte
Geräte die alte Fassung.
