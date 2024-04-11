# webuntis2BlaueBriefe

## Webuntis2BlaueBriefe generiert versandfertige Blaue Briefe als einzelne Word- und PDF-Dokumente.

Technische Voraussetzung sind: 

1. der Einsatz von Atlantis als Schulverwaltungsprogramm
2. der Einsatz von Webuntis und Untis auf Basis einer SQL-Datenbank
3. Installation von Word 

## Organisatorische Voraussetzungen sind:
1. eine Prüfungsart namens  "Blauer Brief" mit dem Langnamen *Mahnung gem. §50 (4) SchulG (Blauer Brief)* ist in Webuntis angelegt
2. die Lehrerinnen und Lehrer haben die Blauen Briefe in der Prüfungseintrag "Blauer Brief" eingetragen 
3. die Datei ```MarksPerLesson.csv``` wurde aus Webuntis exportiert und im Download-Ordner abgelegt.

### Detailierte Schritte der Erstellung der Datei namens MarksPerLesson.csv:

#### Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie:

1. Klassenbuch > Berichte klicken
2. Alle Klassen auswählen
3. Unter "Noten" die Prüfungsart "Alle" auswählen
4. Hinter "Noten pro Schüler" auf CSV klicken
5. Die Datei "MarksPerLesson.csv" im Download-Ordner speichern.