# Quellen und Änderungen

Stand: 29.10.2026.

## Übernommener Ausgangspunkt

- Ursprungsrepository: https://github.com/franziskasuwita-glitch/SHK-Personality-Project
- Fork: https://github.com/LFFLFFLFF/SHK-Personality-Project
- Ausgangscommit: `58f68d06f2ba2158e80ffdc90f155fd57d9eda5c`
- Referenzdatei: `Itemselection.py`
- Referenzdatei als Git-Blob: `b4620749ab0c55f1743f0d9366934975368c1ee1`
- Referenzdatei am festgehaltenen Stand: https://github.com/franziskasuwita-glitch/SHK-Personality-Project/blob/58f68d06f2ba2158e80ffdc90f155fd57d9eda5c/Itemselection.py

Die vorhandene MIT-Lizenz und der ursprüngliche Copyright-Hinweis werden beibehalten. Die zusätzliche Datei `01_embeddings.py` greift das Vorgehen der Embeddingbildung und Itemmittelung auf. Einlesen, Prüfungen, Zuordnung, getrennte Kontrollen, Laufprotokoll und die sichere Trennung von Prüfung und Berechnung wurden neu für diesen Datensatz ausgearbeitet.

## Methodische Referenzen

Hütter beschreibt auf S. 28 bis 29 das Modell `paraphrase-multilingual-MiniLM-L12-v2`, die Einzel-L2-Normierung, die Mittelung von Original und Varianten sowie eine erneute L2-Normierung des Itemmittels. Suwita beschreibt auf den gedruckten S. 11 bis 14 dasselbe Modell und die Verarbeitung von Originalen, Varianten und Definitionsvektoren. Suwitas Anhang B verweist auf das Ursprungsrepository.

Es wird keine exakte technische Replikation der damaligen Läufe behauptet. Die Arbeiten legen insbesondere keine hier unabhängig belegte identische Modellrevision oder vollständige damalige Laufzeitumgebung fest. Die festgehaltene neue Modellrevision ergibt sich aus den öffentlichen [Modellmetadaten](https://huggingface.co/api/models/sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2).

## Änderungen für diese Masterarbeit

1. Einlesen der benannten Excel-Tabelle statt einer fest eingebauten Colab-TXT-Datei.
2. Arbeitsfassung aus C, falls vorhanden, sonst Original aus B. Varianten G bis I. Eindeutige Itemcodes als getrennte Metadaten.
3. Erwartete Auswahl: 53 Items, 34 Rumination und 19 Hostility, 158 Varianten, insgesamt 211 reguläre Texte.
4. Kein reguläres Item wird wegen seiner Position als Kontrollitem behandelt. Kein Rückgriff auf Beispielitems bei fehlender Eingabedatei.
5. Die fünf bereits textuell umgepolten Analysefassungen gehen in hoher Merkmalsrichtung ein. Ursprüngliche niedrige Texte werden in diesen Gruppen nicht mitgemittelt. Keine zusätzliche rechnerische Invertierung.
6. Separat gekennzeichnete Kontrolltexte und Grenzfälle. Zunächst ein selbst formulierter neutraler Kontrolltext KONT01. RRQ3 bleibt aus der Hauptauswahl ausgeschlossen.
7. Standardaufruf ausschließlich zur Prüfung. Modellbibliotheken werden erst im ausdrücklich gestarteten Berechnungszweig importiert.
8. Aufbewahrung ursprünglicher und normierter Einzelvektoren sowie mehrerer eindeutig benannter Itemmittel. Keine Zentrierung oder Projektion in Schritt 1.
9. Prüfungen von Eingabe, Anzahlen, Polung, Rollen und Datenzuordnung. Im späteren Berechnungszweig zusätzlich Prüfung von Tokenlängen, Dimensionen, endlichen Werten und Nullvektoren.
10. Festhalten von Modellrevision, Parametern, Paketversionen und Prüfsummen für den jeweiligen Lauf.

## Was noch nicht empirisch geprüft wurde

Die Textumformulierungen und der neutrale Kontrolltext sind methodisch dokumentierte Arbeitsmaterialien. Die Datenprüfung weist nach, dass der vorgesehene Textbestand korrekt ausgewählt und zugeordnet wird. Sie belegt weder psychometrische Gleichwertigkeit der umformulierten Items noch Validität des Podcastverfahrens.

Der echte Embedding-Berechnungszweig wurde noch nicht mit einem geladenen Modell ausgeführt. Es wurden keine Embeddings erzeugt. SQuID-Zentrierung, Auswahl der Konstruktachsen und Achsenprojektion bleiben gesonderte folgende Arbeitsschritte.
