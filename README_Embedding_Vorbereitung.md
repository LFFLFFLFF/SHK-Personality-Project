# Embedding-Vorbereitung für Rumination und Hostility

Stand: 29.10.2026.

Diese Ergänzung des Forks bereitet die Einbettung der 53 ausgewählten Items vor. Ausgangspunkt ist `Itemselection.py` aus dem [SHK-Personality-Project von Franziska Suwita](https://github.com/franziskasuwita-glitch/SHK-Personality-Project). Quelle und Änderungen sind in [QUELLEN_UND_AENDERUNGEN.md](QUELLEN_UND_AENDERUNGEN.md) dokumentiert.

Der Standardaufruf prüft ausschließlich die Eingabe. Er lädt kein Embedding-Modell und berechnet keine Vektoren. Der Berechnungszweig ist für einen späteren, ausdrücklich gestarteten Lauf vorbereitet.

## Dateien

| Datei | Aufgabe |
| --- | --- |
| `01_embeddings.py` | Einlesen, Datenprüfung und ausdrücklich startbarer späterer Embedding-Lauf |
| `config_embeddings.json` | Eingabepfad, erwartete Anzahlen, festgehaltene Modellversion und Einstellungen |
| `requirements-embeddings.txt` | Benötigte Bibliotheken für diesen Schritt |
| `data/Itemdaten_mit_Kontrolltext.xlsx` | Lokale Arbeitsmappe, nicht im GitHub-Repository enthalten |
| `tests/test_embeddings_vorbereitung.py` | Gezielte Prüfungen der Eingabe- und Verarbeitungslogik |

## Datengrundlage

Die reguläre Auswahl umfasst 53 Items, davon 34 Rumination und 19 Hostility. Pro Item wird eine Basisfassung verwendet. Wenn die Arbeitsfassung in C ausgefüllt ist, wird sie ausgewählt. Sonst wird der Originaltext aus B verwendet. Die Varianten stehen in G bis I. Die Auswahl erfolgt anhand der Spaltenüberschriften innerhalb der benannten Excel-Tabelle `ItemsAugmentierung`.

Das ergibt 53 Basisfassungen und 158 Varianten, zusammen 211 reguläre Texte. 17 Basisfassungen stammen aus C und 36 aus B. RTS-D8 erhält aufgrund seiner Länge zwei Varianten. Die übrigen Items besitzen jeweils drei Varianten.

Die fünf ursprünglich inversen Items haben in C eine textuell umgepolte Analysefassung. Ihre ursprünglichen niedrigen Originaltexte werden nicht zusätzlich eingebettet oder mit den hohen Fassungen gemittelt. Die Analysepolung K muss bei allen regulären Items direkt sein. Die Codes mit `inv.` bleiben stabile Kennungen und lösen keine weitere Invertierung aus. RRQ3 ist in der regulären Auswahl ausgeschlossen.

Itemcodes, Fragebogennamen, Konstruktbezeichnungen und Polung werden als Metadaten gespeichert. Anmerkungen werden mitgeprüft und bei der Prüfsumme des Quelldatensatzes berücksichtigt, ihr vollständiger Inhalt wird nicht in die Zuordnungstabelle kopiert. Diese Angaben sind keine Bestandteile des Textes, den das Modell erhält. Die Textvorverarbeitung entfernt ausschließlich äußeren Leerraum. Wortlaut und innerer Leerraum bleiben erhalten.

## Kontrolltext

Die Arbeitsmappe enthält ein getrenntes Blatt `Kontrolltexte` mit der Tabelle `KontrolltexteTabelle`. Die Spalten heißen `ID`, `Rolle`, `Text` und `Anmerkung`.

KONT01 lautet: „Ich bewahre meine Tassen im Küchenschrank auf.“ Der Satz wurde für diesen Durchgang als neutraler Prüftext formuliert. Er dient einer späteren groben semantischen Plausibilitätsprüfung und liefert keinen unabhängigen Nachweis psychometrischer Validität.

Noch kein zuvor ausgeschlossenes Fragebogenitem wurde als Grenzfall hinzugefügt. Solche Texte können später mit der Rolle `grenzfall` aufgenommen werden. Dann ist `expected_control_count` in der Konfiguration anzupassen. Grenzfälle bleiben außerhalb der regulären 53-Item-Auswahl.

Kontrolltexte und Grenzfälle erhalten eigene Vektoren und werden getrennt gespeichert. Sie bestimmen weder die 53 regulären Itemmittel noch den späteren Zentrierungsmittelwert mit.

## Nur die Eingabe prüfen

Öffne einen Terminal im Ordner dieser Dateien. Voraussetzung für die reine Prüfung ist eine Python-Umgebung mit `openpyxl`. Für numerische Tests wird zusätzlich NumPy benötigt.

```text
python 01_embeddings.py --config config_embeddings.json --pruefen --bericht-dir runs/vorpruefung
```

Auch der Aufruf ohne `--pruefen` bleibt eine reine Prüfung. Fehlende Modellbibliotheken werden im Prüfbericht gemeldet. Sie führen bei korrekt lesbarer Eingabe nicht zu einer falschen Aussage, die Daten selbst seien fehlerhaft. Ein späterer Berechnungslauf erfordert dagegen die vollständigen Laufzeitabhängigkeiten.

Die Prüfung schreibt bei angegebenem Berichtordner einen Prüfbericht und die Zuordnung der tatsächlich ausgewählten Texte. Sie erzeugt keine `.npy`-Dateien. Der Bericht unterscheidet Dateneignung, Modellverfügbarkeit und tatsächlich ausgeführte Schritte.

Die Konfigurationsdatei wird standardmäßig neben dem Skript gesucht. Relative Ein- und Ausgabepfade beziehen sich auf den Ordner der verwendeten Konfiguration. Dadurch hängt die Textauswahl nicht vom gerade geöffneten Terminalordner ab.

## Vorgesehener späterer Berechnungslauf

Die Berechnung ist ausschließlich mit der ausdrücklichen Option `--berechnen` vorgesehen. Sie wurde in diesem Vorbereitungsschritt nicht gestartet. Ein solcher späterer Aufruf kann den erstmaligen Download des Modells auslösen.

Das eingestellte Modell ist `sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2`. Die festgehaltene Revision lautet `e8f8c211226b894fcb81acc59f3b34ba3efd5f42`. Diese Revision wurde am 29.10.2026 anhand der [öffentlichen Modellmetadaten](https://huggingface.co/api/models/sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2) ermittelt. Sie ist eine Festlegung für diesen neuen Lauf. Es wird nicht behauptet, dass Hütter oder Suwita exakt diese Revision benutzt haben.

Es wird keine zusätzliche Embedding-Instruktion vorangestellt. Vorgesehen sind CPU-Ausführung und 16 Texte pro Stapel. Vor der Einbettung muss das Skript die tatsächlichen Tokenlängen prüfen und bei Überschreitung der Modellgrenze stoppen. Texte dürfen nicht unbemerkt abgeschnitten werden.

Die Formulierungen werden zunächst ohne nachträgliche L2-Normierung ausgegeben und gespeichert. Aus diesen ursprünglichen Vektoren entstehen zusätzlich L2-normierte Einzelvektoren. Entsprechend dem Referenzverfahren wird pro regulärem Item über seine normierten Formulierungen gemittelt und dieses Itemmittel erneut L2-normiert.

Die vorgesehenen Ergebnisdateien sind:

| Datei | Inhalt |
| --- | --- |
| `regular_raw.npy` | Unveränderte Einzelvektoren der 211 regulären Formulierungen |
| `regular_l2.npy` | Dieselben Einzelvektoren, nachträglich auf Länge eins normiert |
| `item_mean_of_raw.npy` | 53 Mittelwerte der ursprünglichen Einzelvektoren für spätere methodische Vergleiche |
| `item_mean_of_l2.npy` | 53 Mittelwerte der zuvor normierten Formulierungen, selbst noch nicht erneut normiert |
| `item_mean_of_l2_l2.npy` | 53 erneut normierte Itemmittel, vorgesehener Hauptpfad nach dem Referenzverfahren |
| `controls_raw.npy` | Ursprüngliche Vektoren der separat gekennzeichneten Prüftexte |
| `controls_l2.npy` | Entsprechende normierte Prüftextvektoren |

Die Zuordnungstabellen halten fest, welche Matrixzeile zu welcher Formulierung oder welchem Item gehört. Im Laufprotokoll werden Eingabe- und Skriptprüfsummen, Modellrevision, Einstellungen, Python-Version und tatsächlich installierte Paketversionen erfasst. Die Anforderungsliste ist kein vollständiger bereits getesteter Lockfile. Es wurde keine Embedding-Umgebung installiert oder ein Modelllauf durchgeführt.

## Spätere Schritte

SQuID-Zentrierung, Achsenprojektion, Kosinusvergleiche und Itemauswahl werden in diesem Skript nicht ausgeführt.

Bei der späteren Zentrierung ist zunächst ein gemeinsamer Mittelpunkt aus genau 53 gleich gewichteten regulären Itemrepräsentationen zu bilden. Ein ungewichteter Mittelpunkt aus allen 211 Formulierungen würde Items unterschiedlich gewichten. Kontrollen und Grenzfälle werden anschließend mit demselben gespeicherten Mittelpunkt verarbeitet, ohne dessen Schätzung zu verändern.

Die Reihenfolge von Einzelvektornormierung, Mittelung, Zentrierung und Projektion ist eine methodische Entscheidung. Eine nachträgliche Änderung wird ausdrücklich dokumentiert. Die ursprünglichen Einzelvektoren bleiben dafür verfügbar. Zahl und Inhalte der Ruminations- und Hostility-Achsen müssen noch festgelegt werden.

## Herkunft und KI-Unterstützung

Die ursprünglichen Skripte und ihre Lizenzhinweise bleiben im Fork erhalten. Die hier ergänzte Verarbeitung wurde mit KI-Unterstützung für die Masterarbeit angepasst. Sie ist keine unverändert ausgeführte Kopie der gesamten ursprünglichen Pipeline. Automatisierte Daten- und Logikprüfungen ersetzen keinen empirischen Nachweis der Validität der späteren Podcastwerte.
