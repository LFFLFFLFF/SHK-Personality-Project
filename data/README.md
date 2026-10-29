# Lokale Eingabedatei

Lege hier die vorbereitete Datei `Itemdaten_mit_Kontrolltext.xlsx` ab.

Die Arbeitsmappe enthält die vorhandenen Blätter `Items`, `Protokoll` und `Invertierung` sowie das neue Blatt `Kontrolltexte`.

Die Datei wird lokal eingelesen. Sie ist durch `.gitignore` von der Veröffentlichung im Repository ausgenommen. Als Quelle für die aktuelle Arbeitsfassung dient die zuletzt erstellte Mappe mit 53 Items und textuell umgepolten Analysefassungen.

Das Blatt `Items` enthält 53 reguläre Items und 158 Varianten. Das Blatt `Kontrolltexte` enthält zunächst nur KONT01. Ein später hinzugefügter Grenzfall bleibt von der regulären Auswahl getrennt. Bei geänderter Anzahl der Kontrollen muss `expected_control_count` in der Konfiguration angepasst werden.
