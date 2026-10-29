# Masterarbeit zu Rumination und Hostility

Dieser Fork übernimmt das SHK-Personality-Project von Franziska Suwita. Die ursprünglichen Skripte und Lizenzhinweise bleiben erhalten. Für diese Masterarbeit wird die Verarbeitung ergänzt und gesondert dokumentiert.

## Einstieg in die neue Vorbereitung

Lies zuerst [README_Embedding_Vorbereitung.md](README_Embedding_Vorbereitung.md). Die neue Datei `01_embeddings.py` prüft standardmäßig nur die lokale Excel-Datei. Eine Embedding-Berechnung erfordert später ausdrücklich `--berechnen`.

Die Vorbereitung umfasst 53 reguläre Items mit 158 Varianten und einen getrennten neutralen Kontrolltext. In diesem Vorbereitungsschritt werden keine Embeddings erzeugt. SQuID-Zentrierung und Achsenprojektion folgen gesondert.

Die benötigten Bibliotheken für die Ergänzung stehen in `requirements-embeddings.txt`. Die Herkunft, der festgehaltene Ausgangsstand und die Anpassungen stehen in [QUELLEN_UND_AENDERUNGEN.md](QUELLEN_UND_AENDERUNGEN.md). Die lokale Eingabedatei und Laufberichte werden durch `.gitignore` von der Veröffentlichung ausgenommen.

## Ursprüngliche Projektbeschreibung

Der folgende Abschnitt ist die übernommene Beschreibung des Ursprungsprojekts. Die darin genannten API-Schritte gehören zur ursprünglichen Gesamtpipeline. Die neue reine Embedding-Vorbereitung benötigt keine OpenAI- oder Gemini-API.

# SHK-Personality-Project
Project to assess Personality in a naturalistic paradigm 
Dieses Projekt dokumentiert den vollständigen Prozess von der Vorbereitung psychometrischer Items über die semantische Validierung bis hin zur KI-basierten Generierung klinisch-relevanter Alltagsszenen.
Projektübersicht
Die Pipeline ist in vier aufeinanderfolgende Phasen unterteilt:

    Item-Augmentation: Vermehrung der  Items und anschließende Mittelung der Embeddings zur Stabilisierung von SBERT-Embeddings.

    Item-Ähnlichkeitsanalyse: Validierung der Konstruktkohäsion der Original-Items.

    Szenengenerierung: KI-gestützte Erstellung von narrativen Szenen mit integrierter Validierungsschleife.

    Szenen-Validierung: Statistische Überprüfung der finalen Szenen auf konvergente und diskriminante Validität.

Installation & Setup

git clone https://github.com/franziskasuwita-glitch/SHK-Personality-Project.git
cd SHK-Personality-Project

pip install -r requirements.txt

Hinweis: Für das Skript zur Item-Augmentation und zur Szenengenerierung sind API-Keys erforderlich.

1. Item-Augmentierung (item_augmentation.py)
Dieses Skript nutzt GPT-4o-mini, um kurze psychometrische Items semantisch konsistent auf eine Ziel-Wortanzahl zu erweitern. Dies minimiert Längeneffekte bei der anschließenden Vektorisierung.
    Input: Original-Items (direkt im Skript editierbar).
    Output: Konsolidierte Textblöcke für stabilere Embeddings.

2. Item-Ähnlichkeitsanalyse (item_similarity_analysis.py)

Berechnung der internen Konsistenz der Item-Gruppen mittels Sentence-BERT (SBERT).

    Modell: paraphrase-multilingual-MiniLM-L12-v2.

    Metrik: Cosinus-Ähnlichkeit.

    Visualisierung: Heatmaps der Within-Group und Between-Group Ähnlichkeiten.

3. Validierte Szenengenerierung (scene_generator.py)

Generiert narrative Alltagsszenen aus der Ich-Perspektive, die subklinische Ausprägungen von Impulsivität darstellen.

    Technik: Ein rekursiver Algorithmus prüft jede generierte Szene gegen einen Schwellenwert (Cosine Similarity > 0.3).

    KI-Modell: Google Gemini 2.5 Pro.

4. Szenen-Validierungsanalyse (scene_validation.py)

Abschließende statistische Evaluation der generierten Szenen.

    Diskriminanzprüfung: Vergleich der Ähnlichkeit von Szenen zu ihren eigenen Facetten-Definitionen vs. Fremd-Definitionen.

    Output: Export von PNG-Grafiken zur Dokumentation der Konstruktvalidität.

Anforderungen (requirements.txt)

Für die Ausführung werden folgende Bibliotheken benötigt:

    sentence-transformers

    openai

    google-generativeai

    pandas, numpy

    matplotlib, seaborn

    scikit-learn

    langdetect

