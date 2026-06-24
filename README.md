# Word Topic Extractor

A small Python project that converts one complete Word file into a new Word file containing only a selected topic or section.

## What it does

- Reads a `.docx` Word file.
- Reads a Table of Contents when the document has one.
- Shows topics with page numbers from the contents table.
- Detects document sections from Word heading styles such as `Heading 1`, `Heading 2`, and `Heading 3`.
- Finds the topic you choose.
- Copies that topic and its content into a new `.docx` file.
- Preserves paragraphs, headings, basic text formatting, tables, and images.
- Includes a deployable web app so users can access it from a link.

## Project files

- `word_topic_extractor.py` - main command-line app.
- `word_topic_extractor_gui.py` - professional Honeywell-styled desktop GUI app.
- `web_app.py` - Flask web app for online deployment.
- `templates/` - web page templates.
- `static/` - web styling.
- `Procfile` - web hosting startup command.
- `create_sample_doc.py` - creates a sample Word file for testing.
- `requirements.txt` - Python package list.
- `outputs/` - place for generated Word files.

## Quick start

Create a sample Word document:

```powershell
python create_sample_doc.py
```

Open the GUI:

```powershell
python word_topic_extractor_gui.py
```

On Windows, you can also run:

```powershell
run_gui.bat
```

Open the web app locally:

```powershell
pip install -r requirements.txt
python web_app.py
```

Then open:

```text
http://127.0.0.1:5000
```

In the GUI:

1. Click `Upload Word File`.
2. Select a `.docx` file.
3. Search for a topic if the contents table is long.
4. Highlight one or more rows.
5. Click `Add Highlighted`, or click `Select Visible` to choose all visible search results.
6. Choose where to save the new Word file.
7. Click `Create New Word File`.

The GUI keeps selected topics even when you search again, so you can select topics from different parts of a long contents table.

## Online deployment

This project is ready for Flask hosting. On platforms such as Render, Railway, or Heroku-style hosts:

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn web_app:app`
- Python entry file: `web_app.py`

After deployment, the hosting platform gives you a public link. Anyone who opens that link can upload a Word file, select topics, and download the generated Word file.

Extract a topic from it:

```powershell
python word_topic_extractor.py sample_complete_document.docx "Machine Learning" -o outputs/machine_learning_topic.docx
```

List all available topics:

```powershell
python word_topic_extractor.py sample_complete_document.docx --list-topics
```

## Use your own Word file

```powershell
python word_topic_extractor.py "C:\path\to\complete_file.docx" "Selected Topic" -o outputs/selected_topic.docx
```

## Notes

The app works best when the Word file has a Table of Contents and uses real heading styles. For example:

- Heading 1: Main topic
- Heading 2: Subtopic
- Normal: Body content

The page number shown in the app comes from the Word file's contents table. The actual extraction is done by matching that topic to the document heading and copying everything until the next heading at the same or higher level.

If a document does not use a contents table, the app falls back to heading-style topics.

Images are preserved by keeping the original Word file package and removing unselected sections, instead of rebuilding the output from a blank document.
