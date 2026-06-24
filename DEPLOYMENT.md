# Online Deployment Guide

Use this guide to deploy the Word Topic Extractor as a web app.

## Recommended host settings

Use any Python web host that supports Flask apps.

Common examples:

- Render
- Railway
- Heroku-style platforms
- A VPS with Python and Gunicorn

## Required commands

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn web_app:app
```

## Environment variables

Optional but recommended:

```text
SECRET_KEY=change-this-to-a-long-random-value
MAX_UPLOAD_MB=25
```

`MAX_UPLOAD_MB` controls the maximum upload size in megabytes.

## How users will use it

1. Open the deployed link.
2. Upload a `.docx` Word file.
3. Search and select one or more topics.
4. Click `Create and Download Word File`.
5. Download the generated Word file.

## Important note

Uploaded and generated files are stored on the server filesystem. For large production use, connect cloud storage or add a scheduled cleanup job.
