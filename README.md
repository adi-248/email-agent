# Email Agent

Read-only Gmail intelligence. This project **never modifies Gmail**.

It reads messages, classifies each one with a local Ollama model, and writes an Excel report. It does not send, delete, move, label, archive, or download attachments.

## What you need

- Python 3.10+
- [Ollama](https://ollama.com/) running locally with `qwen2.5:3b` (or change `OLLAMA_MODEL` in `email_agent.py`)
- A Google Cloud OAuth **Desktop** client with the `gmail.readonly` scope

## Setup

```text
pip install -r requirements.txt
```

1. In Google Cloud, create an OAuth client (Desktop app).
2. Download the client JSON.
3. Copy `credentials/credentials.json.example` to `credentials/credentials.json` and replace the placeholders with that client JSON.
4. Start Ollama and confirm the model is installed:

```text
ollama list
```

## Run

```text
python email_agent.py
```

The first run opens a browser so you can authorize Gmail. The token is stored as `credentials/token.json` in this folder only.

Reports are written to `reports/`. Start with `MAX_MESSAGES = 10` in `email_agent.py`, then raise it after you have reviewed the output.

## Safety

Gmail scope: `https://www.googleapis.com/auth/gmail.readonly`

This script cannot delete, move, archive, label, or send email.

## Paths

All paths are relative to this project folder. There is no hardcoded machine-specific root.
