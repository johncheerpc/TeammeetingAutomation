# Teams Transcript Processing - Local Version

A standalone Python script that runs the SAME transcript-processing pipeline as
the Azure Functions version, but from your own machine. You pass the queue
message in as input instead of waiting for a real Azure Queue trigger.

## Project structure

```
teams-transcript-local/
├── main.py                  # thin ENTRY POINT only (bootstrap + handoff)
├── .env.example             # copy to .env and fill in your settings
├── .gitignore
├── requirements.txt
├── sample_message.json      # example message for a quick test
├── .vscode/
│   └── settings.json        # makes Pylance resolve the `shared` package
└── shared/                  # all the real modules live here
    ├── __init__.py
    ├── config.py            # the configuration VARIABLES (read from .env)
    ├── runner.py            # CLI args, logging, reading the message, run()
    ├── models.py            # data containers
    ├── message_parser.py    # Step 1
    ├── sharepoint_client.py # Steps 2, 5, 6
    ├── ai_processor.py      # Step 4 (processAI)
    └── processor.py         # the shared 6-step pipeline
```

## Setup (one time)

```bash
# 1. (recommended) create and activate a virtual environment
python -m venv .venv
# Windows:        .venv\Scripts\activate
# macOS / Linux:  source .venv/bin/activate

# 2. install dependencies
pip install -r requirements.txt

# 3. create your settings file and fill in the values
#    Windows:   copy .env.example .env
#    macOS/Linux: cp .env.example .env
#    then edit .env with your SharePoint site, credentials, and (optional) AI keys
```

## Run it

Pass the queue message in any of these ways:

```bash
# from a file (positional)
python main.py message.json

# from a file (named flag)
python main.py --file message.json

# inline JSON string
python main.py --message '{"value":[{ ... }]}'

# piped via stdin
cat message.json | python main.py            # macOS / Linux
Get-Content message.json | python main.py     # Windows PowerShell

# no argument -> uses sample_message.json (quick smoke test)
python main.py

# add -v / --verbose for extra DEBUG logging
python main.py message.json -v
```

On success you'll see the step-by-step log followed by a JSON `RESULT:` summary.

## What it does (the 6 steps, in shared/processor.py)

1. Extract JoinWebUrl, MeetingId, transcript info, duration, date from the message.
2. Look up the SharePoint list row by JoinWebUrl + MeetingId and read
   AdminUpdatedFolderLink, FileToBeSaved, RecordingsLink, TranscriptLink.
3. Log an audit snapshot.
4. Read the source document and run `processAI` on it.
5. Save the AI output to FileToBeSaved.
6. Move the recordings to RecordingsLink.

## Differences from the Azure version

| | Azure version | Local version (this) |
|---|---|---|
| Entry point | `function_app.py` (triggers) | `main.py` (you run it) |
| How it starts | HTTP request or queue message | you run `python main.py` |
| Where settings live | Azure "Application settings" | `.env` file / env vars |
| Logging | captured by Azure runtime | printed to your terminal |
| Shared logic | `shared/` | `shared/` (identical) |

## Notes

- If `AIEndpoint`/`AIApiKey`/`AIDeployment` are left blank, `processAI` returns a
  labelled placeholder so you can test the whole flow before wiring up a model.
- The SharePoint app registration needs permission on the site (typically
  `Sites.ReadWrite.All`, granted by an admin).
- Verify the four column INTERNAL names match your list; if a value comes back
  as None, a display-vs-internal name mismatch is the usual cause.
