# Teams Transcript Processing Function

Two Azure Functions (Python v2 model) that share one processing pipeline:

- **`manual_trigger`** – HTTP trigger for local dev/testing from VS Code.
- **`queue_trigger`** – Queue trigger for production; fires on each new message.

Both call `shared.processor.process_message`, so there is **no duplicated logic**.

## Project structure

```
teams-transcript-func/
├── function_app.py            # both triggers (entry point)
├── host.json
├── local.settings.json        # local config (DO NOT commit real secrets)
├── requirements.txt
├── sample_message.json        # example queue message for testing
└── shared/
    ├── config.py              # reads Function App Settings
    ├── models.py              # MeetingInfo, SharePointMetadata
    ├── message_parser.py      # Step 1: extract meeting info
    ├── sharepoint_client.py   # Steps 2, 5, 6: SharePoint read/upload/move
    ├── ai_processor.py        # Step 4: processAI(document_content) -> bytes
    └── processor.py           # shared 6-step orchestrator
```

## Pipeline (in `processor.process_message`)

1. **Extract** JoinWebUrl, MeetingId, transcript info, duration, date from the message.
2. **Retrieve** `AdminUpdatedFolderLink`, `FileToBeSaved`, `RecordingsLink`,
   `TranscriptLink` from the SharePoint list, filtered by `JoinWebUrl` + `MeetingId`.
3. **Log** an audit snapshot of everything gathered.
4. **processAI** – read the doc at `AdminUpdatedFolderLink`, run AI, get output.
5. **Save** the output to `FileToBeSaved`.
6. **Move** recordings to `RecordingsLink`.

## Application settings

Set these in the Function App (or `local.settings.json` locally):

| Setting | Purpose |
|---|---|
| `QueueConnection` | Storage connection string for the queue trigger |
| `QueueName` | Queue to listen on |
| `SharePointSiteUrl` | Target SharePoint site |
| `SharePointListName` | Custom list name |
| `SharePointTenantId` / `SharePointClientId` / `SharePointClientSecret` | App-only auth |
| `AIEndpoint` / `AIApiKey` / `AIDeployment` | Azure OpenAI (optional; placeholder if unset) |

The SharePoint app registration needs `Sites.ReadWrite.All` (or site-scoped equivalent).

## Run & test locally

```bash
pip install -r requirements.txt
func start
```

Invoke the manual trigger with the sample message:

```bash
curl -X POST http://localhost:7071/api/process-transcript \
  -H "Content-Type: application/json" \
  --data @sample_message.json
```

When validated, stop using the HTTP endpoint and let `queue_trigger` run automatically.

## Notes

- `processAI` ships with a guarded Azure OpenAI example; replace `_run_model`
  in `shared/ai_processor.py` with your real model call if different.
- Unhandled exceptions in `queue_trigger` cause the runtime to retry and
  eventually poison-queue the message, which preserves it for troubleshooting.
