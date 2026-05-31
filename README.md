# Python project to export OneNote pages to Markdown files
* Uses Microsoft Graph
* Authenticates with registered Azure application Id
* Exports all sections of all notebooks
* Creates a markdown file for each page in a folder structure that follows OneNote
* Exports embedded pdfs docx and Excel files
* Handles throttling by implementing a retry mechanism with exponential backoff.

```
<root_output_folder>
    Notebook1/
        SectionA/
            Page1.md
            Page2.md
        SectionB/
            Page1.md
    Notebook2/
        SectionX/
            Page1.md
```

## Setup
### Install all dependencies
 uv sync

## Configure

1. Copy `config/config.yaml.template` to `config/config.yaml`
2. Edit `config/config.yaml`:
   - Set `export_root:` to your desired folder path
   - Add your Azure Application `CLIENT_ID` from app registration
   - Optionally set `SCOPES` (defaults to Notes.Read and Notes.Read.All)

## Azure App Permissions

The app uses Microsoft Graph API with the following permissions:
- Notes.Read
- Notes.Read.All

## Usage

### First run (or after adding new notebooks)
Run the project to export all OneNote notebooks:

```bash
uv run python main.py
```

### Incremental updates
Subsequent runs will only export new or modified pages based on the last export timestamp. This makes re-runs fast and efficient.

## How to register the app in Azure

This is the hardest part of this whole thing. It's free but you need to register yourself in Azure.
With a personal Microsoft account, I set up a tenant in Azure. Then added an app registration.

https://portal.azure.com/?quickstart=True#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade

### Required permissions for the Azure app:
- Notes.Read (delegated - user's consent)
- Notes.Read.All (application - admin consent required)

When it is registered, you'll see an Application (client) ID with a GUID-like Id.
Copy the ID and paste it into `config/config.yaml` as:
```yaml
CLIENT_ID: "your-client-id-here"
```

### Run the project

```bash
uv run python main.py
```

### Usage examples

```bash
# Export all notebooks (first run or full export)
uv run python main.py

# Incremental update (only new/modified pages)
uv run python main.py
```