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

## How to register the app in Azure
This is the hardest part of this whole thing.
With my personal microsoft account - I set up a tenant in Azure

and then added an app registration
https://portal.azure.com/?quickstart=True#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/558076f0-908d-4247-8334-97f229b74d97

It is visible under App Registrations
https://portal.azure.com/?quickstart=True#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade


