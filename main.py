"""Main entry point for the OneNote Exporter application."""
import msal
import re
from onenote_extract import (
    download_resource,
    get_notebooks,
    get_sections,
    get_pages,
    get_page_html,
    get_page_resources,
    extract_attachments_from_html,
    download_attachment,
)
from markdown_save import html_to_markdown, write_markdown, save_attachment

OUTPUT_FOLDER = "D:\\OneNoteExportedFiles"
CLIENT_ID = "558076f0-908d-4247-8334-97f229b74d97"
SCOPES = ["Notes.Read", "Notes.Read.All"]   

"""
The expected output structure is as follows:
<root_output_folder>/
    Notebook1/
        SectionA/
            Page1.md
            Page2.md
        SectionB/
            Page1.md
    Notebook2/
        SectionX/
            Page1.md

"""

def main():
    print("Hello from onenoteexporter!")

    # app = msal.PublicClientApplication(
    #     client_id=CLIENT_ID,
    #     authority=f"https://login.microsoftonline.com/common"
    # )

    # app = msal.PublicClientApplication(
    #     client_id=CLIENT_ID,
    #     authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    # )

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority="https://login.microsoftonline.com/consumers"
    )


    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        print("Device flow initiation failed:")
        print(flow)
        return

    print("FLOW:", flow)

    # Be careful with this link. It can depend on the app type you have set up in Azure AD.
    # For a multi-tenant app, the link is https://microsoft.com/devicelogin.
    # For a single-tenant app, it might be https://microsoft.com/devicelogin?tenant=YOUR_TENANT_ID
    # For consumer apps, it should be https://www.microsoft.com/link.
    print("Go to https://www.microsoft.com/link and enter this code:")
    print(flow["user_code"])

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print("Authentication failed:")
        print(result)
        return

    token = result["access_token"]
    export_all(OUTPUT_FOLDER, token)

    # notebooks = get_notebooks(token)

    # for nb in notebooks:
    #     print(f"DisplayName: {nb['displayName']}")
    #     print("Notebook ID:", nb["id"])
    #     print(nb["links"]["oneNoteWebUrl"]["href"])
    #     sections = get_sections(nb["id"], token)
    #     for sec in sections:
    #         print(f"  Section: {sec['displayName']}")
    #         print("  Section ID:", sec["id"])
    #         pages = get_pages(sec["id"], token)
    #         for pg in pages:
    #             print(f"    Page: {pg['title']}")
    #             print("    Page ID:", pg["id"])


def export_all(root, token):
    notebooks = get_notebooks(token)

    for nb in notebooks:
        print("Notebook:", nb["displayName"])
        sections = get_sections(nb["id"], token)

        for sec in sections:
            print("  Section:", sec["displayName"])
            pages = get_pages(sec["id"], token)

            for pg in pages:
                modified_datetime = pg.get("lastModifiedDateTime")
                print("    Page:", pg["title"], "Modified:", modified_datetime)
                html = get_page_html(pg, token)
                md_text = html_to_markdown(html)

                attachments = extract_attachments_from_html(html)

                attNo = 0
                for att in attachments:
                    mime = att["mime"]
                    url = att["url"]

                    # Determine extension
                    if mime == "application/pdf":
                        ext = ".pdf"
                    elif mime in (
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "application/vnd.ms-excel",
                        "application/vnd.ms-excel.sheet.macroEnabled.12"
                    ):
                        ext = ".xlsx"
                    elif mime in (
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "application/msword"
                    ):
                        ext = ".docx"
                    else:
                        # Unsupported attachment type
                        attNo += 1
                        continue

                    # Determine filename
                    filename = att["filename"]
                    if filename == "attachment":
                        filename = f"attachment_{attNo}{ext}"

                    attNo += 1

                    # Download and save
                    data = download_attachment(url, token)
                    save_attachment(root, nb, sec, pg, filename, data, modified_datetime)

                    # Append correct link
                    md_text += f"\n\n[Attached file: {filename}]({filename})\n"

                # Write markdown once all attachments are processed, so that the links are correct
                write_markdown(root, nb, sec, pg, md_text, modified_datetime)

if __name__ == "__main__":
    main()
