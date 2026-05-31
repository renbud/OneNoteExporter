"""Main entry point for the OneNote Exporter application."""
import logging
from logging import config
import msal
import yaml
from pathlib import Path
from onenote_extract import (
    get_notebooks,
    get_sections,
    get_pages,
    get_page_html,
    extract_attachments_from_html,
    download_attachment,
)
from markdown_save import html_to_markdown, write_markdown, save_attachment

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)



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

def load_config():
    """Load gmail_user, export_root, CLIENT_ID, and SCOPES from config.yaml."""
    config_path = Path("config/config.yaml")

    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)

        # Validate required keys
        if not isinstance(config, dict):
            logging.error("config.yaml is not a valid YAML mapping.")
            return None

        if "gmail_user" not in config or "export_root" not in config or "CLIENT_ID" not in config:
            logging.error("config.yaml must contain 'gmail_user', 'export_root' and 'CLIENT_ID'.")
            return None

        app_config = {
            "gmail_user": config["gmail_user"],
            "export_root": Path(config["export_root"]),
            "CLIENT_ID": config["CLIENT_ID"]
        }

        if "SCOPES" in config:
            app_config["SCOPES"] = config["SCOPES"]

        return app_config

    except FileNotFoundError:
        logging.error("config.yaml not found in current directory.")
        return None

    except yaml.YAMLError as e:
        logging.critical(f"Error parsing config.yaml: {e}")
        return None


def main():
    config = load_config()
    if not config:
        logger.critical("Failed to load configuration.")
        exit(1)

    OUTPUT_FOLDER = config["export_root"]

    app = msal.PublicClientApplication(
        client_id=config["CLIENT_ID"],
        authority="https://login.microsoftonline.com/consumers"
    )

    SCOPES = config.get("SCOPES", ["Notes.Read", "Notes.Read.All"])


    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        logger.error("Device flow initiation failed:")
        logger.error(flow)
        return

    logger.info("FLOW: %s", flow)

    # Be careful with this link. It can depend on the app type you have set up in Azure AD.
    # For a multi-tenant app, the link is https://microsoft.com/devicelogin.
    # For a single-tenant app, it might be https://microsoft.com/devicelogin?tenant=YOUR_TENANT_ID
    # For consumer apps, it should be https://www.microsoft.com/link.
    logger.info("Go to https://www.microsoft.com/link and enter this code:")
    logger.info(flow["user_code"])

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        logger.error("Authentication failed:")
        logger.error(result)
        return

    token = result["access_token"]
    export_all(OUTPUT_FOLDER, token)

def export_all(root, token):
    notebooks = get_notebooks(token)

    for nb in notebooks:
        logger.info("Notebook: %s", nb["displayName"])
        sections = get_sections(nb["id"], token)

        for sec in sections:
            logger.info("  Section: %s", sec["displayName"])
            pages = get_pages(sec["id"], token)

            for pg in pages:
                modified_datetime = pg.get("lastModifiedDateTime")
                logger.info("    Page: %s Modified: %s", pg["title"], modified_datetime)
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
