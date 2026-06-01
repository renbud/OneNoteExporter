"""Main entry point for the OneNote Exporter application."""
from datetime import datetime
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
from pivot import load_pivot, save_pivot

# Configure logging
def setup_logging(log_level="INFO"):
    """Set up logging configuration from config or default."""
    log_format = "%(asctime)s - %(levelname)s - %(message)s"
    
    if log_level == "DEBUG":
        level = logging.DEBUG
    elif log_level == "WARNING":
        level = logging.WARNING
    elif log_level == "ERROR":
        level = logging.ERROR
    else:
        level = logging.INFO
    
    logging.basicConfig(
        level=level,
        format=log_format,
        datefmt="%Y-%m-%d %H:%M:%S"
    )

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
    """Load export_root, CLIENT_ID, and SCOPES from config.yaml."""
    config_path = Path("config/config.yaml")

    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)

        # Validate required keys
        if not isinstance(config, dict):
            logging.error("config.yaml is not a valid YAML mapping.")
            return None

        if "export_root" not in config or "CLIENT_ID" not in config:
            logging.error("config.yaml must contain 'export_root' and 'CLIENT_ID'.")
            return None

        app_config = {
            "export_root": Path(config["export_root"]),
            "CLIENT_ID": config["CLIENT_ID"]
        }

        if "SCOPES" in config:
            app_config["SCOPES"] = config["SCOPES"]

        # Load log_level from config
        log_level = config.get("log_level", "INFO")
        setup_logging(log_level)

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
        logging.critical("Failed to load configuration.")
        exit(1)

    OUTPUT_FOLDER = config["export_root"]

    app = msal.PublicClientApplication(
        client_id=config["CLIENT_ID"],
        authority="https://login.microsoftonline.com/consumers"
    )

    SCOPES = config.get("SCOPES", ["Notes.Read", "Notes.Read.All"])


    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        logging.error("Device flow initiation failed:")
        logging.error(flow)
        return

    logging.info("FLOW: %s", flow)

    # Be careful with this link. It can depend on the app type you have set up in Azure AD.
    # For a multi-tenant app, the link is https://microsoft.com/devicelogin.
    # For a single-tenant app, it might be https://microsoft.com/devicelogin?tenant=YOUR_TENANT_ID
    # For consumer apps, it should be https://www.microsoft.com/link.
    logging.info("Go to https://www.microsoft.com/link and enter this code:")
    logging.info(flow["user_code"])

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        logging.error("Authentication failed:")
        logging.error(result)
        return

    token = result["access_token"]
    export_all(OUTPUT_FOLDER, token)

def export_all(root, token):
    pivot_ts = load_pivot()
    
    notebooks = get_notebooks(token)

    max_ts = pivot_ts
    for nb in notebooks:
        logging.info("Notebook: %s", nb["displayName"])
        sections = get_sections(nb["id"], token)

        for sec in sections:
            logging.info("  Section: %s", sec["displayName"])
            pages = get_pages(sec["id"], token)

            for pg in pages:
                modified_datetime = pg.get("lastModifiedDateTime")
                if not modified_datetime:
                    continue
                    
                dt = datetime.fromisoformat(modified_datetime.replace("Z", "+00:00"))
                max_ts = max(max_ts, dt.timestamp())
                
                # Only export if page is new or modified since last run
                if dt.timestamp() <= pivot_ts:
                    logging.info("    Page skipped (unchanged): %s", pg["title"])
                    continue
                
                logging.info("    Page exported: %s", pg["title"])
                logging.info("    Page: %s Modified: %s", pg["title"], modified_datetime)
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
            
    # Save the updated pivot timestamp
    save_pivot(max_ts)

if __name__ == "__main__":
    main()
