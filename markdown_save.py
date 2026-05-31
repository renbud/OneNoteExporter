from datetime import datetime
from markdownify import markdownify as md
import logging
import os
import re

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def safe_name(name):
    name = name.strip()
    name = re.sub(r'[\\/*?:"<>|]', '_', name)
    return name


def html_to_markdown(html):
    return md(html, heading_style="ATX")


def write_markdown(root, notebook, section, page, markdown, modified_datetime):
    nb = safe_name(notebook["displayName"])
    sec = safe_name(section["displayName"])
    pg = safe_name(page["title"])

    folder = os.path.join(root, nb, sec)
    os.makedirs(folder, exist_ok=True)

    path = os.path.join(folder, f"{pg}.md")

    # Write the markdown file
    with open(path, "w", encoding="utf-8") as f:
        f.write(markdown)

    # Convert ISO 8601 timestamp → epoch seconds
    # Example OneNote format: "2019-02-12T10:15:23Z"
    dt = datetime.fromisoformat(
        modified_datetime.replace("Z", "+00:00")
    )
    epoch = dt.timestamp()

    # Set both access and modification time
    os.utime(path, (epoch, epoch))

    logger.info(f"Wrote: {path}")


def save_attachment(root, notebook, section, page, filename, data, modified_datetime):
    nb = safe_name(notebook["displayName"])
    sec = safe_name(section["displayName"])
    pg = safe_name(page["title"])
    safe_filename = pg + "_" + safe_name(filename)

    folder = os.path.join(root, nb, sec)
    os.makedirs(folder, exist_ok=True)

    path = os.path.join(folder, safe_filename)

    with open(path, "wb") as f:
        f.write(data)

    # Convert ISO 8601 timestamp → epoch seconds
    dt = datetime.fromisoformat(
        modified_datetime.replace("Z", "+00:00")
    )
    epoch = dt.timestamp()

    # Set both access and modification time
    os.utime(path, (epoch, epoch))

    logger.info(f"Saved attachment: {path}")
