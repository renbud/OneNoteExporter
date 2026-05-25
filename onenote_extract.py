import requests
from bs4 import BeautifulSoup
import time

GRAPH = "https://graph.microsoft.com/v1.0"


def graph_get(url, token, max_retries=7):
    headers = {"Authorization": f"Bearer {token}"}

    for attempt in range(max_retries):
        r = requests.get(url, headers=headers)

        if r.status_code != 429:
            r.raise_for_status()
            return r

        # Throttled
        retry_after = r.headers.get("Retry-After")

        if retry_after:
            wait = int(retry_after)
        else:
            wait = 2 ** attempt  # exponential fallback

        print(f"Throttled. Waiting {wait} seconds…")
        time.sleep(wait)

    r.raise_for_status()
    raise Exception("Too many retries due to throttling")


def get_notebooks(token):
    r = graph_get(f"{GRAPH}/me/onenote/notebooks", token)
    return r.json().get("value", [])

def get_sections(notebook_id, token):
    r = graph_get(f"{GRAPH}/me/onenote/notebooks/{notebook_id}/sections", token)
    return r.json().get("value", [])

def get_pages(section_id, token):
    r = graph_get(f"{GRAPH}/me/onenote/sections/{section_id}/pages", token)
    return r.json().get("value", [])

def get_page_html(page, token):
    url = page["contentUrl"]
    r = graph_get(url, token)
    return r.text

def get_page_resources(page_id, token):
    url = f"{GRAPH}/me/onenote/pages/{page_id}/resources"
    r = graph_get(url, token)
    return r.json().get("value", [])

def download_resource(resource, token):
    url = resource["contentUrl"]
    r = graph_get(url, token)
    return r.content


def extract_attachments_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    attachments = []

    for obj in soup.find_all("object"):
        url = obj.get("data")
        mime = obj.get("type")
        filename = obj.get("data-fullpath") or "attachment"

        if url and mime:
            print("Found attachment in HTML:", filename, mime, url)
            attachments.append({
                "url": url,
                "mime": mime,
                "filename": filename
            })

    return attachments

def download_attachment(url, token):
    r = graph_get(url, token)
    return r.content