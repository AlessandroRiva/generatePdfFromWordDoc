import re
import unicodedata
from pathlib import Path
from docx import Document
from playwright.sync_api import sync_playwright

# ----------------- Filename sanitizer -----------------
def sanitize_filename(title):
    title = unicodedata.normalize('NFKD', title).encode('ASCII', 'ignore').decode('utf-8')
    title = title.lower().strip()
    title = re.sub(r'[^a-z0-9\s-]', '', title)
    title = re.sub(r'[\s-]+', '-', title)
    return title

# ----------------- Extract hyperlinks -----------------
def extract_hyperlinks(docx_path):
    """
    Returns tuples of (nested_folder_path, url)
    folder structure mirrors heading levels
    """
    doc = Document(docx_path)
    links = []
    heading_stack = []

    for para in doc.paragraphs:
        text = para.text.strip()

        # Detect headings like 1, 1.1, 1.2
        if text and all(c.isdigit() or c == '.' for c in text):
            levels = text.split(".")
            heading_stack = levels
            continue

        # Extract hyperlinks from XML
        for rel in para._element.xpath(".//w:hyperlink"):
            r_id = rel.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if not r_id:
                continue
            url = doc.part.rels[r_id]._target

            # Build nested folder path
            folder_parts = []
            for i, level in enumerate(heading_stack):
                if i == 0:
                    folder_parts.append(level)         # top-level folder: "1"
                else:
                    folder_parts.append(f"{folder_parts[-1]}-{level}")  # nested: "1-1", "1-2"
            nested_folder_path = Path(*folder_parts) if folder_parts else Path("root")

            links.append((nested_folder_path, url))

    return links

# ----------------- Save webpage as PDF -----------------
def save_webpage_as_pdf(url, folder_path):
    folder_path.mkdir(parents=True, exist_ok=True)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.set_viewport_size({"width": 1920, "height": 1080})
            page.goto(url, wait_until="load", timeout=60000)
            page.emulate_media(media="screen")
            page.evaluate("document.fonts.ready")

            # Get page title and sanitized filename
            title = page.title()
            filename = sanitize_filename(title) + ".pdf"
            output_path = folder_path / filename

            # Get full page height
            height = page.evaluate("document.body.scrollHeight")

            # Save PDF
            page.pdf(
                path=str(output_path),
                width="1920px",
                height=f"{height}px",
                print_background=True,
                scale=1.2
            )
            browser.close()
        print(f"✅ Saved PDF: {output_path}")
    except Exception as e:
        print(f"❌ Failed for {url}: {e}")

# ----------------- Main -----------------
def main(docx_path):
    links = extract_hyperlinks(docx_path)
    for folder_path, url in links:
        save_webpage_as_pdf(url, folder_path)

if __name__ == "__main__":
    main("example.docx")  # <-- your Word file here
