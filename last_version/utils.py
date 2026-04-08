import os
from datetime import datetime

def read_urls(path="inputs/urls.txt"):
    """
    Reads the input file containing URLs.
    - Skips empty lines.
    - Returns a clean list of URLs.
    """
    urls = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                urls.append(line)
    return urls


def get_timestamp():
    """
    Generates a readable timestamp for filenames.
    Format: YYYYMMDD_HHMMSS
    """
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def extract_instagram_id(url):
    """
    Extracts the Instagram reel/video ID from the URL.
    Example:
        https://www.instagram.com/reel/CqXyZtA123/
    Returns:
        CqXyZtA123
    """
    try:
        parts = url.rstrip("/").split("/")
        return parts[-1]
    except:
        return "unknown"


def ensure_folder(path):
    """
    Creates a folder if it does not already exist.
    Prevents errors when saving files.
    """
    if not os.path.exists(path):
        os.makedirs(path)


def ensure_reports_folder():
    """
    Ensures that the 'reports' folder exists.
    Returns the folder path.
    """
    reports_dir = "reports"
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
    return reports_dir
