import os
import yt_dlp
from datetime import datetime

def make_timestamp():
    """
    Generates a detailed timestamp for filenames.
    Format: YYYY_MM_DD_HH_MM_SS_microseconds
    """
    return datetime.now().strftime("%Y_%m_%d_%H_%M_%S_%f")


def download_instagram_video(url, output_dir="recipes_videos"):
    """
    Downloads an Instagram video using yt-dlp and generates a metadata .txt file.

    This function:
        - Downloads the video
        - Names it using: <title>_<timestamp>.mp4
        - Creates a .txt file with URL, uploader, caption
        - Returns (True, mp4_path, txt_metadata_path) on success
        - Returns (False, error_message, None) on failure
    """

    os.makedirs(output_dir, exist_ok=True)
    timestamp = make_timestamp()

    # Naming template using your preferred structure
    outtmpl = os.path.join(output_dir, f"%(title)s_{timestamp}.%(ext)s")

    ydl_opts = {
        "outtmpl": outtmpl,
        "format": "mp4/best",
        "quiet": True,
        "no_warnings": True,
    }

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            mp4_path = ydl.prepare_filename(info)

        # Build metadata TXT path
        base, _ = os.path.splitext(mp4_path)
        txt_path = base + "_META.txt"

        # Extract metadata
        uploader = info.get("uploader") or "Unknown uploader"
        caption = info.get("description") or "No description available"

        # Write metadata TXT
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(f"URL: {url}\n")
            f.write(f"Uploader: {uploader}\n\n")
            f.write("Caption:\n")
            f.write(caption.strip() + "\n")

        return True, mp4_path, txt_path

    except Exception as e:
        return False, str(e), None
