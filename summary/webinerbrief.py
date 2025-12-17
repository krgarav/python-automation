import os
import requests
import tempfile
import shutil
import subprocess
import openai
from fastapi import Request
from dotenv import load_dotenv

load_dotenv()

MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")
MONDAY_API_URL = "https://api.monday.com/v2"

openai.api_key = os.getenv("OPENAI_API_KEY")

# Track processing jobs to avoid duplicates
PROCESSING_JOBS = set()
PROCESSED_ITEMS = set()


# --------------------------------------------------------
# Helper: Monday API Caller
# --------------------------------------------------------
def monday_api(query, variables):
    headers = {
        "Authorization": MONDAY_API_KEY,
        "Content-Type": "application/json"
    }
    response = requests.post(
        MONDAY_API_URL,
        json={"query": query, "variables": variables},
        headers=headers
    )
    return response.json()


# --------------------------------------------------------
# Helper: Add Comment to Monday Item
# --------------------------------------------------------
def add_comment(item_id, text):
    mutation = """
    mutation ($item_id: ID!, $text: String!) {
      create_update (item_id: $item_id, body: $text) { 
        id 
        body 
      }
    }
    """
    monday_api(mutation, {"item_id": str(item_id), "text": text})
    print(f"📝 Comment added to item {item_id}")


# --------------------------------------------------------
# Google Drive Download
# --------------------------------------------------------
def download_google_drive_file(file_id, destination_path):
    session = requests.Session()
    URL = "https://drive.google.com/uc?export=download"

    print("⬇ Downloading from Google Drive...")
    response = session.get(URL, params={"id": file_id}, stream=True)

    def get_confirm_token(response):
        for key, value in response.cookies.items():
            if key.startswith("download_warning"):
                return value
        return None

    token = get_confirm_token(response)
    if token:
        response = session.get(URL, params={"id": file_id, "confirm": token}, stream=True)

    if "text/html" in response.headers.get("Content-Type", ""):
        print("❌ Google Drive returned HTML — likely access issue.")
        return False

    os.makedirs(os.path.dirname(destination_path), exist_ok=True)

    with open(destination_path, "wb") as f:
        for chunk in response.iter_content(32768):
            if chunk:
                f.write(chunk)

    print("✅ Downloaded:", destination_path)
    return True


# --------------------------------------------------------
# AUDIO / TRANSCRIPTION PIPELINE
# --------------------------------------------------------
def extract_audio(video_path, audio_path):
    print("🎧 Extracting audio...")

    command = [
        "ffmpeg", "-y",
        "-i", video_path,
        "-vn",
        "-acodec", "libmp3lame",
        "-b:a", "128k",
        audio_path
    ]

    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        raise Exception("FFmpeg failed")

    print("✅ Audio extracted:", audio_path)


def split_audio_if_needed(audio_path, chunk_dir, max_size=20 * 1024 * 1024):
    size = os.path.getsize(audio_path)
    print("📦 Audio size:", size)

    if size <= max_size:
        return [audio_path]

    print("⚠ Splitting audio...")

    chunk_template = os.path.join(chunk_dir, "chunk_%03d.mp3")
    command = [
        "ffmpeg", "-i", audio_path,
        "-f", "segment",
        "-segment_time", "300",
        "-c", "copy",
        chunk_template
    ]
    subprocess.run(command)

    chunks = [
        os.path.join(chunk_dir, f)
        for f in os.listdir(chunk_dir)
        if f.endswith(".mp3")
    ]

    print("✅ Created chunks:", chunks)
    return chunks


def transcribe_chunk(path):
    print("📝 Transcribing:", path)

    with open(path, "rb") as f:
        transcript = openai.Audio.transcribe(
            model="whisper-1",
            file=f
        )
    return transcript["text"]


def summarize_text(text):
    print("🧠 Summarizing...")

    response = openai.ChatCompletion.create(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": "Summarize transcripts clearly and professionally."},
            {"role": "user", "content": "Summarize this transcript:\n\n" + text}
        ]
    )
    return response["choices"][0]["message"]["content"]


# --------------------------------------------------------
# PROCESS WEBHOOK (called by main.py)
# --------------------------------------------------------
async def process_webinerbrief(data):
    print("🔔 Webhook received:", data)

    try:
        item_id = data["event"]["pulseId"]
        video_url = data["event"]["value"]["url"]
        file_id = video_url.split("/d/")[1].split("/")[0]

        if file_id in PROCESSED_ITEMS:
            print("⏭️ Already processed — skipping:", file_id)
            return {"status": "already_done"}

        if file_id in PROCESSING_JOBS:
            print("🚫 Skipping duplicate processing attempt:", file_id)
            return {"status": "skipped"}

        PROCESSING_JOBS.add(file_id)
        print("🔒 Locked processing for:", file_id)

        video_path = f"videos/{file_id}.mp4"
        ok = download_google_drive_file(file_id, video_path)

        if not ok:
            add_comment(item_id, "❌ Could not download video. Check Drive permissions.")
            return {"error": "download_failed"}

        temp_dir = tempfile.mkdtemp()
        try:
            chunk_dir = os.path.join(temp_dir, "chunks")
            os.makedirs(chunk_dir, exist_ok=True)

            audio_path = os.path.join(temp_dir, "audio.mp3")

            extract_audio(video_path, audio_path)
            chunks = split_audio_if_needed(audio_path, chunk_dir)

            transcript = ""
            for c in chunks:
                transcript += transcribe_chunk(c) + "\n"

            summary = summarize_text(transcript)

            add_comment(item_id, summary)

            print("🎉 DONE — Summary added to Monday!")

            PROCESSED_ITEMS.add(file_id)

            return {"status": "success", "summary": summary}

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            print("🧽 Temp files removed.")

            if os.path.exists(video_path):
                os.remove(video_path)
                print("🗑️ Deleted local video file:", video_path)

    except Exception as e:
        print("❌ ERROR:", e)
        return {"error": str(e)}

    finally:
        if "file_id" in locals() and file_id in PROCESSING_JOBS:
            PROCESSING_JOBS.remove(file_id)
            print("🔓 Unlocked job:", file_id)
