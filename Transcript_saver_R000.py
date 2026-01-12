import os
import re
from pytube import YouTube
from youtube_transcript_api import YouTubeTranscriptApi, TranscriptsDisabled, NoTranscriptFound

def sanitize_filename(name):
    # Remove characters not allowed in filenames
    return re.sub(r'[\\/*?:"<>|]', "", name)

def fetch_and_save_transcript(video_id, preferred_languages=['en'], output_folder='transcripts'):
    try:
        # Get video title using pytube
        yt = YouTube(f"https://www.youtube.com/watch?v={video_id}")
        video_title = sanitize_filename(yt.title)
    except Exception as e:
        print(f"⚠️ Could not fetch video title: {e}")
        video_title = video_id  # fallback to video ID

    try:
        # Try to fetch transcript in preferred languages
        transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=preferred_languages)
    except NoTranscriptFound:
        print("❌ No transcript found in preferred languages. Trying any available transcript...")
        try:
            transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)
            transcript = transcript_list.find_transcript(
                transcript_list._manually_created_transcripts.keys() or
                transcript_list._generated_transcripts.keys()
            ).fetch()
        except Exception as e:
            print(f"❌ Failed to fetch any transcript: {e}")
            return
    except TranscriptsDisabled:
        print("🚫 Transcripts are disabled for this video.")
        return
    except Exception as e:
        print(f"⚠️ Unexpected error: {e}")
        return

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Save to file using video title
    filename = os.path.join(output_folder, f"{video_title}.txt")
    with open(filename, "w", encoding="utf-8") as f:
        for entry in transcript:
            f.write(f"{entry['start']:.2f}s: {entry['text']}\n")

    print(f"✅ Transcript saved to '{filename}'")

# 🔧 Replace with your desired video ID and output folder
video_id = "IgF3OX8nT0w"
fetch_and_save_transcript(video_id, output_folder=r"C:\Users\manga\OneDrive\####Mind_Palace\####Technical\##Veritasium")