import os
import sys
import pandas as pd
import time
from datetime import datetime
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import subprocess
import psutil

LOCK_FILE = 'script.lock'

# Function to create a lock file
def create_lock():
    if os.path.exists(LOCK_FILE):
        print("Another instance of the script is already running.")
        sys.exit()
    with open(LOCK_FILE, 'w') as f:
        f.write("This file is used to lock the script execution.")

# Function to remove the lock file
def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

# Function to list all audio sessions
def list_audio_sessions():
    sessions = AudioUtilities.GetAllSessions()
    session_info = []
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        session_info.append({
            "process_id": session.ProcessId,
            "process_name": session.Process.name() if session.Process else "Unknown",
            "mute": volume.GetMute(),
        })
    return session_info

# Function to mute or unmute audio devices
def set_mute(except_process_id=None, mute=True):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.ProcessId == except_process_id:
            continue
        volume.SetMute(mute, None)

# Function to play the audio file using the default media player
def play_audio(file_path):
    process = subprocess.Popen(['start', '/wait', '', file_path], shell=True)
    return process


def interruptible_wait(seconds):
    """Sleep in small intervals so the script stays responsive."""
    end_time = time.time() + seconds
    while time.time() < end_time:
        time.sleep(min(1, end_time - time.time()))

# Function to find the media player process
def find_media_player_process():
    media_player_names = ["wmplayer.exe", "vlc.exe", "mpc-hc.exe", "mpc-be.exe", "foobar2000.exe", "potplayer.exe", "aimp.exe", "microsoft.media.player.exe"]
    for process in psutil.process_iter(['name']):
        if process.info['name'].lower() in media_player_names:
            return process
    return None

def main():
    # Ensure only one instance of the script is running
    create_lock()

    try:
        # Path to the Excel schedule file (created by xlsxBuilder.py)
        xlsx_file = 'path/to/your/schedule.xlsx'

        # Read the Excel file
        try:
            df = pd.read_excel(xlsx_file, dtype=str)
        except FileNotFoundError:
            print(f"Error: The file {xlsx_file} was not found.")
            return
        except Exception as e:
            print(f"Error reading {xlsx_file}: {e}")
            return

        for index, row in df.iterrows():
            try:
                # Parse the date and time
                scheduled_time = datetime.strptime(row['datetime'], '%m/%d/%Y %H:%M')
                audio_file = row['file_path']
            except KeyError as e:
                print(f"Error: Missing column in CSV file: {e}")
                return
            except ValueError as e:
                print(f"Error: {e}")
                return

            # Calculate the time difference
            time_diff = (scheduled_time - datetime.now()).total_seconds()

            if time_diff <= 0:
                print(f"Skipping past-scheduled entry at {scheduled_time} (row {index})")
                continue

            print(f"Waiting {time_diff:.1f} seconds until {scheduled_time}")
            interruptible_wait(time_diff)

            # Play the audio file
            print(f"Playing audio: {audio_file}")
            audio_process = play_audio(audio_file)

            # Wait briefly to allow the media player to start before finding its process
            time.sleep(2)

            # Find the media player process and mute everything else
            media_player_process = find_media_player_process()
            if media_player_process is None:
                print("Warning: Media player process not found; muting all sessions.")
                set_mute(mute=True)
            else:
                print(f"Media player process ID: {media_player_process.pid}")
                set_mute(except_process_id=media_player_process.pid, mute=True)
            print("Audio sessions after muting:")
            print(list_audio_sessions())

            # Wait for audio to finish (start /wait holds until player closes)
            audio_process.wait()

            # Unmute all audio devices
            set_mute(mute=False)
            print("Audio playback finished and system audio unmuted")
            print("Audio sessions after unmuting:")
            print(list_audio_sessions())
    finally:
        # Remove the lock file
        remove_lock()

if __name__ == "__main__":
    main()
