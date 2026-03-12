import os
import pandas as pd
import time
from datetime import datetime
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import subprocess
import psutil
import sys


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
def set_mute(mute=True):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
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


def main():
    # Path to the CSV file
    csv_file = 'test.csv'

    # Read the CSV file
    try:
        df = pd.read_csv(csv_file)
    except FileNotFoundError:
        print(f"Error: The file {csv_file} was not found.")
        return
    except pd.errors.EmptyDataError:
        print(f"Error: The file {csv_file} is empty.")
        return
    except pd.errors.ParserError:
        print(f"Error: There was a problem parsing the file {csv_file}.")
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

        # Mute all audio devices
        set_mute(mute=True)
        print("Audio sessions after muting:")
        print(list_audio_sessions())

        # Play the audio file and wait for it to finish
        print(f"Playing audio: {audio_file}")
        audio_process = play_audio(audio_file)
        audio_process.wait()  # blocks until 'start /wait' returns (file done playing)

        # Unmute all audio devices
        set_mute(mute=False)
        print("Audio playback finished and system audio unmuted")
        print("Audio sessions after unmuting:")
        print(list_audio_sessions())

if __name__ == "__main__":
    main()