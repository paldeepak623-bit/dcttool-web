import time
import win32com.client
from pathlib import Path

audio_file = r"C:\Users\DEEPAK PAL\Desktop\Domain_Cost_Truth_Audio.mp3"
if not Path(audio_file).exists():
    audio_file = r"C:\Users\DEEPAK PAL\OneDrive\Desktop\Domain_Cost_Truth_Audio.mp3"

wmp = win32com.client.Dispatch("WMPlayer.OCX")
wmp.settings.volume = 100
wmp.URL = audio_file

# Wait for playback to initialize and play to completion
print("Playing audio directly through system speakers...")
start = time.time()
time.sleep(2)

while True:
    state = wmp.playState
    # 3 = Playing, 2 = Paused, 1 = Stopped, 8 = MediaEnded
    pos = wmp.controls.currentPosition
    dur = wmp.currentMedia.duration if wmp.currentMedia else 0
    if dur > 0 and pos >= dur - 0.5:
        break
    if state == 1 and time.time() - start > 10:
        break
    time.sleep(1)

print("Audio playback completed successfully!")
