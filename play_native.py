import ctypes
import time
import os

path = r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\fixed_guide.mp3"
if not os.path.exists(path):
    print("File does not exist:", path)
    exit(1)

winmm = ctypes.windll.winmm
buf = ctypes.create_unicode_buffer(256)

# Close any previous instance
winmm.mciSendStringW("close all", None, 0, None)

# Open audio
r1 = winmm.mciSendStringW(f'open "{path}" type mpegvideo alias my_audio', None, 0, None)
print("Open result:", r1)

# Play audio
r2 = winmm.mciSendStringW("play my_audio", None, 0, None)
print("Play result:", r2)

time.sleep(1)
winmm.mciSendStringW("status my_audio mode", buf, 256, None)
print("Status mode:", buf.value)

# Query length in ms
winmm.mciSendStringW("status my_audio length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 20000
print("Length (ms):", length_ms)

# Wait until playback finishes
start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status my_audio mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close my_audio", None, 0, None)
print("Playback complete!")
