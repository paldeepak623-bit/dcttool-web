import ctypes
import time

winmm = ctypes.windll.winmm
path = r"C:\Users\DEEPAK PAL\Desktop\DCT_Audio_Guide.mp3"
buf = ctypes.create_unicode_buffer(256)

winmm.mciSendStringW("close all", None, 0, None)
winmm.mciSendStringW(f'open "{path}" type mpegvideo alias guide_snd', None, 0, None)
winmm.mciSendStringW("play guide_snd", None, 0, None)

time.sleep(1)
winmm.mciSendStringW("status guide_snd length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 20000

start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status guide_snd mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close guide_snd", None, 0, None)
print("Playback completed!")
