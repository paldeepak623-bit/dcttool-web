from gtts import gTTS
import ctypes
import time
import os

text = (
    "दीपक भाई, ऑडियो मैंने सीधे आपके स्पीकर्स पर प्ले कर दिया है। "
    "अब बस अपनी स्क्रीन पर ऊपर काले रंग के 'Manual Deploy' बटन पर क्लिक करें, "
    "और 'Clear build cache and deploy' चुन लें। "
    "नया फिक्स तुरंत डिप्लॉय हो जाएगा और आपकी वेबसाइट लाइव हो जाएगी।"
)

audio_path = r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\deploy_now.mp3"
tts = gTTS(text=text, lang='hi', slow=False)
tts.save(audio_path)

winmm = ctypes.windll.winmm
buf = ctypes.create_unicode_buffer(256)
winmm.mciSendStringW("close all", None, 0, None)
winmm.mciSendStringW(f'open "{audio_path}" type mpegvideo alias my_audio', None, 0, None)
winmm.mciSendStringW("play my_audio", None, 0, None)

time.sleep(1)
winmm.mciSendStringW("status my_audio length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 15000

start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status my_audio mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close my_audio", None, 0, None)
print("Deploy now audio completed!")
