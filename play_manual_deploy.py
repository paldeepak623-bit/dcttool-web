from gtts import gTTS
import os
import subprocess
import ctypes
import time

text = (
    "दीपक भाई, एरर 500 अभी इसलिए आया क्योंकि रेंडर पर अभी भी पुराना वाला कोड ही चल रहा था, "
    "नया फिक्स अभी डिप्लॉय नहीं हुआ था। "
    "आप क्रोम में छठवें नंबर वाले टैब पर क्लिक करें, जहाँ 'dcttool-cloud' का रेंडर डैशबोर्ड खुला है। "
    "वहाँ ऊपर दाहिनी तरफ काले रंग के Manual Deploy बटन पर क्लिक करके "
    "'Clear build cache and deploy' चुन लें। "
    "बस एक से दो मिनट में नया कोड लाइव हो जाएगा। "
    "उसके बाद सातवें टैब पर जाकर पेज को रिफ्रेश करें और प्रोसेस करें, "
    "इस बार 100% बिना किसी एरर के सारे टेम्पलेट्स प्रोसेस होकर डाउनलोड हो जाएँगे।"
)

desktop_path = r"C:\Users\DEEPAK PAL\Desktop\DCT_Audio_Guide.mp3"
tts = gTTS(text=text, lang='hi', slow=False)
tts.save(desktop_path)
print("Saved to Desktop:", desktop_path)

# Open in Windows Media Player
wmplayer = r"C:\Program Files (x86)\Windows Media Player\wmplayer.exe"
if not os.path.exists(wmplayer):
    wmplayer = r"C:\Program Files\Windows Media Player\wmplayer.exe"
if os.path.exists(wmplayer):
    subprocess.Popen([wmplayer, desktop_path])
else:
    subprocess.Popen(["cmd.exe", "/c", "start", "", desktop_path])

# Also play natively through winmm
winmm = ctypes.windll.winmm
buf = ctypes.create_unicode_buffer(256)
winmm.mciSendStringW("close all", None, 0, None)
winmm.mciSendStringW(f'open "{desktop_path}" type mpegvideo alias fix_audio', None, 0, None)
winmm.mciSendStringW("play fix_audio", None, 0, None)

time.sleep(1)
winmm.mciSendStringW("status fix_audio length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 25000

start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status fix_audio mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close fix_audio", None, 0, None)
print("Playback complete!")
