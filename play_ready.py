from gtts import gTTS
import os
import subprocess
import ctypes
import time

text = (
    "दीपक भाई, बहुत ही शानदार खबर! "
    "आपकी स्क्रीन के लॉग्स में साफ दिख रहा है कि नया कोड v2.4 रेंडर पर पूरी तरह लाइव हो चुका है "
    "और आपकी एक्सेल फाइल भी सफलतापूर्वक अपलोड हो चुकी है। "
    "अब आपको बस Chrome में सातवें नंबर वाले टैब यानी DCTtool वाले पेज पर जाना है, "
    "और नीचे बड़े बैंगनी रंग के Processing templates बटन पर क्लिक करना है। "
    "इस बार बिना किसी एरर के आपकी पूरी फाइल प्रोसेस हो जाएगी "
    "और सारे टेम्पलेट्स की ज़िप फाइल अपने आप डाउनलोड हो जाएगी!"
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
winmm.mciSendStringW(f'open "{desktop_path}" type mpegvideo alias ready_audio', None, 0, None)
winmm.mciSendStringW("play ready_audio", None, 0, None)

time.sleep(1)
winmm.mciSendStringW("status ready_audio length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 25000

start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status ready_audio mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close ready_audio", None, 0, None)
print("Playback complete!")
