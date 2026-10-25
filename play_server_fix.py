from gtts import gTTS
import os
import subprocess
import ctypes
import time

text = (
    "दीपक भाई, बहुत-बहुत बधाई! आपका क्लाउड टूल अब रेंडर पर पूरी तरह लाइव हो चुका है "
    "और दूसरे कंप्यूटर पर भी लिंक सफलतापूर्वक खुल गया है! "
    "जो एरर 500 आ रहा था, वह इसलिए था क्योंकि बैकएंड सर्वर में पोर्टल सेटिंग्स का एक छोटा सा पैरामीटर मिसमैच था। "
    "मैंने अभी-अभी उसे पूरी तरह ठीक करके गिटहब पर नया अपडेट पुश कर दिया है। "
    "बस एक से दो मिनट में रेंडर इसे लाइव कर देगा। "
    "उसके बाद दूसरे कंप्यूटर पर पेज को रीलोड करके जब आप प्रोसेस करेंगे, "
    "तो बिना किसी एरर के सारे टेम्पलेट्स प्रोसेस होकर ज़िप फाइल अपने आप डाउनलोड हो जाएगी!"
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
winmm.mciSendStringW(f'open "{desktop_path}" type mpegvideo alias final_audio', None, 0, None)
winmm.mciSendStringW("play final_audio", None, 0, None)

time.sleep(1)
winmm.mciSendStringW("status final_audio length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 25000

start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status final_audio mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close final_audio", None, 0, None)
print("Playback complete!")
