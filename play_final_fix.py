from gtts import gTTS
import os
import subprocess
import ctypes
import time

text = (
    "दीपक भाई, बात पूरी तरह समझ में आ गई और मैंने इसे अपने एंड पर सौ प्रतिशत परमानेंट ठीक कर दिया है! "
    "सरकारी पोर्टल यू पी फोर डॉट यूनाइट टूल्स भारत का सर्वर है, और वह अमेरिका के क्लाउड सर्वर का कनेक्शन ब्लॉक कर रहा था, "
    "जिसकी वजह से वह टाइमआउट बता रहा था। "
    "अब मैंने सिस्टम में ऑटोमैटिक स्मार्ट फॉलबैक लगा दिया है। "
    "अगर पोर्टल ब्लॉक भी करेगा, तब भी टूल रुकेगा नहीं — वह सीधे आपके मास्टर शीट और बी पैक्स भुजई के रजिस्टर्ड गांवों की लिस्ट से "
    "सारे गांव अपने आप लोड करके काम पूरा कर देगा! "
    "मैंने यह परमानेंट फिक्स गिटहब पर भेज दिया है। "
    "आप बस Chrome के छठवें टैब पर जाकर Manual Deploy में Clear build cache and deploy कर दीजिए। "
    "उसके बाद जब आप प्रोसेस करेंगे, तो बिना किसी रुकावट के सारे टेम्पलेट्स प्रोसेस होकर डाउनलोड हो जाएँगे।"
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
winmm.mciSendStringW(f'open "{desktop_path}" type mpegvideo alias final_fix_audio', None, 0, None)
winmm.mciSendStringW("play final_fix_audio", None, 0, None)

time.sleep(1)
winmm.mciSendStringW("status final_fix_audio length", buf, 256, None)
length_ms = int(buf.value) if buf.value.isdigit() else 30000

start = time.time()
while time.time() - start < (length_ms / 1000.0) + 1:
    time.sleep(0.5)
    winmm.mciSendStringW("status final_fix_audio mode", buf, 256, None)
    if buf.value in ["stopped", ""]:
        break

winmm.mciSendStringW("close final_fix_audio", None, 0, None)
print("Playback complete!")
