from gtts import gTTS
import subprocess

text = (
    "दीपक भाई, समस्या पूरी तरह पकड़ में आ गई है! "
    "पहले लिनक्स क्लाउड पर एक पुरानी लाइब्रेरी मिसिंग थी, जिसकी वजह से वह 47 सेकंड में फेल हो गया था। "
    "मैंने अभी-अभी उसे पूरी तरह ठीक करके गिटहब पर नया अपडेट पुश कर दिया है। "
    "अब आप स्क्रीन पर दाहिनी तरफ काले रंग के बटन 'Manual Deploy' पर क्लिक करके "
    "'Clear build cache & deploy' चुन सकते हैं, या बस एक मिनट रुकें, नया डिप्लॉय अपने आप शुरू होकर वेबसाइट को लाइव कर देगा।"
)

audio_path = r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\fixed_guide.mp3"
tts = gTTS(text=text, lang='hi', slow=False)
tts.save(audio_path)
print("Saved to", audio_path)

# Play via wmplayer
subprocess.Popen([r'C:\Program Files\Windows Media Player\wmplayer.exe', audio_path])
