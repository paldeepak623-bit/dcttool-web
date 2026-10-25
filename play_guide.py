from gtts import gTTS
import os
import subprocess

text = (
    "दीपक भाई, बहुत बढ़िया खबर है! आपकी स्क्रीन पर जो दो हरे टिक मार्क दिख रहे हैं, "
    "इसका मतलब है कि नया फिक्स रेंडर क्लाउड पर सफलतापूर्वक सिंक हो चुका है। "
    "अब अपनी लाइव वेबसाइट का 24 घंटे चलने वाला लिंक देखने के लिए स्क्रीन पर जो बैंगनी रंग में dcttool-cloud लिखा हुआ दिख रहा है, "
    "उस पर क्लिक करें। या फिर बाईं तरफ मेनू में रिसोर्सेज पर क्लिक करके dcttool-cloud खोलें। "
    "वहाँ सबसे ऊपर ही आपको आपकी वेबसाइट का लाइव यूआरएल मिल जाएगा। "
    "वहाँ नीचे लाइव लॉग्स भी दिखेंगे कि नया कोड डिप्लॉय हो रहा है। "
    "बस एक से दो मिनट में वह लाइव हो जाएगा और आप उस लिंक को पूरे ऑफिस में शेयर कर सकते हैं।"
)

audio_path = r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\render_guide.mp3"
tts = gTTS(text=text, lang='hi', slow=False)
tts.save(audio_path)
print("Saved to", audio_path)

# Play audio in background
subprocess.Popen(["powershell", "-c", f'Start-Process "{audio_path}"'])
