from gtts import gTTS
import subprocess

text = (
    "दीपक भाई, बिल्कुल भी घबराने की ज़रूरत नहीं है! "
    "क्लाउड पर सब कुछ पूरी तरह सुरक्षित है और अपने आप चल रहा है। "
    "मैंने आपके कंप्यूटर पर सीधे रेंडर का डैशबोर्ड खोल दिया है। "
    "जैसे ही डैशबोर्ड खुले, आपको सामने dcttool-cloud नाम की सर्विस दिखेगी। "
    "बस उस पर क्लिक करें, वहाँ सबसे ऊपर आपकी वेबसाइट का लाइव लिंक मिल जाएगा। "
    "आप उस लिंक को कॉपी करके अपने पूरे ऑफिस में शेयर कर सकते हैं।"
)

audio_path = r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\reopen_guide.mp3"
tts = gTTS(text=text, lang='hi', slow=False)
tts.save(audio_path)
print("Saved to", audio_path)

# Play audio via play_wpf
ps_code = f'''
Add-Type -AssemblyName presentationCore
$player = New-Object system.windows.media.mediaplayer
$player.open([System.Uri]"{audio_path}")
$player.Play()
Start-Sleep -Seconds 20
'''
with open(r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\play_reopen.ps1", "w", encoding="utf-8") as f:
    f.write(ps_code)

# Open browser to Render dashboard
subprocess.Popen(["cmd.exe", "/c", "start", "https://dashboard.render.com"])
# Play audio
subprocess.Popen(["powershell", "-ExecutionPolicy", "Bypass", "-File", r"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\play_reopen.ps1"])
