# -*- coding: utf-8 -*-
import os
import subprocess
import time
from pathlib import Path

# Script to synthesize clear Hindi/Hinglish audio explanation and play it
text = (
    "नमस्ते दीपक भाई! आपकी दोनों समस्याओं को पूरी तरह जड़ से ठीक कर दिया गया है। "
    "पहला: रेंडर क्लाउड पर मास्टर शीट नॉट फाउंड या एक्सपायर्ड की जो एरर आ रही थी, "
    "उसमें ऑटो-रिकवरी जोड़ दी गई है। अब अगर रेंडर रीस्टार्ट भी होता है, तो टूल बिना कोई एरर दिखाए "
    "ब्राउज़र मेमोरी से फ़ाइल को तुरंत रिस्टोर करके ऑटोमेशन पूरा कर देगा। "
    "दूसरा: विलेज वाला नियम! मेंबरशिप टेम्पलेट में अब सिर्फ और सिर्फ पोर्टल ड्रॉपडाउन के रजिस्टर्ड गांव ही लिए जाएंगे। "
    "जैसे मुदरासन पैक्स के पोर्टल पर एक भी गांव रजिस्टर नहीं था, केवल अदर्स था, तो टूल ने नियमों के अनुसार प्रोसेस तुरंत रोक दिया "
    "और साफ एरर डायलॉग दिखा दिया कि बिना पोर्टल गांव के एक्सेल के पुराने गांव इस्तेमाल नहीं हो सकते। "
    "और भूलनपुर जैसी पैक्स जहां पोर्टल पर 10 गांव रजिस्टर्ड हैं, वहां मात्र 5 सेकंड में सभी टेम्पलेट्स असली गांवों के साथ तैयार हो रहे हैं। "
    "सभी फाइल्स गिटहब पर पुश हो चुकी हैं, रेंडर पर नया डिप्लॉय ट्रिगर हो गया है, "
    "और साथ ही आपका लाइव क्लाउड टनल लिंक भी एकदम तैयार है। आप बेफिक्र होकर टेस्ट कर सकते हैं!"
)

ps_script = f'''
Add-Type -AssemblyName System.Speech
$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synth.Rate = 0
$synth.Volume = 100

$voice = $synth.GetVoices() | Where-Object {{ $_.VoiceInfo.Culture.Name -like "*IN*" -or $_.VoiceInfo.Name -like "*Kalpana*" -or $_.VoiceInfo.Name -like "*Heera*" -or $_.VoiceInfo.Name -like "*Ravi*" }} | Select-Object -First 1
if ($voice) {{
    $synth.SelectVoice($voice.VoiceInfo.Name)
}}

$outPath1 = "C:\\Users\\DEEPAK PAL\\Desktop\\DCT_Issue_Resolved_Guide.mp3"
$outPath2 = "C:\\Users\\DEEPAK PAL\\OneDrive\\Desktop\\DCT_Issue_Resolved_Guide.mp3"
$synth.SetOutputToWaveFile("C:\\Users\\DEEPAK PAL\\Desktop\\DCT_temp.wav")
$synth.Speak(@"
{text}
"@)
$synth.Dispose()

# Convert or copy wav
if (Test-Path "C:\\Users\\DEEPAK PAL\\Desktop\\DCT_temp.wav") {{
    Copy-Item "C:\\Users\\DEEPAK PAL\\Desktop\\DCT_temp.wav" "C:\\Users\\DEEPAK PAL\\Desktop\\DCT_Issue_Resolved_Guide.wav" -Force
    if (Test-Path "C:\\Users\\DEEPAK PAL\\OneDrive\\Desktop") {{
        Copy-Item "C:\\Users\\DEEPAK PAL\\Desktop\\DCT_temp.wav" "C:\\Users\\DEEPAK PAL\\OneDrive\\Desktop\\DCT_Issue_Resolved_Guide.wav" -Force
    }}
}}
'''

with open("generate_audio.ps1", "w", encoding="utf-8") as f:
    f.write(ps_script)

subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", "generate_audio.ps1"], check=True)

# Play audio via Windows default media player / sound player
play_ps = '''
$player = New-Object System.Media.SoundPlayer "C:\\Users\\DEEPAK PAL\\Desktop\\DCT_Issue_Resolved_Guide.wav"
$player.PlaySync()
'''
with open("play_sound.ps1", "w", encoding="utf-8") as f:
    f.write(play_ps)

subprocess.Popen(["powershell", "-ExecutionPolicy", "Bypass", "-File", "play_sound.ps1"])
print("Audio guide created and started playing!")
