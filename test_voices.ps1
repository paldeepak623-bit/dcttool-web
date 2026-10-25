
Add-Type -AssemblyName System.Speech
 = New-Object System.Speech.Synthesis.SpeechSynthesizer
foreach ( in .GetInstalledVoices()) {
    Write-Host .VoiceInfo.Name (.VoiceInfo.Culture.Name)
}
