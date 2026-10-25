
Add-Type -AssemblyName presentationCore
$player = New-Object system.windows.media.mediaplayer
$player.open([System.Uri]"C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\reopen_guide.mp3")
$player.Play()
Start-Sleep -Seconds 20
