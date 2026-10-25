$wmp = New-Object -ComObject WMPlayer.OCX
$wmp.URL = "C:\Users\DEEPAK PAL\Desktop\DCT-Cloud-WebApp\render_guide.mp3"
$wmp.controls.play()
Start-Sleep -Seconds 1
while ($wmp.playState -eq 3 -or $wmp.playState -eq 9 -or $wmp.playState -eq 6) {
    Start-Sleep -Milliseconds 300
}
