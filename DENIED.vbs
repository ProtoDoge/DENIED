x=msgbox("DENIED" ,48+4096, " ")

' MUST BE AN .MP3

Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

oPlayer.URL = "[PATH TO MP3]"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 10
Wend

oPlayer.close

' MUST BE A .PNG

CreateObject("WScript.Shell").Run "[FILE PATH TO PNG]"

' MUST BE A .BAT

Set oShell = CreateObject("Wscript.Shell")
oShell.CurrentDirectory = "c:\"
oShell.Run ("[FILEPATH TO BATCH") 
SET oShell=nothing


