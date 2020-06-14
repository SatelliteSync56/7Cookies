Dim oPlayer
Public stpswrd
Set WshShell = WScript.CreateObject("WScript.Shell")
a = MsgBox("7Cookies by Trangle APG 2019, support development at trangleproject.wordpress.com",1,"7Cookies 1.6")
If a = 2 Then
Wscript.Quit
End If
a = MsgBox("You do not have a 7Cookies Account. Create one now?",3,"7Cookies 1.6")
If a = 2 Then
Wscript.Quit
End If
If a = 7 Then
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\WINDOWS\Media\Windows Critical Stop.wav"
oPlayer.controls.play
a = MsgBox("Error. Could not connect to server due to network error. Code:0500",5,"Connecting...")
If a = 2 Then
Wscript.Quit
End If
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend
oPlayer.close
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\WINDOWS\Media\Windows Critical Stop.wav"
oPlayer.controls.play
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend
oPlayer.close
a = MsgBox("Error. Could not connect to server due to network error. Code:0500",0,"Connecting...")
Wscript.Quit
End If
stpswrd = InputBox("Password:","Create Account")
uspswrd = InputBox("Please enter your 7Cookies password below.","Login")
If uspswrd = stpswrd Then
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\WINDOWS\Media\Windows Critical Stop.wav"
oPlayer.controls.play
a = MsgBox("Error. Could not connect to server due to network error. Code:0500",5,"Connecting...")
If a = 2 Then
Wscript.Quit
End If
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend
oPlayer.close
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\WINDOWS\Media\Windows Critical Stop.wav"
oPlayer.controls.play
a = MsgBox("Error. Could not connect to server due to network error. Code:0500",0,"Connecting...")
Wscript.Quit
End If
uspswrd = InputBox("Please enter your 7Cookies password below.","Login")
If uspswrd = stpswrd Then
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\WINDOWS\Media\Windows Critical Stop.wav"
oPlayer.controls.play
a = MsgBox("Error. Could not connect to server due to network error. Code:0500",5,"Connecting...")
If a = 2 Then
Wscript.Quit
End If
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend
oPlayer.close
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\WINDOWS\Media\Windows Critical Stop.wav"
oPlayer.controls.play
a = MsgBox("Error. Could not connect to server due to network error. Code:0500",0,"Connecting...")
Wscript.Quit
End If
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend
oPlayer.close
uspswrd = MsgBox("The password is incorrect",0,"Login")
