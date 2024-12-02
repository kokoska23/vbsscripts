Set WshShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("WScript.Shell")

WshShell.Run "notepad"
WScript.Sleep 1000

message = "I'm watching you..."
For i = 1 To Len(message)
    WshShell.SendKeys Mid(message, i, 1)
    WScript.Sleep 150
Next

WScript.Sleep 1000

url = "https://youtu.be/dQw4w9WgXcQ?si=Ec71cR51e-QwreS5"
objShell.Run url
