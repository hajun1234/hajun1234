Set ws = CreateObject("WScript.Shell") ws.Run "notepad.exe"
WScript.Sleep 300 ws.SendKeys "it is cool! I like it! Do you like it?{ENTER}"
MsgBox "Continue? (Maker/limo not responsible for any damages, you are!)", 1
Dim x
x = 0 ' Initialize x to 0 Do Until x = 600
ws.SendKeys "you NOOB!{ENTER}" createObject("Sapi.SpVoice").speak "you NOOB!"
WScript.Sleep 500 
x = x + 1
Loop
