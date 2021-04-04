Set wshShell = wscript.CreateObject("WScript.Shell") 

Dim delay
delay = InputBox("Max ping (If your ping is bigger than this number at any time, the bot will crash, 100 recommended)", 1)


wscript.sleep 5000

filePath = WScript.Arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")

Do

Set file = fso.OpenTextFile(filePath)

Do Until file.AtEndOfStream

wscript.sleep delay
wshShell.sendkeys "{ENTER}"
wscript.sleep 100 
wshShell.sendkeys file.ReadLine
wscript.sleep 100
wshShell.sendkeys "{ENTER}"

Loop

file.Close

Loop