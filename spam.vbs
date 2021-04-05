Set wshShell = wscript.CreateObject("WScript.Shell") 

IF InputBox("This bot works based on simulated key presses. Once it starts, you cannot stop it until it finishes, unless you somehow manage to kill wscript.exe in task manager with your keyboard spamming.") THEN

Dim delay
delay = InputBox("Max ping (If your ping is bigger than this number at any time, the bot will crash, 100 recommended)", "Max ping?", 1)

Dim filePath

IF Wscript.Arguments.Count THEN
filePath = WScript.Arguments(0)

ELSE
filePath = InputBox("Spam file path?", "File path?", "bee-movie-script.txt")
END IF

Dim repeat
repeat = CInt(InputBox("How many times to repeat after end of file? (-1 for infinite)", "Repeat EOF?", True))

wscript.sleep 5000

Set fso = CreateObject("Scripting.FileSystemObject")

WHILE repeat

Set file = fso.OpenTextFile(filePath)

Do Until file.AtEndOfStream

wscript.sleep delay
wshShell.sendkeys "{ENTER}"
wscript.sleep 150
wshShell.sendkeys file.ReadLine
wscript.sleep 150
wshShell.sendkeys "{ENTER}"

Loop

file.Close

repeat = repeat - 1

WEND

END IF