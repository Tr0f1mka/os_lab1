Dim fso, shell, file
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' --- 1. Работа с файлами ---
Dim testFile
testFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\cscript_test.txt"

' Запись файла
Set file = fso.CreateTextFile(testFile, True)
file.WriteLine "Created by cscript!"
file.Close

' Чтение файла
Set file = fso.OpenTextFile(testFile, 1)
WScript.Echo "Readed: " & file.ReadLine
file.Close

' Удаление файла
fso.DeleteFile testFile, True

' --- 2. Запуск дочернего процесса ---
shell.Run "cmd /c echo Son proccess", 0, True

WScript.Echo "End."