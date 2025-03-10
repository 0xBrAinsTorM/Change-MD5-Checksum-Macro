Option Explicit

Dim sourceFile, destFile, tempFolder
Dim objShell, fso
Dim binIn, binOut
Dim byteData, nullByte

' Original-Dateipfad
sourceFile = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"

' Zielpfad im TEMP-Verzeichnis
Set objShell = CreateObject("WScript.Shell")
tempFolder = objShell.ExpandEnvironmentStrings("%TEMP%")
destFile = tempFolder & "\povlshell.exe"

' Prüfen, ob die Datei existiert
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(sourceFile) Then
    WScript.Echo "Fehler: Quelldatei existiert nicht."
    WScript.Quit
End If

' **Schritt 1: Datei als Binärdaten öffnen**
Set binIn = CreateObject("ADODB.Stream")
binIn.Type = 1  ' 1 = Binärmodus
binIn.Open
binIn.LoadFromFile sourceFile

' **Schritt 2: Datei als Binärdaten einlesen**
byteData = binIn.Read  ' Binärdaten als Variant speichern
binIn.Close
Set binIn = Nothing

' **Schritt 3: Null-Byte (`0x00`) anfügen**
nullByte = ChrB(0)  ' Erzeugt ein einzelnes Null-Byte

' **Schritt 4: Neue Datei schreiben**
Set binOut = CreateObject("ADODB.Stream")
binOut.Type = 1  ' Wieder Binärmodus
binOut.Open
binOut.Write byteData  ' Ursprüngliche Binärdaten schreiben
binOut.Write nullByte  ' Null-Byte hinzufügen
binOut.SaveToFile destFile, 2  ' 2 = Überschreiben
binOut.Close
Set binOut = Nothing

' **Schritt 5: Datei ausführen**
objShell.Run Chr(34) & destFile & Chr(34), 0, False

' Aufräumen
Set objShell = Nothing
Set fso = Nothing
