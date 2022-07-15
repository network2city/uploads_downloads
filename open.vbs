Dim objShell
Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run("C:\Users\Patri\Downloads\Doc5.docx")
wscript.sleep (100)
objShell.Run("C:\Users\Patri\OneDrive\Desktop\MBR_Virii\build\py_file\file2.txt")
wscript.sleep (100)
wscript.quit