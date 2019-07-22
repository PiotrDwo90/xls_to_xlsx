
Dim app, fso, file, fName, wb, diropen, dirsave, strLogLocation 

diropen = "c:\prowizje\xls\"
dirsave = "c:\prowizje\"
strLogLocation = "C:\prowizje\bin\"

Set objShell = WScript.CreateObject("WScript.Shell")
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
strScript = Wscript.ScriptName

strLogName = Left(strScript, Len(strScript)-4) & "HistorycznyLog.txt"
strLogNameTymczas = Left(strScript, Len(strScript)-4) & "OstatniLog.txt"

Set objLogFile = objFSO.OpenTextFile(strLogLocation & strLogName, 8, True)
Set objLogFileTymczas = objFSO.CreateTextFile(strLogLocation & strLogNameTymczas, 8, True)

Set app = CreateObject("Excel.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
	
	objLogFile.WriteLine("PROCES START: "& Now & vbCrLf)
	objLogFileTymczas.WriteLine("PROCES START: "& Now & vbCrLf)
	

	For Each file In fso.GetFolder(diropen).Files	
		   		
	    If LCase(fso.GetExtensionName(file)) = "xls" Then  
		objLogFile.WriteLine( "PROCES dla pliku: "&file&" o rozszerzeniu XLS Rozpoczêty !")
		objLogFileTymczas.WriteLine( "PROCES dla pliku: "&file&" o rozszerzeniu XLS Rozpoczêty !")
	    fName = fso.GetBaseName(file)
	    Set wb = app.Workbooks.Open(file) 
	    app.Application.Visible = True
	    app.Application.DisplayAlerts = False
	    app.ActiveWorkbook.SaveAs dirsave & fName & ".xlsx", 51
	
	    app.ActiveWorkbook.Close
	    app.Application.DisplayAlerts = True
	    app.Application.Quit
			If Err.Number <> 0 Then
			objLogFile.WriteLine(vbCrLf &Now & vbTab & strMsg)
				objLogFile.WriteLine( "   ERROR       : " & strScript & " file: "&file)
				objLogFile.WriteLine( "   ERROR       : " & Err.Number)
				objLogFile.WriteLine( "   ERROR (Hex) : " & Hex(Err.Number))
    				objLogFile.WriteLine( "   SOURCE      : " & Err.Source)
   				objLogFile.WriteLine( "   DESCRIPTION : " & Err.Description)
			objLogFileTymczas.WriteLine(vbCrLf &Now & vbTab & strMsg)
				objLogFileTymczas.WriteLine( "   ERROR       : " & strScript & " file: "&file)
				objLogFileTymczas.WriteLine( "   ERROR       : " & Err.Number)
				objLogFileTymczas.WriteLine( "   ERROR (Hex) : " & Hex(Err.Number))
    				objLogFileTymczas.WriteLine( "   SOURCE      : " & Err.Source)
   				objLogFileTymczas.WriteLine( "   DESCRIPTION : " & Err.Description)
			Else
				objLogFile.WriteLine("Success: " & strScript & " Copy file: "&file)
				objLogFileTymczas.WriteLine("Success: " & strScript & " Copy file: "&file)
			End If
		Err.Clear

	    Else

objLogFile.WriteLine( "   ERROR       : " & strScript & " file: "&file)
objLogFileTymczas.WriteLine( "   ERROR       : " & strScript & " file: "&file)


		objLogFile.WriteLine("   ERROR       : "& file &" To nie plik *.XLS")
		objLogFileTymczas.WriteLine("   ERROR       : "& file &" To nie plik *.XLS")

	    End If
	Next


	'Kasuje pliki *.tmp'
	For Each file In fso.GetFolder(dirsave).Files
		If LCase(fso.GetExtensionName(file)) = "tmp" Then  
			objLogFile.WriteLine(vbCrLf &"KASOWANIE pliku: "&file)
			objLogFileTymczas.WriteLine(vbCrLf & "KASOWANIE pliku: "&file)
			ObjFSO.DeleteFile dirsave&"*.tmp", True	
			If Err.Number <> 0 Then
				objLogFile.WriteLine(Now & vbTab & strMsg)
					objLogFile.WriteLine( "   ERROR       KASOWANIE: " & strScript & " file: "&file)
					objLogFile.WriteLine( "   ERROR       : " & Err.Number)
					objLogFile.WriteLine( "   ERROR (Hex) : " & Hex(Err.Number))
    					objLogFile.WriteLine( "   SOURCE      : " & Err.Source)
   					objLogFile.WriteLine( "   DESCRIPTION : " & Err.Description)
				objLogFileTymczas.WriteLine( Now & vbTab & strMsg)
					objLogFileTymczas.WriteLine( "   ERROR       : " & strScript & " file: "&file)
					objLogFileTymczas.WriteLine( "   ERROR       : " & Err.Number)
					objLogFileTymczas.WriteLine( "   ERROR (Hex) : " & Hex(Err.Number))
    					objLogFileTymczas.WriteLine( "   SOURCE      : " & Err.Source)
   					objLogFileTymczas.WriteLine( "   DESCRIPTION : " & Err.Description)
			End If
		End If
	Next

objShell.Run "taskkill /im EXCEL.EXE", , False
objLogFile.WriteLine(vbCrLf &"PROCES STOP: "& Now & vbCrLf & vbCrLf)
objLogFileTymczas.WriteLine(vbCrLf &"PROCES STOP: "& Now & vbCrLf & vbCrLf)
Set fso = Nothing
Set wb = Nothing    
Set app = Nothing

wScript.Quit

