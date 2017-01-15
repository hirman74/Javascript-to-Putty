class filingName
	public name
end class 

Function resultNameNow ()
	dim m
	set m = new filingName
	If Len (month(now)) < 2 Then
		monthly = "0" & (month(now))
	Else
		monthly = (month(now))
	End If

	If Len (day(now)) < 2 Then
		dayly = "0" & (day(now))
	Else
		dayly = (day(now))
	End If

	If Len (hour(now)) < 2 Then
		hourly = "0" & (hour(now))
	Else
		hourly = (hour(now))
	End If

	If Len (minute(now)) < 2 Then
		minutely = "0" & (minute(now))
	Else
		minutely = (minute(now))
	End If

	If Len (second(now)) < 2 Then
		secondly = "0" & (second(now))
	Else
		secondly = (second(now))
	End If
	m.name = year(now) & monthly & dayly & "_" & hourly & minutely & secondly
	set resultNameNow = m
End Function

    
	addingFileName = WScript.Arguments(0)
	Dim objFSO
	Set m = resultNameNow()
	resultName = m.name
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objRead = objFSO.OpenTextFile (addingFileName, 1, True,-2)
    Set objWrite = objFSO.OpenTextFile (resultName & ".csv", 8, True,-2)
        Do Until objRead.AtEndOfStream 
        strLine = objRead.ReadLine
        arrSpaceGap = Split (strLine, " ")
        p = ""
            For each n in arrSpaceGap
                If len(n) > 0 AND n <> " " Then
                    If p = " " Then
                        p = Trim(n)				
                    Else
                        p = p & "," & Trim(n)
                    End If
                End If
            Next
            If Len(p) > 1 Then
                objWrite.writeline p
            End If
        Loop
    objRead.close
    objWrite.close
    Set objFSO = Nothing
    Set objWrite = Nothing
    Set objRead = Nothing
	
	