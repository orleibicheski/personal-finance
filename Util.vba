REM  *****  BASIC  *****

Sub Util

End Sub

Function IsBlank(text_ As String) As Boolean

  If text_ = "" Then
  	IsBlank = True
  Else
  	IsBlank = False
  End If

End Function

Function Contains(target As String, search As String) As Boolean

	If InStr(UCase(target), UCase(search)) > 0 Then
		Contains = False
	Else
		Contains = True
	End If

End Function

Function YearMonth(date_ As String)

   	Dim year, month
   	Dim spllited
   	spllited = split(date_, "/")
   	
   	year = spllited(2)
   	month = spllited(0)
   	
   	If Len(month) < 2 Then
   		month = "0" & month
   	End If
   	
   	YearMonth = year & "-" & month

End Function

Function YearMonthDay(date_ As String)

   	Dim year, month, day
   	Dim spllited
   	spllited = split(date_, "/")
   	
   	year = spllited(2)
   	month = spllited(0)
   	day = spllited(1)
   	
   	If Len(month) < 2 Then
   		month = "0" & month
   	End If

   	If Len(day) < 2 Then
   		day = "0" & day
   	End If
   	
   	YearMonthDay = year & "-" & month & "-" & day

End Function

Function Load(fileName As String) As Collection

	Dim lines As New Collection

	Dim iNumber As Integer
	Dim sLine As String
	
    iNumber = Freefile()
    
    Open fileName For Input As iNumber
    
    While Not EOF(iNumber)
        Line Input #iNumber, sLine
        
        If sLine <> "" Then
        	lines.add(sLine)
        End If
    Wend
    
    Close #iNumber
	
	Load = lines

End Function