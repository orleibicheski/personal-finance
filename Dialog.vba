REM  *****  BASIC  *****

Option Explicit

Dim dialog As Object
 
Sub Main

	DialogLibraries.LoadLibrary("Standard")
	dialog = CreateUnoDialog(DialogLibraries.Standard.LoadFiles)
	dialog.execute()

End Sub

Sub btn_close()

	dialog.endexecute()
	dialog.dispose()

End Sub

Sub btn_chequing()

	execute("Balance")

End Sub

Sub btn_credit()

	execute("Credit")

End Sub

Function execute(layout_ As String)

	Dim lines As Object
	Dim file As String

	file = fileName(layout_)
	If Util.IsBlank(file) Then
		Exit Function
	End If
	
	lines = Util.Load(file)
	
	Dim i As Integer
	
	For i = 1 To lines.Count
	
		Select Case layout_

		Case "Balance"
			createBalance(lines(i))
			MsgBox "File balance loaded!"
		
		Case "Credit"
			createCredit(lines(i))
			MsgBox "File credit loaded!"
		
		End Select
    	
	Next i

	registerFile(file)
	
End Function

Function registerFile(file_ As String)

	Dim file As Object
	Set file = New File
	
	file.build(getSingleName(file_))
	file.save()
	
End Function

Function createBalance(line_ As String)

	Dim balance As Object
	Set balance = New Balance
	
	balance.build(line_)
	balance.save()
	
	Dim expense As Object	
	Set expense = New Expense

	expense.buildBalance(line_, balance.transactionID(), balance.exchangeID())
	expense.save()
	

End Function

Function createCredit(line_ As String)

	Dim credit As Object
	Set credit = New Credit
	
	credit.build(line_)
	credit.save()

	Dim expense As Object	
	Set expense = New Expense

	expense.buildCredit(line_, credit.transactionID(), credit.exchangeID())
	expense.save()

End Function

Function fileName(search As String)

	Dim file As String
	
	file = dialog.GetControl("txt_fileCntrl").Text
	
	REM Check if File Name is blank
	If Util.IsBlank(file) Then
		MsgBox "Please, select a file!"
		Exit Function
	End If
	
	REM Check if File Name has words "Balance" or "Credit"
	If Util.Contains(file, search) Then
		MsgBox "It isn't " & search & " file!"
		Exit Function
	End If

	REM Check if File already was loaded
	Dim singleName
	Dim fileNameSql As Object
	Dim fields As Object
	Dim resultSet

	singleName = getSingleName(file)
	
	Set fileNameSql = New SqlField
	Set fields = New Collection
	
	fileNameSql.oSelect("txt_FileName", singleName, "String", "=", "")
	fields.add(fileNameSql)
	
	resultSet= Sql.query("tb_File", fields)

	If resultSet.next() Then
	
		MsgBox "File " & file & " was loaded!"
		Exit Function

	End If

	fileName = file

End Function

Function getSingleName(file_ As String)

	Dim splitted
	splitted = Split(file_, "/")

	getSingleName = splitted(UBound(splitted))
	
End Function
