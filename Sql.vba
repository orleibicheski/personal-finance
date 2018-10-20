REM  *****  BASIC  *****

Option Explicit

Sub Sql

End Sub

Function oConnection()

	Dim oContext
	Dim oDB
	Dim oConn

	'CREATE A DATABASE CONTEXT OBJECT	
	oContext = CreateUnoService("com.sun.star.sdb.DatabaseContext")

	'GET DATABASE BY NAME
	oDB = oContext.getByName("Finance")
	
	'ESTABLISH CONNECTION TO DATABASE
	oConn = oDB.getConnection("","")
	
	' oDatasource = thisDatabaseDocument.CurrentController
	' If Not (oDatasource.isConnected()) Then 
	' 	oDatasource.connect()
	' End If
	
	' oConn = oDatasource.ActiveConnection()
	
	oConnection = oConn

End Function

Function insert(tableName_ As String, fields_ As Collection)

	Dim oConn
	Dim statement
	
	Dim sql As String
	Dim fields As String
	Dim values As String
	
	oConn = oConnection()
	
	sql = "INSERT INTO """ & tableName_  & """ "
	
	fields = "("
	values = "VALUES("

	Dim field As Object
	Dim i
			
	for i = 1 to fields_.count
		
		If i = 1 Then
			fields = fields & fields_(i).toInsert()
			values = values & " ? "
		Else
			fields = fields & ", " & fields_(i).toInsert()
			values = values & ", ? "
		End If
	
	Next i

	fields = fields & ") "
	values = values & ") "
	sql = sql & fields & values
	
	statement = oConn.prepareStatement(sql)

	for i = 1 to fields_.count
		
		setVal(statement, i, fields_(i))
	
	Next i
	
	insert = statement.executeUpdate()
	
End Function

Function query(tableName_ As String, fields_ As Collection)

	query = queryWithSelect(tableName_, "*", fields_)

End Function

Function queryWithSelect(tableName_ As String, select_ As String, fields_ As Collection)

	Dim oConn
	Dim statement
	
	Dim where As Object
	
	Dim sql As String
	Dim fields As String
	Dim values As String
	
	oConn = oConnection()
	
	sql = "SELECT " & select_ & " FROM """ & tableName_  & """ "
	
	Set where = New SqlWhere
	
	Dim i
	
	for i = 1 to fields_.count
		
		If i = 1 Then
			where.oWhere(fields_(i))
		Else
			Select Case fields_(i).clause()

			Case "AND"
				where.addAnd(fields_(i))
			
			Case "OR"
				where.addOr(fields_(i))
			
			End Select
		End If
	
	Next i
	
	sql = sql & where.toString()
	
	statement = oConn.prepareStatement(sql)

	for i = 1 to fields_.count
		
		setVal(statement, i, fields_(i))
	
	Next i
	
	queryWithSelect = statement.executeQuery()
	
End Function

Function lastRecord(tableName_ As String)

	Dim oConn
	Dim statement
	Dim result
	
	oConn = oConnection()
	
	statement = oConn.prepareStatement("SELECT MAX(ID) AS id FROM """ & tableName_  & """ ")
	
	result = statement.executeQuery()
	If Not IsNull(result) Then
		If result.next() Then
			lastRecord = result.getInt(1)
			
			Exit Function
		End If
	End If
	
	lastRecord = 0
	
End Function

Function bestExchange(tableName_ As String, fields_ As Collection)

	Dim oConn
	Dim statement
	Dim result
	Dim where As Object
	
	Dim sql As String
	
	oConn = oConnection()
	
	sql = "SELECT MAX(ID) AS id FROM """ & tableName_  & """ "

	Set where = New SqlWhere
	
	Dim i
	
	for i = 1 to fields_.count
		
		If i = 1 Then
			where.oWhere(fields_(i))
		Else
			Select Case fields_(i).clause()

			Case "AND"
				where.addAnd(fields_(i))
			
			Case "OR"
				where.addOr(fields_(i))
			
			End Select
		End If
	
	Next i
	
	sql = sql & where.toString()

	statement = oConn.prepareStatement(sql)
	
	for i = 1 to fields_.count
		
		setVal(statement, i, fields_(i))
	
	Next i

	result = statement.executeQuery()

	If Not IsNull(result) Then
		If result.next() Then
			bestExchange = result.getInt(1)
			
			Exit Function
		End If
	End If
	
	bestExchange = 0
	
End Function

Function setVal(statement As Object, index As Integer, field As SqlField)

	Dim value As String
	value = field.value()
	
	Select Case field.typeField
	
	Case "String"
		statement.setString(index, value)
	
	Case "Integer"
		statement.setInt(index, CInt(value))
	
	Case "Long"
		statement.setLong(index, CLng(value))
	
	Case "Double"
		Dim dbl As Double
		dbl = CDbl(value)
		
		If dbl < 0 Then
			dbl = (dbl * -1)
		End If
		statement.setDouble(index, dbl)
	
	Case "Date"
		Dim startDate As Date
		Dim aDate As New com.sun.star.util.Date
		
		startDate = CDate(Util.YearMonthDay(value))
		' Need to convert 'startDate' into a util.Date UNO structure
		' before using it with to bind the parameter with setDate()
		With aDate
		   .Year = Year(startDate)
		   .Month = Month(startDate)
		   .Day = Day(startDate)
		End With
		
		statement.setDate(index, aDate)
	
	End Select

End Function
