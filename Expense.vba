REM  *****  BASIC  *****

Option Compatible
Option ClassModule

Option Explicit

private _expenseType As String
private _transactionID As String
private _exchangeID As String

private _date As String
private _month As String

private _fields As Collection

REM --------------------------------------------------------------
REM --- CONSTRUCTORS / DESTRUCTORS                             ---
REM --------------------------------------------------------------

Private Sub Class_Initialize()

	_expenseType = "57" ' The start value is Unspecified

	Set _fields = New Collection
	

End Sub      '   Constructor

Private Sub Class_Terminate()

End Sub

REM ---------------------------------------------------------------

Sub Expense

End Sub

Function buildBalance(line_ As String, transaction As String, exchange As String) 

	Dim columns
   	columns = Split(line_, ",")

   	_date = columns(0)

	_transactionID = transaction
	_exchangeID = exchange
   	
   	_month = Util.YearMonth(_date)
   	
   	createFields()

End Function

Function buildCredit(line_ As String, transaction As String, exchange As String) 

	Dim columns
   	columns = Split(line_, ",")

   	_date = columns(0)

	_transactionID = transaction
	_exchangeID = exchange
   	
   	_month = Util.YearMonth(_date)
   	
   	createFields()

End Function

private Function createFields() 

	Dim expenseTypeF As Object
	Dim transactionF As Object
	Dim exchangeF As Object
	Dim dateF As Object
	Dim monthF As Object

	Set expenseTypeF = New SqlField
	Set transactionF = New SqlField
	Set exchangeF = New SqlField
	Set dateF = New SqlField
	Set monthF = New SqlField

	REM Preparing SqlField to Insert
	expenseTypeF.oInsert("fk_ExpenseType_ID", _expenseType, "Integer")
	transactionF.oInsert("fk_Transaction_ID", _transactionID, "Integer")
	exchangeF.oInsert("fk_Exchange_ID", _exchangeID, "Integer")
	dateF.oInsert("dt_Date", _date, "Date")
	monthF.oInsert("txt_Month", _month, "String")
	
	REM Adding SqlFields in Collection
	_fields.Add(expenseTypeF)
	_fields.Add(transactionF)
	_fields.Add(exchangeF)
	_fields.Add(dateF)
	_fields.Add(monthF)

End Function

Function save()

	Sql.insert("tb_Expense", _fields)

End Function