REM  *****  BASIC  *****

Option Compatible
Option ClassModule

Option Explicit

private _date As String
private _transaction As String
private _description As String
private _value As String

private _transactionType As String
private _currency As String
private _origin As String

private _fields As Collection

private _transactionID As String
private _exchangeID As String

REM --------------------------------------------------------------
REM --- CONSTRUCTORS / DESTRUCTORS                             ---
REM --------------------------------------------------------------

Private Sub Class_Initialize()

	_transaction = "Credit" ' Credit is standard value
	_currency = "0" ' Default Currency is Canadian Dollar
	_origin = "2" ' Default Origin to credit is Scotiabank Visa

	Set _fields = New Collection

End Sub      '   Constructor

Private Sub Class_Terminate()

End Sub

REM ---------------------------------------------------------------

Sub Credit

End Sub

Public Property Get fields As String

	fields = _fields

End Property

Public Property Get transactionID As String

	transactionID = _transactionID

End Property

Public Property Get exchangeID As String

	exchangeID = _exchangeID

End Property

Function build(line_ As String) 

	Dim columns
	
   	columns = Split(line_, ",")

   	_date = columns(0)
   	_description = columns(1)
   	_value = columns(2)
   	
   	If Util.Contains(_value, "-") Then 
   		_transactionType = "0" ' Value is negative and Transaction Type will be define as Withdrawal
   	Else
   		_transactionType = "2" ' Value is negative and Transaction Type will be define as Credit
   	End If
   	
   	createFields()

End Function

private Function createFields() 

	Dim dateF As Object
	Dim transactionF As Object
	Dim descriptionF As Object
	Dim valueF As Object
	Dim transactionTypeF As Object
	Dim currencyF As Object
	Dim originF As Object

	Set dateF = New SqlField
	Set transactionF = New SqlField
	Set descriptionF = New SqlField
	Set valueF = New SqlField
	Set transactionTypeF = New SqlField
	Set currencyF = New SqlField
	Set originF = New SqlField

	REM Preparing SqlField to Insert
	dateF.oInsert("dt_Date", _date, "Date")
	transactionF.oInsert("txt_Transaction", _transaction, "String")
	descriptionF.oInsert("txt_Description", _description, "String")
	valueF.oInsert("fl_Value", _value, "Double")

	transactionTypeF.oInsert("fk_TransactionType_ID", _transactionType, "Integer")
	currencyF.oInsert("fk_Currency_ID", _currency, "Integer")
	originF.oInsert("fk_Origin_ID", _origin, "Integer")
	
	REM Adding SqlFields in Collection
	_fields.Add(dateF)
	_fields.Add(transactionF)
	_fields.Add(descriptionF)
	_fields.Add(valueF)
	
	_fields.Add(transactionTypeF)
	_fields.Add(currencyF)
	_fields.Add(originF)

End Function


Function save()

	Sql.insert("tb_Transaction", _fields)

	Dim expenseF As Object
	Set expenseF = New Collection

	Dim transactionID
	Dim exchangeID
		
	transactionID = Sql.lastRecord("tb_Transaction")
	_transactionID = CStr(transactionID)

	Dim dateF As Object
	Dim currencyFromF As Object
	Dim currencyToF As Object

	Set dateF = New SqlField
	Set currencyFromF = New SqlField
	Set currencyToF = New SqlField

	dateF.oSelect("dt_Date", _date, "Date", "<=", "")
	currencyFromF.oSelect("fk_Currency_From", "0", "Integer", "=", "AND") ' Defaul Exchange Canadian Dollar 
	currencyToF.oSelect("fk_Currency_To", "1", "Integer", "=", "AND") ' Defaul Exchange Brazilian Real

	expenseF.Add(dateF)
	expenseF.Add(currencyFromF)
	expenseF.Add(currencyToF)

	exchangeID = Sql.bestExchange("tb_Exchange", expenseF) 
	_exchangeID = CStr(exchangeID)

End Function
