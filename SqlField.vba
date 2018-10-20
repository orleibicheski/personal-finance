REM  *****  BASIC  *****


Option Compatible
Option ClassModule

Option Explicit

private _name As String
private _data As String
private _typeField As String
private _operator As String
private _clause As String

REM -----------------------------------------------------------------------------------------------------------------------
REM --- CONSTRUCTORS / DESTRUCTORS                                                                    ---
REM -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()


End Sub      '   Constructor

REM -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

Function oInsert(nm As String, dt As String, tf As String)

	_name = nm
	_data = dt
	_typeField = tf

End Function


Function oSelect(nm As String, dt As String, tf As String, op As String, cl As String)

	_name = nm
	_data = dt
	_typeField = tf
	_operator = op
	_clause = cl

End Function


Public Property Get value As String

	value = _data

End Property

Public Property Get typeField As String

	typeField = _typeField

End Property

Public Property Get field As String

	field = _name

End Property

Public Property Get clause As String

	clause = _clause

End Property

public Function toInsert() As String

	toInsert = " """ & _name & """ "

End Function

public Function toSelect() As String

	toSelect = " """ & _name & """ " & _operator & " ? "

End Function
