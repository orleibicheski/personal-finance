REM  *****  BASIC  *****

Option Compatible
Option ClassModule

Option Explicit

private _where As String
private hasWhere As Boolean

REM -----------------------------------------------------------------------------------------------------------------------
REM --- CONSTRUCTORS / DESTRUCTORS                                                                    ---
REM -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()

	hasWhere = False
	
End Sub      '   Constructor

REM -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

public Function oWhere(field As SqlField)

	If hasWhere <> True Then
		_where = "WHERE " & field.toSelect()
		hasWhere = True
	End If

End Function

public Function addOr(field As SqlField)

	_where = _where & " OR " & field.toSelect()
	
End Function


public Function addAnd(field As SqlField)

	_where = _where & " AND " & field.toSelect()

End Function

public Function toString()

	toString = _where
	
End Function
