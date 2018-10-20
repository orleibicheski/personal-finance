REM  *****  BASIC  *****

Option Compatible
Option ClassModule

Option Explicit

private _fields As Collection

REM --------------------------------------------------------------
REM --- CONSTRUCTORS / DESTRUCTORS                             ---
REM --------------------------------------------------------------

Private Sub Class_Initialize()

	Set _fields = New Collection

End Sub      '   Constructor

Private Sub Class_Terminate()

End Sub

REM ---------------------------------------------------------------

Sub File

End Sub

Function build(fileName_ As String) 

	Dim fileNameF As Object
	Set fileNameF = New SqlField

	REM Preparing SqlField to Insert
	fileNameF.oInsert("txt_FileName", fileName_, "String")
	
	REM Adding SqlFields in Collection
	_fields.Add(fileNameF)

End Function

Function save()

	Sql.insert("tb_File", _fields)

End Function
