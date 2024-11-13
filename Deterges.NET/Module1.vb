Option Strict Off
Option Explicit On
Module Module1
	Public IniFile As String
	
	Public Declare Function GetPrivateProfileStringByKeyName Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpszKey As String, ByVal lpszDefault As String, ByVal lpszReturnBuffer As String, ByVal cchReturnBuffer As Integer, ByVal lpszFile As String) As Integer
	Public Declare Function WritePrivateProfileStringByKeyName Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer
	
	Function GetConfig(ByRef section As String, ByRef key As String, ByRef File As String) As String
		Dim KeyValue As String
		Dim characters As Integer
		
		IniFile = VB6.GetPath & "\config.ini"
		KeyValue = New String(Chr(0), 128)
		characters = GetPrivateProfileStringByKeyName(section, key, "", KeyValue, 127, File)
		
		If characters > 1 Then
			KeyValue = Left(KeyValue, characters)
		End If
		
		GetConfig = KeyValue
		
	End Function
	
	Public Function SaveConfig(ByRef SectionName As String, ByRef KeyName As String, ByRef KeyValue As String, ByRef FileName As String) As Integer
        IniFile = Application.CommonAppDataPath & "\config.ini"
		If Len(KeyValue) = 0 Then KeyValue = " "
		SaveConfig = WritePrivateProfileStringByKeyName(SectionName, KeyName, KeyValue, FileName)
	End Function
End Module