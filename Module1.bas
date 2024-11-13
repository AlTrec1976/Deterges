Attribute VB_Name = "Module1"
Global IniFile As String

Public Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Public Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)

Function GetConfig(section$, key$, File$) As String
    Dim KeyValue$
    Dim characters As Long
    
    IniFile = App.Path & "\config.ini"
    KeyValue$ = String$(128, 0)
    characters = GetPrivateProfileStringByKeyName(section$, key$, "", KeyValue$, 127, File$)

    If characters > 1 Then
        KeyValue$ = Left$(KeyValue$, characters)
    End If
    
    GetConfig = KeyValue$

End Function

Public Function SaveConfig(SectionName As String, KeyName As String, KeyValue As String, FileName As String) As Long
    IniFile = App.Path & "\config.ini"
    If Len(KeyValue) = 0 Then KeyValue = " "
    SaveConfig = WritePrivateProfileStringByKeyName(SectionName, KeyName, KeyValue, FileName)
End Function
