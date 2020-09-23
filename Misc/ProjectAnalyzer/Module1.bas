Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public strProjectPath As String
Public strProjectName As String
Public strProjectFileName As String
Public colModules As Collection
Public intCounter As Integer
Public colProcedures As Collection
Public colControls As Collection
Public colVariables As Collection
Public lngHeigth As Long
Public lngWidth As Long

Public Sub WriteDataToINI(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)
    Dim strFileLoc As String
    
    strFileLoc = App.Path + "\" & "ProjectAnalyzer.INI"
    
    WritePrivateProfileString strSection, strKey, strValue, strFileLoc

End Sub

Public Function GetINIData(strSection As String, strKey As String) As String
    Dim strBuffer As String
    Dim strFileLoc As String
    
    strBuffer = Space(145)
    strFileLoc = App.Path + "\" & "ProjectAnalyzer.INI"
    GetPrivateProfileString strSection, strKey, "", strBuffer, 144, strFileLoc
    
    GetINIData = Left(strBuffer, (InStr(1, strBuffer, Chr(0)) - 1))
End Function

