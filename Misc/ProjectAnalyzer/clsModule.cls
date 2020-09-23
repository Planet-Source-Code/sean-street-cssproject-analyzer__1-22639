VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strName As String
Dim strType As String
Dim strLoc As String
Dim strProcedures As String
Dim strControls As String
Dim strVariables As String
Dim strCode As String

Public Property Get ModName() As String
    ModName = strName
End Property

Public Property Let ModName(ByVal Value As String)
    strName = Value
End Property

Public Property Get ModType() As String
    ModType = strType
End Property

Public Property Let ModType(ByVal Value As String)
    strType = Value
End Property

Public Property Get ModLoc() As String
    ModLoc = strLoc
End Property

Public Property Let ModLoc(ByVal Value As String)
    strLoc = Value
End Property

Public Property Get ModCode() As String
    ModCode = strCode
End Property

Public Property Let ModCode(ByVal Value As String)
    strCode = Value
End Property

Public Function ProcedureCount() As Integer
    Dim strTemp() As String
    strTemp() = Split(strProcedures, ",")
    ProcedureCount = UBound(strTemp) + 1
 End Function

Public Function GetProcIndex(intIndex As Integer) As Integer
    Dim strTemp() As String
    
    strTemp() = Split(strProcedures, ",")
    
    GetProcIndex = strTemp(intIndex)
    
End Function

Public Sub AddProcedure(objProc As clsProcedure)
    objProc.ProcParent = strName
    colProcedures.Add objProc
    If Len(strProcedures) = 0 Then
        strProcedures = colProcedures.Count
    Else
        strProcedures = strProcedures & "," & colProcedures.Count
    End If
End Sub

Public Function VarCount() As Integer
    Dim strTemp() As String
    strTemp() = Split(strVariables, ",")
    VarCount = UBound(strTemp) + 1
 End Function

Public Function GetVarIndex(intIndex As Integer) As Integer
    Dim strTemp() As String
    
    strTemp() = Split(strVariables, ",")
    
    GetVarIndex = strTemp(intIndex)
    
End Function

Public Sub AddVariable(objVar As clsVariable)
    objVar.VarParent = strName
    colVariables.Add objVar
    If Len(strVariables) = 0 Then
        strVariables = colVariables.Count
    Else
        strVariables = strVariables & "," & colVariables.Count
    End If
End Sub

Public Sub AddControl(objControl As clsControl)
    objControl.CtrParent = strName
    colControls.Add objControl
    If Len(strControls) = 0 Then
        strControls = colControls.Count
    Else
        strControls = strControls & "," & colControls.Count
    End If
End Sub

Public Function ControlCount() As Integer
    Dim strTemp() As String
    strTemp() = Split(strControls, ",")
    ControlCount = UBound(strTemp) + 1
 End Function

Public Function GetCtrIndex(intIndex As Integer) As Integer
    Dim strTemp() As String
    
    strTemp() = Split(strControls, ",")
    
    GetCtrIndex = strTemp(intIndex)
    
End Function


Private Sub Class_Initialize()
    strProcedures = ""
    strControls = ""
    strVariables = ""
End Sub

