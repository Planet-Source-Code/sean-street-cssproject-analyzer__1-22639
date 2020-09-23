VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim strName As String
Dim strScope As String
Public colArguments As Collection
Dim strCode As String
Dim strReturnVal As String
Dim strParent As String
Dim intParent As Integer

Public Sub AddArguments(objVariable As clsVariable)
    objVariable.VarParent = strName
    colArguments.Add objVariable
End Sub

Public Function ArgCount() As Integer
    ArgCount = colArguments.Count
End Function

Public Property Get ProcName() As String
    ProcName = strName
End Property

Public Property Let ProcName(ByVal Value As String)
    strName = Value
End Property

Public Property Get ProcType() As String
    ProcType = strType
End Property

Public Property Let ProcType(ByVal Value As String)
    strType = Value
End Property

Public Property Get ProcParent() As String
    ProcParent = strParent
End Property

Public Property Let ProcParent(ByVal Value As String)
    strParent = Value
End Property

Public Property Get ProcParentID() As Integer
    ProcParentID = intParent
End Property

Public Property Let ProcParentID(ByVal Value As Integer)
    intParent = Value
End Property

Public Property Get ProcScope() As String
    ProcScope = strScope
End Property

Public Property Let ProcScope(ByVal Value As String)
    strScope = Value
End Property

Public Property Get ProcCode() As String
    ProcCode = strCode
End Property

Public Property Let ProcCode(ByVal Value As String)
    strCode = Value
End Property

Public Property Get ProcReturn() As String
    ProcReturn = strReturnVal
End Property

Public Property Let ProcReturn(ByVal Value As String)
    strReturnVal = Value
End Property

Private Sub Class_Initialize()
    Set colArguments = New Collection
End Sub



