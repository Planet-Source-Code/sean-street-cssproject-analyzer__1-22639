VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim strName As String
Dim strScope As String
Dim strParent As String
Dim strCode As String

Public Property Get VarName() As String
    VarName = strName
End Property

Public Property Let VarName(ByVal Value As String)
    strName = Value
End Property

Public Property Get VarType() As String
    VarType = strType
End Property

Public Property Let VarType(ByVal Value As String)
    strType = Value
End Property

Public Property Get VarParent() As String
    VarParent = strParent
End Property

Public Property Let VarParent(ByVal Value As String)
    strParent = Value
End Property

Public Property Get VarScope() As String
    VarScope = strScope
End Property

Public Property Let VarScope(ByVal Value As String)
    strScope = Value
End Property

Public Property Get VarCode() As String
    VarCode = strCode
End Property

Public Property Let VarCode(ByVal Value As String)
    strCode = Value
End Property


