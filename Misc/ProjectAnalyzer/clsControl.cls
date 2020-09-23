VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim strName As String
Dim strParent As String
Dim strCode As String

Public Property Get CtrName() As String
    CtrName = strName
End Property

Public Property Let CtrName(ByVal Value As String)
    strName = Value
End Property

Public Property Get CtrType() As String
    CtrType = strType
End Property

Public Property Let CtrType(ByVal Value As String)
    strType = Value
End Property

Public Property Get CtrParent() As String
    CtrParent = strParent
End Property

Public Property Let CtrParent(ByVal Value As String)
    strParent = Value
End Property

Public Property Get CtrCode() As String
    CtrCode = strCode
End Property

Public Property Let CtrCode(ByVal Value As String)
    strCode = Value
End Property

