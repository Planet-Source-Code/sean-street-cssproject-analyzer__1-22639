VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   1230
      TabIndex        =   6
      Top             =   2580
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Threat window location..."
      Height          =   795
      Left            =   150
      TabIndex        =   3
      Top             =   1560
      Width           =   3315
      Begin VB.OptionButton optLoc 
         Caption         =   "Right"
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   5
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton optLoc 
         Caption         =   "Bottom"
         Height          =   315
         Index           =   0
         Left            =   540
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "When opening project..."
      Height          =   1245
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3315
      Begin VB.CheckBox chkOptions 
         Caption         =   "Automatically search for errors"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   2955
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Automatically search for threat words"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intOptLoc As Integer

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    chkOptions(0).Value = Val(GetINIData("Options", "ThreatOnOpen"))
    chkOptions(1).Value = Val(GetINIData("Options", "ErrorsOnOpen"))
    optLoc(Val(GetINIData("ThreatWords", "WindowLocation"))).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteDataToINI "Options", "ThreatOnOpen", chkOptions(0).Value
    WriteDataToINI "Options", "ErrorsOnOpen", chkOptions(1).Value
    WriteDataToINI "ThreatWords", "WindowLocation", intOptLoc
End Sub

Private Sub optLoc_Click(Index As Integer)
    intOptLoc = Index
End Sub
