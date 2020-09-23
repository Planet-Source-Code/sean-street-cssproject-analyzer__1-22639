VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThreats 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Threats found"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstThreat 
      Height          =   4335
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lvl"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Word(s)"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Location"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuJump 
         Caption         =   "&Jump to Procedure"
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmThreats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    mnuPopup.Visible = False
    If GetINIData("ThreatWords", "WindowLocation") = 0 Then
        Me.Left = frmMain.Left
        Me.Top = frmMain.Height + frmMain.Top
        Me.Width = frmMain.Width
        Me.Height = 2500
    Else
        If frmMain.Left > 0 Then
            frmMain.Left = Me.Width / 2
            Me.Left = frmMain.Width + frmMain.Left
            Me.Top = frmMain.Top
            Me.Height = frmMain.Height
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With lstThreat
        .Width = Me.Width - 270
        .Height = Me.Height - 540
        .ColumnHeaders(2).Width = .Width * 0.31
        .ColumnHeaders(3).Width = .Width * 0.57
    End With
End Sub

Private Sub lstThreat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Me.PopupMenu mnuPopup
End Sub

Private Sub mnuJump_Click()
    Dim strTemp As String
    
    strTemp = Mid(lstThreat.SelectedItem.Key, 2, InStr(1, lstThreat.SelectedItem.Key, "X") - 2)
    
    With frmMain
        .trvView.Nodes("PP" & strTemp).Selected = True
        .txtCode.Text = colProcedures(Val(strTemp)).ProcCode
        .MapProcedure (Val(strTemp))
    End With
End Sub
