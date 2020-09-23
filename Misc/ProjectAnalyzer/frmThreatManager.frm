VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThreatManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Threat Manager"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   3270
      TabIndex        =   7
      Top             =   4890
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Threat"
      Height          =   1035
      Left            =   900
      TabIndex        =   1
      Top             =   3630
      Width           =   5985
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Add"
         Height          =   345
         Left            =   4590
         TabIndex        =   6
         Top             =   510
         Width           =   1215
      End
      Begin VB.TextBox txtThreat 
         Height          =   315
         Left            =   870
         TabIndex        =   4
         Top             =   510
         Width           =   3615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   315
         Left            =   150
         Max             =   1
         Min             =   5
         TabIndex        =   2
         Top             =   510
         Value           =   5
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Level                                  Word(s)"
         Height          =   225
         Left            =   450
         TabIndex        =   5
         Top             =   270
         Width           =   4035
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   195
         Left            =   450
         TabIndex        =   3
         Top             =   570
         Width           =   285
      End
   End
   Begin MSComctlLib.ListView lstThreat 
      Height          =   3315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   5847
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Level"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Threat Word(s)"
         Object.Width           =   11112
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Threat"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete Threat"
      End
      Begin VB.Menu SEP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmThreatManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    SaveThreats
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    If Len(Trim(txtThreat.Text)) = 0 Then Exit Sub
    With lstThreat
        .Sorted = False
        .ListItems.Add , "D" & lstThreat.ListItems.Count, Val(lblLevel.Caption)
        .ListItems(.ListItems.Count).SubItems(1) = "|" & txtThreat.Text & "|"
        .Sorted = True
    End With
    txtThreat.Text = ""

End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    lstThreat.SortKey = 0
    LoadThreats

End Sub

Private Sub LoadThreats()
    Dim strParced() As String
    Dim strData As String
    Dim intLooper As Integer
    
    lstThreat.ListItems.Clear
    lstThreat.Sorted = False
    
    For intCounter = 1 To 5
        strData = GetINIData("ThreatWords", "DEFCON" & Trim(intCounter))
        
        strParced() = Split(strData, ",")
        
        For intLooper = 0 To UBound(strParced())
            With lstThreat
                .ListItems.Add , "D" & .ListItems.Count, intCounter
                .ListItems(.ListItems.Count).SubItems(1) = strParced(intLooper)
            End With
        Next intLooper
        
    Next intCounter
    
    lstThreat.Sorted = True
End Sub

Private Sub lstThreat_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then DeleteThreat
End Sub

Private Sub lstThreat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuDel.Visible = lstThreat.ListItems.Count > 0
        Me.PopupMenu mnuPop
    End If
End Sub

Private Sub mnuAdd_Click()
    txtThreat.SetFocus
End Sub

Private Sub mnuDel_Click()
    DeleteThreat
End Sub

Private Sub VScroll1_Change()
    lblLevel.Caption = 6 - VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    lblLevel.Caption = 6 - VScroll1.Value
End Sub

Private Sub DeleteThreat()
    If lstThreat.ListItems.Count < 1 Then Exit Sub
    If MsgBox("Are you sure you want to remove this threat from the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    lstThreat.ListItems.Remove (lstThreat.SelectedItem.Key)
End Sub

Private Sub SaveThreats()
    Dim strData(1 To 5) As String
    Dim intDefcon As Integer
    
    For intCounter = 1 To lstThreat.ListItems.Count
        intDefcon = Val(lstThreat.ListItems(intCounter).Text)
        If Len(strData(intDefcon)) = 0 Then
            strData(intDefcon) = lstThreat.ListItems(intCounter).SubItems(1)
        Else
            strData(intDefcon) = strData(intDefcon) & "," & lstThreat.ListItems(intCounter).SubItems(1)
        End If
    Next intCounter
    
    For intCounter = 1 To 5
        WriteDataToINI "ThreatWords", "DEFCON" & Trim(intCounter), strData(intCounter)
    Next intCounter
    
End Sub

