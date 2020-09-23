VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "No Project Loaded"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgMain 
      Left            =   5670
      Top             =   5550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1138
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158A
            Key             =   "Procedure"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DC
            Key             =   "Attribute"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E2E
            Key             =   "Argument"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2280
            Key             =   "Control"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   4515
      Left            =   4050
      TabIndex        =   3
      Top             =   210
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   7964
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Procedures that I call"
      TabPicture(0)   =   "frmMain.frx":2512
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstProcMapping(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Procedures that call me"
      TabPicture(1)   =   "frmMain.frx":252E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstProcMapping(1)"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView lstProcMapping 
         Height          =   3600
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   6350
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
            Text            =   "Parent"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Procedure Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Possible Errors"
            Object.Width           =   4057
         EndProperty
      End
      Begin MSComctlLib.ListView lstProcMapping 
         Height          =   3600
         Index           =   1
         Left            =   -74760
         TabIndex        =   5
         Top             =   660
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   6350
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
            Text            =   "Parent"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Procedure Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Possible Errors"
            Object.Width           =   4057
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   2865
      Left            =   150
      TabIndex        =   2
      Top             =   5040
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   5054
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":254A
   End
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   330
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView trvView 
      Height          =   4485
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   7911
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   5
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Caption         =   "Code of selected module/procedure"
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   4830
      Width           =   9735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintC 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu SEP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuThreatMan 
         Caption         =   "&Threat Manager"
      End
      Begin VB.Menu SEP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScan 
         Caption         =   "&Scan for Errors..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuThreat 
         Caption         =   "&View Possible Threats..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu SEP4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuJump 
         Caption         =   "&Jump to Procedure"
      End
      Begin VB.Menu SEP5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objModule As clsModule
Dim intProcMapping As Integer

Private Sub Form_Load()
    lngHeigth = 8730
    lngWidth = Me.Width
    trvView.ImageList = imgMain
    mnuPopup.Visible = False
    mnuPrintC.Enabled = False
    
    OpenProject
    
End Sub

Private Sub LoadProject()
    Dim intFreeFile As Integer
    Dim strData As String
    Dim intCounter2 As Integer
    Dim intLooper As Integer
    
    intFreeFile = FreeFile
    'opening Visual Basic project file
    Open strProjectPath & strProjectFileName For Input As intFreeFile
    
    Do While Not EOF(intFreeFile)
        'looping through each line within the VBP file
        'to determine their purpose in life
        Line Input #intFreeFile, strData
        If InStr(1, UCase(strData), "FORM=") = 1 Then
            'we have a reference to a form here, let's add to our
            'module collection
            Set objModule = New clsModule
            objModule.ModLoc = Replace(Mid(strData, InStr(1, strData, "=") + 1), Chr(34), "")
            objModule.ModType = "Form"
            'now let's load any procedures, controls or variables for this module
            LoadModule strProjectPath & objModule.ModLoc
            colModules.Add objModule
            Set objModule = Nothing
        ElseIf InStr(1, UCase(strData), "MODULE=") = 1 Then
            'we have a module here, let's add it to our
            'module collection
            Set objModule = New clsModule
            If InStr(1, strData, ";") > 1 Then
                'sometimes visual basic places the name
                'within the VBP file, so let's parce it
                objModule.ModLoc = Trim(Replace(Mid(strData, InStr(1, strData, ";") + 1), Chr(34), ""))
            Else
                objModule.ModLoc = Replace(Mid(strData, InStr(1, strData, "=") + 1), Chr(34), "")
            End If
            'now let's load any procedures, controls or variables for this module
            objModule.ModType = "Module"
            LoadModule strProjectPath & objModule.ModLoc
            colModules.Add objModule
            Set objModule = Nothing
        ElseIf InStr(1, UCase(strData), "CLASS=") = 1 Then
            'we have a class module here, let's add it to our
            'module collection
            Set objModule = New clsModule
            If InStr(1, strData, ";") > 1 Then
                'sometimes visual basic places the name
                'within the VBP file, so let's parce it
                objModule.ModLoc = Trim(Replace(Mid(strData, InStr(1, strData, ";") + 1), Chr(34), ""))
            Else
                objModule.ModLoc = Replace(Mid(strData, InStr(1, strData, "=") + 1), Chr(34), "")
            End If
            objModule.ModType = "Class"
            'now let's load any procedures, controls or variables for this module
            LoadModule strProjectPath & objModule.ModLoc
            colModules.Add objModule
            Set objModule = Nothing
        ElseIf InStr(1, UCase(strData), "NAME=") = 1 Then
            strProjectName = Replace(Mid(strData, InStr(1, strData, "=") + 1), Chr(34), "")
        End If
    Loop
    
    Close
    
    With trvView.Nodes
        .Add , , "ROOT", strProjectName, 1
        
        intLooper = 0
        For Each objModule In colModules
            intLooper = intLooper + 1
            
            .Add "ROOT", tvwChild, "M" & intLooper, objModule.ModName & ":" & objModule.ModType, objModule.ModType
            
            'adding procedures here
            For intCounter = 0 To objModule.ProcedureCount - 1
                .Add "M" & intLooper, tvwChild, "PP" & objModule.GetProcIndex(intCounter), colProcedures(objModule.GetProcIndex(intCounter)).ProcName, "Procedure"
                .Add "PP" & objModule.GetProcIndex(intCounter), tvwChild, "PS" & objModule.GetProcIndex(intCounter), "Scope: " & colProcedures(objModule.GetProcIndex(intCounter)).ProcScope, "Attribute"
                .Add "PP" & objModule.GetProcIndex(intCounter), tvwChild, "PT" & objModule.GetProcIndex(intCounter), "Type: " & colProcedures(objModule.GetProcIndex(intCounter)).ProcType, "Attribute"
                .Add "PP" & objModule.GetProcIndex(intCounter), tvwChild, "PR" & objModule.GetProcIndex(intCounter), "ReturnType: " & colProcedures(objModule.GetProcIndex(intCounter)).ProcReturn, "Attribute"
                .Add "PP" & objModule.GetProcIndex(intCounter), tvwChild, "PA" & objModule.GetProcIndex(intCounter), "Arguments: " & colProcedures(objModule.GetProcIndex(intCounter)).ArgCount, "Attribute"
                
                For intCounter2 = 1 To colProcedures(objModule.GetProcIndex(intCounter)).ArgCount
                    .Add "PA" & objModule.GetProcIndex(intCounter), tvwChild, , colProcedures(objModule.GetProcIndex(intCounter)).colArguments(intCounter2).VarName, "Argument"
                Next intCounter2
            Next intCounter
            
            'adding controls here
            For intCounter = 0 To objModule.ControlCount - 1
                .Add "M" & intLooper, tvwChild, "C" & objModule.GetCtrIndex(intCounter), colControls(objModule.GetCtrIndex(intCounter)).CtrName & ":" & colControls(objModule.GetCtrIndex(intCounter)).CtrType, "Control"
            Next intCounter
            
            'adding modular level variables here
            For intCounter = 0 To objModule.VarCount - 1
                .Add "M" & intLooper, tvwChild, "V" & objModule.GetVarIndex(intCounter), colVariables(objModule.GetVarIndex(intCounter)).VarName & ":" & colVariables(objModule.GetVarIndex(intCounter)).VarType, "Argument"
            Next intCounter
            
        Next

    End With

    Set objModule = Nothing
End Sub

Private Sub LoadModule(strPath As String)
    Dim intFreeFile As Integer
    Dim strData As String
    Dim objControl As clsControl
    Dim objProcedure As clsProcedure
    Dim objVariable As clsVariable
    Dim strCode As String
    Dim strModCode As String
    Dim intStart As Integer
    Dim strTemp As String
    Dim strArguments() As String
    
    intFreeFile = FreeFile
    Open strPath For Input As intFreeFile
    strModCode = ""
    
    Do While Not EOF(intFreeFile)
        Line Input #intFreeFile, strData
        If InStr(1, UCase(strData), "ATTRIBUTE VB_NAME = ") = 1 Then
            'we have the name of the module here;
            'let's set the name property
            objModule.ModName = Replace(Mid(strData, InStr(1, UCase(strData), "ATTRIBUTE VB_NAME = ") + 20), Chr(34), "")
        ElseIf (InStr(1, UCase(Trim(strData)), "BEGIN VB.") = 1) And (InStr(1, UCase(Trim(strData)), UCase(objModule.ModType)) = 0) Then
            'we have a control of some sort here;
            'let's add it to our control collection
            Set objControl = New clsControl
            intStart = InStr(1, UCase(strData), "BEGIN VB.") + 9
            With objControl
                .CtrType = Mid(strData, intStart, InStr(intStart, strData, " ") - intStart)
                .CtrName = Trim(Mid(strData, InStr(intStart, strData, " ")))
                .CtrCode = "Name: " & .CtrName & vbCrLf & "Type: " & .CtrType
            End With
            objModule.AddControl objControl
            Set objControl = Nothing
        ElseIf ((InStr(1, UCase(strData), " SUB ") > 1) Or (InStr(1, UCase(strData), " FUNCTION ") > 1) Or (InStr(1, UCase(strData), " PROPERTY ") > 1)) And (InStr(1, strData, "(") > 5) And (InStr(1, strData, ")") > 6) Then
            'we have a procedure here;
            'let's add it to our procedure collection
            strCode = strData & vbCrLf
            
            Set objProcedure = New clsProcedure
            With objProcedure
                .ProcParentID = colModules.Count + 1
                .ProcScope = "Private"
                If (InStr(1, UCase(strData), "PRIVATE ")) Or (InStr(1, UCase(strData), "PUBLIC ")) Or (InStr(1, UCase(strData), "FRIEND ")) Or (InStr(1, UCase(strData), "STATIC ")) Then .ProcScope = Left(strData, InStr(1, strData, " ") - 1)
                
                intStart = InStr(1, UCase(strData), "SUB ") + InStr(1, UCase(strData), "FUNCTION ") + InStr(1, UCase(strData), "PROPERTY ")
                
                .ProcType = Mid(strData, intStart, (InStr(intStart + 1, strData, " ")) + (Abs((InStr(1, UCase(strData), "PROPERTY ") > 0) * 4)) - intStart)
                                
                intStart = (InStr(intStart + 1, strData, " ")) + (Abs((InStr(1, UCase(strData), " PROPERTY GET ") > 0) * 4)) + (Abs((InStr(1, UCase(strData), " PROPERTY LET ") > 0) * 4)) + (Abs((InStr(1, UCase(strData), " PROPERTY SET ") > 0) * 4)) + 1

                .ProcName = Mid(strData, intStart, InStr(1, strData, "(") - intStart)
                
                intStart = InStr(1, strData, "(") + 1
                strTemp = Mid(strData, intStart, InStr(intStart, strData, ")") - intStart)
                If Len(strTemp) > 2 Then
                    strArguments() = Split(strTemp, ",")
                    For intCounter = 0 To UBound(strArguments)
                        Set objVariable = New clsVariable
                        With objVariable
                            .VarName = Trim(strArguments(intCounter))
                            intStart = InStr(1, UCase(strArguments(intCounter)), " AS ") + 4
                            If InStr(intStart, strArguments(intCounter), " = ") Then
                                .VarType = Mid(strArguments(intCounter), intStart, InStr(intStart, strArguments(intCounter), " = ") - intStart)
                            Else
                                .VarType = Mid(strArguments(intCounter), intStart)
                            End If
                        End With
                        .AddArguments objVariable
                        Set objVariable = Nothing
                    Next intCounter
                End If
                
                .ProcReturn = "None"
                intStart = InStr(1, strData, ")") + 1
                If InStr(intStart, UCase(strData), " AS ") Then .ProcReturn = Mid(strData, intStart + 4)
                
                If InStr(1, UCase(strData), "DECLARE") = 0 Then

                    If InStr(1, .ProcType, " ") > 0 Then
                        strTemp = UCase(Left(.ProcType, InStr(1, .ProcType, " ") - 1))
                    Else
                        strTemp = UCase(.ProcType)
                    End If
                    
                    Line Input #intFreeFile, strData
                    strCode = strCode & strData & vbCrLf
                    
                    Do While ((InStr(1, UCase(strData), "END " & strTemp)) = 0)
                        Line Input #intFreeFile, strData
                        strCode = strCode & strData & vbCrLf
                    Loop
                End If
                 .ProcCode = strCode

            End With
            objModule.AddProcedure objProcedure
            Set objProcedure = Nothing
            strModCode = strModCode & strCode & vbCrLf
        ElseIf ((InStr(1, UCase(strData), "PRIVATE ")) Or (InStr(1, UCase(strData), "PUBLIC ")) Or (InStr(1, UCase(strData), "DIM ")) Or (InStr(1, UCase(strData), "GLOBAL "))) And ((InStr(1, UCase(strData), " TYPE ") = 0)) And ((InStr(1, UCase(strData), " ENUM ") = 0)) Then

            'we have a variable here;
            'let's add it to our variable collection
            Set objVariable = New clsVariable
            With objVariable
                .VarScope = Left(strData, InStr(1, strData, " ") - 1)
                intStart = Len(.VarScope) + InStr(1, UCase(strData), "CONST") + 1
                .VarName = Trim(Mid(strData, intStart, InStr(intStart, UCase(strData), " AS ") - intStart))
                intStart = InStr(1, UCase(strData), " AS ") + 4
                If InStr(intStart, strData, " ") > 0 Then
                    .VarType = Mid(strData, intStart, InStr(intStart, strData, " ") - intStart)
                Else
                    .VarType = Mid(strData, intStart)
                End If
                .VarCode = "Location: " & objModule.ModName & vbCrLf & "Scope: " & .VarScope & vbCrLf & "Name: " & .VarName & vbCrLf & "Type: " & .VarType
            End With
            
            
            objModule.AddVariable objVariable
            Set objVariable = Nothing
            
            strModCode = strModCode & strData & vbCrLf
        End If
        
    Loop
    
    objModule.ModCode = strModCode
    
    Close intFreeFile
    
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Me.Height = lngHeigth
        Me.Width = lngWidth
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colModules = Nothing
    Set colProcedures = Nothing
    Set colControls = Nothing
    Set colVariables = Nothing
    Close
End Sub

Private Sub lstProcMapping_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 2) And (lstProcMapping(Index).ListItems.Count > 0) Then
        intProcMapping = Index
        mnuJump.Visible = True
        mnuPrint.Visible = False
        Me.PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuJump_Click()
    Dim strTemp As String
    strTemp = Mid(lstProcMapping(intProcMapping).SelectedItem.Key, 2, InStr(1, lstProcMapping(intProcMapping).SelectedItem.Key, "X") - 2)
    trvView.Nodes("PP" & strTemp).Selected = True
    txtCode.Text = colProcedures(Val(strTemp)).ProcCode
    MapProcedure (Val(strTemp))
End Sub

Private Sub mnuOpen_Click()
    OpenProject
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo EvalErr
    With cdgMain
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
    
        If txtCode.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        Printer.Print ""
        txtCode.SelPrint .hDC
    End With
    
    Exit Sub
    
EvalErr:
    If Err.Number <> 32755 Then MsgBox Err.Number & vbCr & Err.Description
    Exit Sub
End Sub

Private Sub mnuPrintC_Click()
    mnuPrint_Click
End Sub

Private Sub mnuThreat_Click()
    PopulateThreats
    frmThreats.Show
End Sub

Private Sub mnuThreatMan_Click()
    frmThreatManager.Show vbModal
End Sub

Private Sub trvView_NodeClick(ByVal Node As MSComctlLib.Node)
    'a node was clicked, so we need to show the code/information
    
    If InStr(1, Node.Key, "P") = 1 Then
        txtCode.Text = colProcedures(Val(Mid(Node.Key, 3))).ProcCode
        MapProcedure (Val(Mid(Node.Key, 3)))
    ElseIf InStr(1, Node.Key, "M") = 1 Then
        txtCode.Text = colModules(Val(Mid(Node.Key, 2))).ModCode
    ElseIf InStr(1, Node.Key, "C") = 1 Then
        txtCode.Text = "ACTIVE X CONTROL" & vbCrLf & colControls(Val(Mid(Node.Key, 2))).CtrCode
    ElseIf InStr(1, Node.Key, "V") = 1 Then
        txtCode.Text = "VARIABLE/OBJECT" & vbCrLf & colVariables(Val(Mid(Node.Key, 2))).VarCode
    Else
        txtCode.Text = ""
        lstProcMapping(0).ListItems.Clear
        lstProcMapping(1).ListItems.Clear
    End If
    
    mnuPrintC.Enabled = Len(txtCode.Text) > 0
    
End Sub

Private Sub OpenProject()
    On Error GoTo EvalErr
    
    Set colModules = Nothing
    Set colProcedures = Nothing
    Set colControls = Nothing
    Set colVariables = Nothing
    trvView.Nodes.Clear
    Unload frmThreats
    
    Set colModules = New Collection
    Set colProcedures = New Collection
    Set colControls = New Collection
    Set colVariables = New Collection
    
    With cdgMain
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        .Filter = "Visual Basic Projects |*.vbp|"
        .DialogTitle = "Project Analyzer"
        .ShowOpen
        strProjectFileName = .FileTitle
        strProjectPath = Left(.FileName, InStr(1, .FileName, strProjectFileName) - 1)
        
        Me.Caption = "Common Sense Software Project Analyzer   [" & strProjectFileName & "]"
        
        LoadProject
        
        txtCode.Text = ""
        
    End With
    
    If trvView.Nodes.Count > 0 Then trvView.Nodes.Item(1).Selected = True
    
    Exit Sub
    
EvalErr:
    If Err.Number <> 32755 Then MsgBox Err.Number & vbCr & Err.Description
    Unload Me
End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 2) And (Len(txtCode.Text) > 0) Then
        mnuJump.Visible = False
        mnuPrint.Visible = True
        Me.PopupMenu mnuPopup
    End If
End Sub

Public Sub MapProcedure(intProcID As Integer)
    Dim strParced() As String
    Dim intLooper As Integer
    Dim intLooper2 As Integer
    Dim strTmpProcName As String
    Dim strTmpProcParent As String
    Dim blnProcResolved As Boolean
    Dim intLocalProcID As Integer
    Dim intStart As Integer
    Dim intEnd As String
    
    lstProcMapping(0).ListItems.Clear
    lstProcMapping(1).ListItems.Clear
    
    strParced() = Split(UCase(colProcedures(intProcID).ProcCode), vbCrLf)
    
    'procs that I call
    
    For intCounter = 1 To UBound(strParced) - 1
        'for each line of code within the
        'procedure's life, excluding header and footer
        blnProcResolved = False
        
        For intLooper = 0 To colModules(colProcedures(intProcID).ProcParentID).ProcedureCount - 1
            'First, check to see if procedure is local
                        
            intLocalProcID = colModules(colProcedures(intProcID).ProcParentID).GetProcIndex(intLooper)
            strTmpProcName = UCase(colProcedures(intLocalProcID).ProcName)
            strTmpProcParent = UCase(colProcedures(intLocalProcID).ProcParent)
            
            If ((InStr(1, strParced(intCounter) & " ", strTmpProcName & " ") > 0) Or (InStr(1, strParced(intCounter), strTmpProcName & "(") > 0)) And ((InStr(1, strParced(intCounter), " " & strTmpProcName) > 0) Or (InStr(1, strParced(intCounter), strTmpProcName) = 1)) Then
                
                'A procedure exists locally with the same name,
                'let's see if it's the one that's referenced
                
                If InStr(1, strParced(intCounter), "." & strTmpProcName) > 0 Then
                    'there's a leading . here...either we have a redundant
                    'reference to a local public procedure, or a call is made to a
                    'different procedure...let's see which it is
                    
                    If (InStr(1, strParced(intCounter), strTmpProcParent & "." & strTmpProcName) > 0) Or (InStr(1, strParced(intCounter), "ME." & strTmpProcName) > 0) Then
                        'the procedure is local, but the proc is using the local
                        'module's reference...no need for that, let's report it
                        If UCase(colProcedures(intLocalProcID).ProcScope) = "PRIVATE" Then
                            AddProcToCallList 0, intLocalProcID, "Rdnt Ref./Out of Scope"
                            blnProcResolved = True
                            Exit For
                        Else
                            AddProcToCallList 0, intLocalProcID, "Redundant Reference"
                            blnProcResolved = True
                            Exit For
                        End If
                    ElseIf (InStr(1, strParced(intCounter), " ." & strTmpProcName) > 0) Or (InStr(1, strParced(intCounter), "." & strTmpProcName) = 1) Then
                        'looks like we have a With block here,
                        'let's see if the local proc is the parent
                        For intLooper2 = intCounter To 1 Step -1
                            If InStr(1, UCase(strParced(intLooper2)), "WITH ") > 0 Then
                                'we found the with block, now let's see
                                'who this procedure's parent really is
                                
                                intStart = InStr(1, UCase(strParced(intLooper2)), "WITH") + 5
                                If InStr(intStart, strParced(intLooper2), "'") > 0 Then
                                    intEnd = InStr(intStart, strParced(intLooper2), "'") - intStart
                                Else
                                    intEnd = 50
                                End If
                                
                                If (Mid(strParced(intLooper2), intStart, 2) = "ME") Or (Mid(strParced(intLooper2), intStart, intEnd) = strTmpProcParent) Then
                                    'the local procedure is the parent
                                    If UCase(colProcedures(intLocalProcID).ProcScope) = "PRIVATE" Then
                                        AddProcToCallList 0, intLocalProcID, "Rdnt Ref./Out of Scope"
                                        blnProcResolved = True
                                        Exit For
                                    Else
                                        AddProcToCallList 0, intLocalProcID, "Redundant Reference"
                                        blnProcResolved = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next intLooper2
                        If blnProcResolved Then Exit For
                    End If
                    
                Else
                    'everything's ok, this is the proc used
                    'so let's add it to the list
                    
                    AddProcToCallList 0, intLocalProcID
                    blnProcResolved = True
                    Exit For
                End If
            End If
        Next intLooper
        If Not blnProcResolved Then
            'not a local proc, must be
            'a public one
            
            For intLooper = 1 To colProcedures.Count - 1
                If intLooper <> intProcID Then
                    
                    strTmpProcName = UCase(colProcedures(intLooper).ProcName)
                    strTmpProcParent = UCase(colProcedures(intLooper).ProcParent)
            
                    If ((InStr(1, strParced(intCounter) & " ", strTmpProcName & " ") > 0) Or (InStr(1, strParced(intCounter), strTmpProcName & "(") > 0)) And ((InStr(1, strParced(intCounter), " " & strTmpProcName) > 0) Or (InStr(1, strParced(intCounter), strTmpProcName) = 1)) Then
                    
                        'A procedure exists outside with the same name,
                        'let's see if it's the one that's referenced
                        
                        If (InStr(1, strParced(intCounter), "." & strTmpProcName) > 0) Then
                            'looks like we are referencing using the . method
                            
                            If (InStr(1, strParced(intCounter), strTmpProcParent & "." & strTmpProcName) > 0) Then
                                'ok, we've identified the parent,
                                'now let's see if the procedure
                                'is private
                                
                                If UCase(colProcedures(intLooper).ProcScope) = "PRIVATE" Then
                                    'oops, it's private, so this baby is
                                    'out of scope
                                    AddProcToCallList 0, intLooper, "Out of Scope"
                                    blnProcResolved = True
                                    Exit For
                                Else
                                    AddProcToCallList 0, intLooper
                                    blnProcResolved = True
                                    Exit For
                                End If
                                
                            ElseIf (InStr(1, strParced(intCounter), " ." & strTmpProcName) > 0) Or (InStr(1, strParced(intCounter), "." & strTmpProcName) = 1) Then
                                'looks like we have a With block here,
                                'let's see if the current proc is the parent
                                
                                For intLooper2 = intCounter To 1 Step -1
                                    If InStr(1, UCase(strParced(intLooper2)), "WITH") > 0 Then
                                        'we found the with block, now let's see
                                        'who this procedure's parent really is
                                        
                                        'We do this to ensure we don't include any comments in our search
                                        intStart = InStr(1, UCase(strParced(intLooper2)), "WITH") + 5
                                        If InStr(intStart, strParced(intLooper2), "'") > 0 Then
                                            intEnd = InStr(intStart, strParced(intLooper2), "'") - intStart
                                        Else
                                            intEnd = 50
                                        End If
                                        
                                        If (Mid(strParced(intLooper2), intStart, intEnd) = "ME") Or (Mid(strParced(intLooper2), intStart, intEnd) = strTmpProcParent) Then
                                            'the local procedure is the parent
                                            If UCase(colProcedures(intLooper).ProcScope) = "PRIVATE" Then
                                                'oopsies, it's out of scope
                                                AddProcToCallList 0, intLooper, "Out of Scope"
                                                blnProcResolved = True
                                                Exit For
                                            Else
                                                AddProcToCallList 0, intLooper
                                                blnProcResolved = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next intLooper2
                                If blnProcResolved Then Exit For
                            End If
                        Else
                            'the procedure must be stored in a module
                            'and must be public, or else...
                            blnProcResolved = False
                            For intLooper2 = 1 To colProcedures.Count
                                strTmpProcName = UCase(colProcedures(intLooper2).ProcName)
                                strTmpProcParent = UCase(colProcedures(intLooper2).ProcParent)
                            
                                If ((InStr(1, strParced(intCounter) & " ", strTmpProcName & " ") > 0) Or (InStr(1, strParced(intCounter), strTmpProcName & "(") > 0)) And (colModules(colProcedures(intLooper2).ProcParentID).ModType = "Module") And (colProcedures(intLooper2).ProcScope <> "Private") Then
                                    'ok, we have a match in name, a bas module and
                                    'a public procedure.  This is the one that is
                                    'being called, so let's add it
                                    AddProcToCallList 0, intLooper
                                    blnProcResolved = True
                                    Exit For
                                    
                                End If
                            Next intLooper2
                            
                            If Not blnProcResolved Then
                                'ok, we couldn't find a procedure that
                                'matches all required criteria, so we
                                'assume that the procedure is out of scope
                                AddProcToCallList 0, intLooper, "Out of Scope"
                                blnProcResolved = True
                                Exit For
                            End If
                            
                        End If
                    End If
                    
                    
                End If
            Next intLooper
            
        End If
        
        
    Next intCounter

    Erase strParced
    
    'procs that call me
'
'    For intCounter = 1 To colProcedures.Count
'        strTmpProcName = UCase(colProcedures(intProcID).ProcName)
'        strTmpProcParent = UCase(colProcedures(intProcID).ProcParent)
'        blnProcResolved = False
'
'        strParced() = Split(UCase(colProcedures(intCounter).ProcCode), vbCrLf)
'
'        For intLooper = 1 To UBound(strParced) - 1
'            'skipping first and last lines of code
'
'            If (InStr(1, strParced(intLooper) & " ", strTmpProcName & " ") > 0) Or (InStr(1, strParced(intLooper), strTmpProcName & "(") > 0) Then
'                'current procedure contains a possible
'                'reference to me
'
'                If (InStr(1, strParced(intLooper), "." & strTmpProcName) > 0) Then
'                    'there's a . preceeding the procedure
'                    'let's see who the parent is
'
'                    If (InStr(1, strParced(intLooper), " ." & strTmpProcName) > 0) Or (InStr(1, strParced(intLooper), "." & strTmpProcName) = 1) Then
'                        'looks like we have a With block here,
'                        'let's see if the current proc is the parent
'
'                        For intLooper2 = intLooper To 1 Step -1
'                            If InStr(1, UCase(strParced(intLooper2)), "WITH") > 0 Then
'                                'we found the with block, now let's see
'                                'who this procedure's parent really is
'
'
'
'                                If ((Mid(strParced(intLooper2), InStr(1, UCase(strParced(intLooper2)), "WITH") + 5) = "ME") And (colProcedures(intCounter).ProcParent = strTmpProcParent)) Or (Mid(strParced(intLooper2), InStr(1, UCase(strParced(intLooper2)), "WITH") + 5) = strTmpProcParent) Then
'                                    'the local procedure is the parent
'                                    If UCase(colProcedures(intLooper).ProcScope) = "PRIVATE" Then
'                                        'oopsies, it's out of scope
'                                        AddProcToCallList 0, intLooper, "Out of Scope"
'                                        blnProcResolved = True
'                                        Exit For
'                                    Else
'                                        AddProcToCallList 0, intLooper
'                                        blnProcResolved = True
'                                        Exit For
'                                    End If
'                                End If
'                            End If
'                        Next intLooper2
'                        If blnProcResolved Then Exit For
'                    End If
'
'
'                End If
'            End If
'
'        Next intLooper
'
'    Next intCounter
'
End Sub

Private Sub PopulateThreats()
    Dim strParced() As String
    Dim strData As String
    Dim intLooper As Integer
    Dim intLooper2 As Integer
    
    frmThreats.lstThreat.ListItems.Clear
    
    For intCounter = 1 To 5
        strData = Replace(GetINIData("ThreatWords", "DEFCON" & Trim(intCounter)), "|", "")
        strParced = Split(strData, ",")
        For intLooper = 0 To UBound(strParced)
            For intLooper2 = 1 To colProcedures.Count
                If InStr(1, UCase(colProcedures(intLooper2).ProcCode), UCase(strParced(intLooper))) > 0 Then
                    With frmThreats.lstThreat
                        .ListItems.Add , "P" & Trim(intLooper2) & "X" & .ListItems.Count, intCounter
                        .ListItems(.ListItems.Count).SubItems(1) = strParced(intLooper)
                        .ListItems(.ListItems.Count).SubItems(2) = colProcedures(intLooper2).ProcParent & "." & colProcedures(intLooper2).ProcName
                    End With
                End If
            Next intLooper2
        Next intLooper
    Next intCounter
End Sub

Private Sub AddProcToCallList(intListViewID As Integer, intProcedureID As Integer, Optional strErrors As String = "None")
    With lstProcMapping(intListViewID)
        .ListItems.Add , "P" & intProcedureID & "X" & .ListItems.Count, colProcedures(intProcedureID).ProcParent
        .ListItems(.ListItems.Count).SubItems(1) = colProcedures(intProcedureID).ProcName
        .ListItems(.ListItems.Count).SubItems(2) = strErrors
    End With
End Sub
