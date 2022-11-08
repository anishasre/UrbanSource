VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "winxpc.ocx"
Begin VB.Form frmScheduleReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Schedule"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSchedules 
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.OptionButton OptMinorMnrMjrDet 
      Caption         =   "Detailed"
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton OptMinorMnrMjrDet 
      Caption         =   "Minor"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.OptionButton OptMinorMnrMjrDet 
      Caption         =   "Major"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   7560
      TabIndex        =   14
      Top             =   5280
      Width           =   1245
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4620
      TabIndex        =   12
      Top             =   3075
      Width           =   450
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4620
      TabIndex        =   9
      Top             =   2475
      Width           =   450
   End
   Begin VB.ListBox lstAccountHeads 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   30
      TabIndex        =   8
      Top             =   1170
      Width           =   4485
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   120
      Width           =   2085
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   4740
      Width           =   3825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   5640
      TabIndex        =   13
      Top             =   5280
      Width           =   1245
   End
   Begin VB.CommandButton cmdSearchSchedules 
      Caption         =   "..."
      Height          =   285
      Left            =   3390
      TabIndex        =   0
      Top             =   120
      Width           =   405
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3600
      Top             =   6330
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3885
      Left            =   5160
      TabIndex        =   11
      Top             =   1200
      Width           =   4575
      _cx             =   8070
      _cy             =   6853
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmScheduleReports.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "       Selected Heads"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4875
      TabIndex        =   18
      Top             =   840
      Width           =   4845
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "  Account Heads"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   -30
      TabIndex        =   17
      Top             =   840
      Width           =   4890
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   4440
      TabIndex        =   16
      Top             =   165
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Schedule Title"
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   165
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Search"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   4785
      Width           =   510
   End
End
Attribute VB_Name = "frmScheduleReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mCount As Integer
    Dim mLevel As Integer

Public Sub ClearFields()
    
    txtDescription.Text = ""
    txtSearch.Text = ""
    'txtTitle.Text = ""
    lstAccountHeads.Clear
    lstSchedules.Clear
    vsGrid.Clear
    vsGrid.TextMatrix(0, 0) = "Account Heads"
    OptMinorMnrMjrDet(0).Value = False
    OptMinorMnrMjrDet(1).Value = False
    OptMinorMnrMjrDet(2).Value = False
End Sub
Public Sub AssignSerial()
    
    For i = 1 To vsGrid.Rows - 1
        If vsGrid.TextMatrix(i, 0) <> "" Then vsGrid.TextMatrix(i, 2) = i
    Next
End Sub
Private Sub cmdAdd_Click()

    If lstAccountHeads.Selected(lstAccountHeads.ListIndex) = True Then
        vsGrid.TextMatrix(mCount, 0) = lstAccountHeads.List(lstAccountHeads.ListIndex)
        vsGrid.TextMatrix(mCount, 1) = lstAccountHeads.ItemData(lstAccountHeads.ListIndex)
        vsGrid.TextMatrix(mCount, 2) = mCount
        vsGrid.TextMatrix(mCount, 3) = mLevel
        mCount = mCount + 1
        OptMinorMnrMjrDet_Click (0)
    End If
End Sub

Private Sub cmdRemove_Click()
    If vsGrid.IsSelected(vsGrid.Row) Then
        vsGrid.RemoveItem (vsGrid.Row)
        mCount = mCount - 1
        OptMinorMnrMjrDet_Click (0)
        AssignSerial
    End If
End Sub

Private Sub cmdSave_Click()
    Dim mConn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim arrIn(6) As Variant
    If txtDescription.Text = "" Then MsgBox "Description Cannot be Left Blank", vbInformation, "Saankhya": Exit Sub
    If txtTitle.Text = "" Then MsgBox "Schedule Title Cannot be leftBlank", vbInformation, "Saankhya": Exit Sub
    If vsGrid.TextMatrix(1, 0) = "" Then MsgBox "Select at Least One Account Head", vbInformation, "Saankhya": Exit Sub
    arrIn(0) = txtTitle.Text
    arrIn(1) = txtDescription.Text
    arrIn(2) = ""
    For i = 1 To vsGrid.Rows
     If vsGrid.TextMatrix(i, 1) <> "" Then
        arrIn(3) = vsGrid.TextMatrix(i, 1)
        arrIn(4) = vsGrid.TextMatrix(i, 2)
        arrIn(5) = mID(vsGrid.TextMatrix(i, 0), 1, 9)
        arrIn(6) = vsGrid.TextMatrix(i, 3)
        objdb.ExecuteSP "spSaveSchedules", arrIn, , , mConn, adCmdStoredProc
     Else
        Exit For
    End If
    Next
    ClearFields
End Sub

Private Sub cmdSearchSchedules_Click()
    Dim rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    Dim objdb As New clsDB
    lstSchedules.Left = 3360
    lstSchedules.Height = 2000
    lstSchedules.Visible = True
    mSQL = "Select * from faScheduleReports"
    mCnn.ConnectionString = objdb.GetConnectionString(enuSourceString.Saankhya)
    mCnn.Open
    rec.Open mSQL, mCnn
    i = 0
    While Not rec.EOF
        lstSchedules.AddItem rec.Fields(1)
        lstSchedules.ItemData(i) = rec.Fields(0)
        rec.MoveNext
        i = i + 1
    Wend
    lstSchedules.SetFocus
End Sub




Private Sub Command1_Click()
    
    Unload Me
End Sub

Private Sub lstAccountHeads_DblClick()
    
    Call cmdAdd_Click
End Sub

Private Sub lstSchedules_DblClick()
    
    txtTitle.Text = lstSchedules.Text
    lstSchedules.Visible = False
    Call txtTitle_LostFocus
    cmdSearchSchedules.SetFocus
End Sub

Private Sub lstSchedules_LostFocus()
    
    lstSchedules.Visible = False
End Sub

Private Sub OptMinorMnrMjrDet_Click(Index As Integer)

    Dim mSQL As String
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim rec As New ADODB.Recordset
    lstAccountHeads.Clear
    mSQL = "Select * from faMinorAccountHeads where vchMinorAccountHeadCode not in('0',"
    If OptMinorMnrMjrDet(0).Value = True Then
        mSQL = "Select * from faMajorAccountHeads where vchMajorAccountHeadCode not in('0',"
        mLevel = 1
    End If
    If OptMinorMnrMjrDet(1).Value = True Then
        mSQL = "Select * from faMinorAccountHeads where vchMinorAccountHeadCode not in('0',"
        mLevel = 2
    End If
    If OptMinorMnrMjrDet(2).Value = True Then
        mSQL = "Select * from faAccountHeads where vchAccountHeadCode not in('0',"
        mLevel = 3
    End If
    For i = 1 To vsGrid.Rows - 1
        If vsGrid.TextMatrix(i, 0) <> "" Then
            mSQL = mSQL + "'" & mID(vsGrid.TextMatrix(i, 0), 1, 9) + "',"
        Else
            Exit For
        End If
    Next
    mSQL = Left(mSQL, Len(mSQL) - 1) + ")"
    mCnn.ConnectionString = objdb.GetConnectionString(enuSourceString.Saankhya)
    mCnn.Open
    rec.Open mSQL, mCnn
    While Not rec.EOF
        lstAccountHeads.AddItem rec.Fields(1) + "    " + rec.Fields(2)
        lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = rec.Fields(0)
        'lstAccountHeads.ToolTipText = rec.Fields(0)
        rec.MoveNext
    Wend
End Sub

Private Sub txtSearch_Change()
    For i = 1 To lstAccountHeads.ListCount
        If txtSearch.Text = mID(lstAccountHeads.List(i), 1, Len(txtSearch.Text)) Or txtSearch.Text = mID(lstAccountHeads.List(i), 14, Len(txtSearch.Text)) Then
            lstAccountHeads.Selected(i) = True
            Exit Sub
        End If
    Next
End Sub

Private Sub txtTitle_LostFocus()
    
    Dim rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mCon As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim mSQL As String
    Call ClearFields
    mCount = 1
    Dim objdb As New clsDB
    mCnn.ConnectionString = objdb.GetConnectionString(enuSourceString.Saankhya)
    mCnn.Open
    cmdSave.Caption = "&Save"
    rec.Open "spGetSchedules '" & Trim(txtTitle.Text) & "'", mCnn
       While Not rec.EOF
            cmdSave.Caption = "&Update"
            vsGrid.Rows = vsGrid.Rows + 1
            txtDescription.Text = rec!vchDescription
            Label6.Caption = IIf(IsNull(rec!intScheduleGroupId), "", rec!intScheduleGroupId)
            If rec!tinAccountHeadLevel = 1 Then
                mSQL = "select * from faMajorAccountHeads where intMajorAccountHeadId=" & rec!intAccountHeadID
                ElseIf rec!tinAccountHeadLevel = 2 Then
                    mSQL = "select * from faMinorAccountHeads where intMinorAccountHeadId=" & rec!intAccountHeadID
                Else
                    mSQL = "select * from faAccountHeads where intAccountHeadId=" & rec!intAccountHeadID
            End If
            mCon.ConnectionString = objdb.GetConnectionString(enuSourceString.Saankhya)
            mCon.Open
            rs.Open mSQL, mCon
            vsGrid.TextMatrix(mCount, 0) = rec!vchAccountHeadCode + "    " + rs.Fields(2)
            vsGrid.TextMatrix(mCount, 1) = rec!intAccountHeadID
            vsGrid.TextMatrix(mCount, 2) = mCount
            vsGrid.TextMatrix(mCount, 3) = rec!tinAccountHeadLevel
            mCon.Close
            rec.MoveNext
            mCount = mCount + 1
        Wend
End Sub

Private Sub vsGrid_DblClick()

    Call cmdRemove_Click
End Sub
