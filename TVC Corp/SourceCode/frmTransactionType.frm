VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmTransactionType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Transaction Type"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdvancedSearch 
      Caption         =   "Search"
      Height          =   345
      Left            =   2400
      TabIndex        =   5
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ListBox lstAccountHeads 
      Height          =   4350
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1440
      Width           =   3735
   End
   Begin VB.ListBox lstSelected 
      Height          =   1035
      Left            =   8490
      TabIndex        =   16
      Top             =   150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstTransactionType 
      BackColor       =   &H80000018&
      Height          =   450
      Left            =   6360
      TabIndex        =   15
      Top             =   5970
      Visible         =   0   'False
      Width           =   345
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3600
      Top             =   6420
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CheckBox chkSystemTransaction 
      Caption         =   "System Transaction"
      Height          =   375
      Left            =   4470
      TabIndex        =   9
      Top             =   5970
      Width           =   1755
   End
   Begin VB.ComboBox cmbGroup 
      Height          =   315
      ItemData        =   "frmTransactionType.frx":0000
      Left            =   1740
      List            =   "frmTransactionType.frx":0017
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   660
      Width           =   4455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   6030
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<"
      Height          =   345
      Left            =   3900
      TabIndex        =   8
      Top             =   3630
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">"
      Height          =   345
      Left            =   3900
      TabIndex        =   6
      Top             =   2730
      Width           =   495
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4545
      Left            =   4470
      TabIndex        =   7
      Top             =   1320
      Width           =   4185
      _cx             =   7382
      _cy             =   8017
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
      BackColorAlternate=   -2147483624
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransactionType.frx":0051
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
      Editable        =   0
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
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   30
      TabIndex        =   12
      Top             =   30
      Width           =   8715
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   7110
         TabIndex        =   11
         Top             =   420
         Width           =   915
      End
      Begin VB.TextBox txtTransactionTypeID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton cmdTransactionType 
         Caption         =   "---"
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   210
         Width           =   465
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   210
         Width           =   3765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   960
         TabIndex        =   14
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Menu mnuSub 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuSubDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmTransactionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Call AddToGrid
End Sub

Private Sub cmdAdvancedSearch_Click()
On Error GoTo Err:
    Dim mIndex As Long
    Dim mTempString As String
    frmSearchAccountHeads.Show vbModal
   
    If gbSearchStr <> "" Then
        mTempString = Trim(mID(gbSearchStr, InStr(1, gbSearchStr, " ", vbTextCompare)))
        mIndex = SendMyMessage(lstAccountHeads.hwnd, LB_FINDSTRING, -1, mTempString)
        If mIndex <> -1 Then
        
            lstAccountHeads.Selected(mIndex) = True
            Call AddToGrid
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End If
    Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub

Private Sub cmdNew_Click()
On Error GoTo Err:
    FormInitialize
    txtTransactionType.Locked = False
    Dim mCon As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    If objDB.SetConnection(mCon) Then
        Set Rec = objDB.ExecuteSP("Select isnull(max(intTransactionTypeID),0)+1 as TransTypeID from faTransactionType where intTransactionTypeID<> 9999", , , , mCon, adCmdText)
        If Not Rec.EOF Then
            txtTransactionTypeID.Text = Rec!TransTypeID
        End If
    End If
    Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub

Private Sub cmdRemove_Click()
    Call RemoveFromGrid
    Call FillAccountHeads
End Sub

Private Sub cmdSave_Click()
    Call SaveTransactionType
    Call FormInitialize
End Sub

Private Sub cmdTransactionType_Click()
On Error GoTo Err:
    lstTransactionType.Visible = True
    lstTransactionType.SetFocus
    lstTransactionType.ZOrder (0)
    PopulateList lstTransactionType, "SELECT vchTransactionType,intTransactionTypeID FROM faTransactionType ORDER BY intTransactionTypeID", , True, True, True, enuSourceString.Saankhya
    Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Load()
    WindowsXPC.InitIDESubClassing
    Call FormInitialize
    'FillAccountHeads
End Sub



Private Sub lstAccountHeads_DblClick()
    Call AddToGrid
End Sub

Private Sub lstAccountHeads_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call AddToGrid
    End If
End Sub

Private Sub lstTransactionType_DblClick()
    If lstTransactionType.ListIndex > 0 Then
        Call FillGrid
        FillAccountHeads
    End If
    Call lstTransactionType_LostFocus
End Sub

Private Sub lstTransactionType_GotFocus()
    lstTransactionType.Top = txtTransactionType.Top
    lstTransactionType.Width = txtTransactionType.Width
    lstTransactionType.Left = txtTransactionType.Left
    lstTransactionType.Height = 2500
End Sub

Private Sub lstTransactionType_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call lstTransactionType_DblClick
    End If
End Sub

Private Sub lstTransactionType_LostFocus()
    lstTransactionType.Visible = False
    Me.Refresh
End Sub
Private Sub FillAccountHeads()
On Error GoTo Err:
    lstAccountHeads.Clear
        
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mIndex As Integer
        Dim mTempString As String
        Dim mQry As String
        
        mQry = "Select vchAccountHead,intAccountHeadID,vchAccountHeadCode from faAccountHeads Order By vchAccountHead"
       
        If objDB.SetConnection(mCnn) Then
            Set Rec = objDB.ExecuteSP(mQry, , , , mCnn, adCmdText)
        End If
        While Not Rec.EOF
            mTempString = Rec!vchAccountHead
            mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTempString)
            If mIndex = -1 Then
                lstAccountHeads.AddItem Rec!vchAccountHead
                lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = Rec!intAccountHeadID
            End If
            Rec.MoveNext
        Wend
        vsGrid.AutoSize 0, vsGrid.Cols - 1, , True
        Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub

Private Sub AddToGrid()
On Error GoTo Err:
        Dim mCount As Integer
        If lstAccountHeads.ListCount > 0 Then
            For mCount = 0 To lstAccountHeads.ListCount - 1
                If lstAccountHeads.Selected(mCount) Then
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 0) = lstAccountHeads.List(mCount)
                    vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1) = lstAccountHeads.ItemData(mCount)
                    lstSelected.AddItem vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 0)
                    lstSelected.ItemData(lstSelected.NewIndex) = vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1)
                End If
            Next mCount
            Call FillAccountHeads
        End If
        Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub
Private Sub FormInitialize()
    txtTransactionType.Text = ""
    txtTransactionType.Tag = ""
    txtTransactionTypeID.Text = ""
    txtTransactionTypeID.Tag = ""
    If cmbGroup.ListCount > 0 Then cmbGroup.ListIndex = -0
    lstTransactionType.Clear
    lstAccountHeads.Clear
    lstSelected.Clear
    vsGrid.Rows = 1
    txtTransactionType.Locked = True
    txtTransactionTypeID.Locked = True
    chkSystemTransaction.Value = vbUnchecked
    FillAccountHeads
End Sub
Private Sub RemoveFromGrid()
On Error GoTo Err:
        Dim mLoop As Long
        Dim mChildLoop As Long
        Dim mIndex As Long
        Dim mTempString As String
        Dim mCount As Integer
        If vsGrid.Rows > 0 Then
            mCount = 1
            While mCount <= vsGrid.Rows - 1
                If mCount <= vsGrid.Rows - 1 Then
                    If vsGrid.IsSelected(mCount) = True Then
                        mTempString = vsGrid.Cell(flexcpText, mCount, 0)
                        mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTempString)
                        If mIndex <> -1 Then
                            lstSelected.RemoveItem (mIndex)
                        End If
                        vsGrid.RemoveItem (mCount)
                    Else
                        mCount = mCount + 1
                    End If
                End If
            Wend
        End If
        Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub

'Private Sub txtSearch_Change()
'On Error GoTo err:
'    Dim mTempString As String
'    Dim mIndex As Long
'    mTempString = Trim(txtSearch.Text)
'    If lstAccountHeads.ListCount > 0 Then
'        mIndex = SendMyMessage(lstAccountHeads.hwnd, LB_FINDSTRING, -1, mTempString)
'        If mIndex <> -1 Then
'            lstAccountHeads.ListIndex = mIndex
'        End If
'    End If
'    Exit Sub
'err:
'    MsgBox Error, vbCritical, "Saankhya"
'End Sub
Private Sub FillGrid()
On Error GoTo Err:
    Dim objDB As New clsDB
    Dim mCon As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim mCount As Integer
    vsGrid.Rows = 1
    If objDB.SetConnection(mCon) Then
        mSQL = "SELECT * FROM faTransactionType Where intTransactionTypeID=" & lstTransactionType.ItemData(lstTransactionType.ListIndex)
'        mSQL = "SELECT     faTransactionType.intGroupID, faTransactionType.intTransactionTypeID, faTransactionType.vchTransactionType, "
'        mSQL = mSQL & " faTransactionTypeChild.intAccountHeadID , faTransactionTypeChild.intOrder, faAccountHeads.vchAccountHead"
'        mSQL = mSQL & " FROM         faTransactionType LEFT OUTER JOIN"
'        mSQL = mSQL & " faTransactionTypeChild ON faTransactionType.intTransactionTypeID = faTransactionTypeChild.intTransactionTypeID"
'        mSQL = mSQL & " LEFT OUTER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faTransactionTypeChild.intAccountHeadID "
'        mSQL = mSQL & " WHERE faTransactionType.intTransactionTypeID=" & lstTransactionType.ItemData(lstTransactionType.ListIndex)
        Set Rec = objDB.ExecuteSP(mSQL, , , , mCon, adCmdText)
        If Not Rec.EOF Then
            txtTransactionType.Text = Rec!vchTransactionType
            txtTransactionTypeID.Text = Rec!intTransactionTypeID
            chkSystemTransaction.Value = IIf(IsNull(Rec!tnySystemType), vbUnchecked, Rec!tnySystemType)
            For mCount = 1 To cmbGroup.ListCount - 1
                If cmbGroup.ItemData(mCount) = Rec!intGroupID Then
                    cmbGroup.ListIndex = mCount
                    Exit For
                End If
            Next
        End If
        mSQL = "Select faTransactionTypeChild.intAccountHeadID , faTransactionTypeChild.intOrder, faAccountHeads.vchAccountHead"
        mSQL = mSQL & " From faTransactionTypeChild Inner Join faAccountHeads"
        mSQL = mSQL & " ON faTransactionTypeChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
        mSQL = mSQL & " Where faTransactionTypeChild.intTransactionTypeID=" & lstTransactionType.ItemData(lstTransactionType.ListIndex)
        Set Rec = objDB.ExecuteSP(mSQL, , , , mCon, adCmdText)
        If Not Rec.EOF Then
            
            While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                lstSelected.AddItem vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 0)
                lstSelected.ItemData(lstSelected.NewIndex) = Val(vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1))
                Rec.MoveNext
            Wend
            vsGrid.AutoSize 0, vsGrid.Cols - 1, , True
        End If
    End If
    Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub

Private Sub SaveTransactionType()
On Error GoTo Err:
    If Val(txtTransactionTypeID.Text) = 0 Or Trim(txtTransactionType.Text) = "" Then
        MsgBox "No TransactionType Selected", vbInformation, "Saankhya"
        Exit Sub
    End If
    If cmbGroup.ListIndex <= 0 Then
        MsgBox "Select Group", vbInformation, "Saankhya"
        cmbGroup.SetFocus
        Exit Sub
    End If
    Dim mCon As New ADODB.Connection
    Dim objDB As New clsDB
    Dim mVarrIn As Variant
    Dim mCount As Integer
    Dim str As String
    If cmbGroup.ItemData(cmbGroup.ListIndex) = 10 Then
        str = "R"
    ElseIf cmbGroup.ItemData(cmbGroup.ListIndex) = 20 Then
        str = "P"
    ElseIf cmbGroup.ItemData(cmbGroup.ListIndex) = 30 Then
        str = "CV"
    ElseIf cmbGroup.ItemData(cmbGroup.ListIndex) = 40 Then
        str = "JV"
    End If
    mVarrIn = Array(Val(txtTransactionTypeID.Text), _
                    Trim(txtTransactionType.Text), _
                    0, 0, 0, 0, _
                    cmbGroup.ItemData(cmbGroup.ListIndex), _
                    str, _
                    0, _
                    IIf((chkSystemTransaction.Value = vbChecked), 1, 0) _
                    )
    If objDB.SetConnection(mCon) Then
        objDB.ExecuteSP "spSaveTransactionType", mVarrIn, , , mCon, adCmdStoredProc
        mCon.Execute ("DELETE FROM faTransactionTypeChild where intTransactionTypeID=" & Val(txtTransactionTypeID.Text))
        For mCount = 1 To vsGrid.Rows - 1
            mVarrIn = Array(Val(txtTransactionTypeID.Text), _
                            mCount, _
                            Val(vsGrid.TextMatrix(mCount, 1)), _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null _
                        )
            objDB.ExecuteSP "spSaveTransactionTypeChild", mVarrIn, , , mCon, adCmdStoredProc
        Next
        MsgBox "Transaction Type Saved", vbInformation, "Saankhya"
    End If
    Exit Sub
Err:
    MsgBox Error, vbCritical, "Saankhya"
End Sub


Private Sub vsGrid_DblClick()
    Call cmdRemove_Click
End Sub

Private Sub vsGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdRemove_Click
    End If
End Sub
