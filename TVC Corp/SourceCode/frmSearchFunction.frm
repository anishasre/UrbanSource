VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchFunction 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Functions"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4695
      Left            =   15
      TabIndex        =   2
      Top             =   345
      Width           =   6225
      _cx             =   10980
      _cy             =   8281
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   -2147483632
      TreeColor       =   16777215
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchFunction.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3480
      Top             =   5400
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtSearchKey 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   5130
      Width           =   4905
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   315
      Left            =   5730
      TabIndex        =   1
      Top             =   5130
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "  Functions"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Function"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   5190
      Width           =   615
   End
End
Attribute VB_Name = "frmSearchFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'    Private Sub cmdSearch_Click()
'        On Error GoTo Err:
''            Dim mQuery1 As String
''            Dim mQuery2 As String
''
''            If txtSearchKey.Text = "" Then
''                mQuery1 = "Select (vchFunctionCode + '  ' + vchFunction) as FunctionHead, intFunctionID  From faFunctions"
''                mQuery2 = "Select (vchFunction) as FunctionHead, intFunctionID  From faFunctions"
''
''                PopulateList lstFunction, mQuery1, , True, True, True
''                PopulateList lstHead, mQuery2, , True, True, True
''
''                Exit Sub
''            End If
''
''            If IsNumeric(txtSearchKey.Text) Then
''                mQuery1 = "Select (vchFunctionCode + '  ' + vchFunction) as FunctionHead, intFunctionID  From faFunctions Where vchFunctionCode Like '" & Val(txtSearchKey.Text) & "%'"
''                mQuery2 = "Select (vchFunction) as FunctionHead, intFunctionID  From faFunctions Where vchFunctionCode Like '" & Val(txtSearchKey.Text) & "%'"
''
''            Else
''                mQuery1 = "Select (vchFunctionCode + '  ' + vchFunction) as FunctionHead, intFunctionID  From faFunctions Where vchFunction Like '%" & Trim(txtSearchKey.Text) & "%'"
''                mQuery2 = "Select (vchFunction) as FunctionHead, intFunctionID  From faFunctions Where vchFunction Like '%" & Trim(txtSearchKey.Text) & "%'"
''            End If
''
''            PopulateList lstFunction, mQuery1, , True, True, True
''            PopulateList lstHead, mQuery2, , True, True, True
''
''        Exit Sub
'Err:
'        MsgBox (Error$)
'    End Sub
'
'    Private Sub Form_Load()
'        On Error GoTo Err:
'
''            mQuery1 = "Select (vchFunctionCode + '  ' + vchFunction) as FunctionHead, intFunctionID  From faFunctions"
''            mQuery2 = "Select (vchFunction) as FunctionHead, intFunctionID  From faFunctions"
''
''            PopulateList lstFunction, mQuery1, , True, True, True
''            PopulateList lstHead, mQuery2, , True, True, True
'        Exit Sub
'Err:
'        MsgBox (Error$)
'    End Sub
'
'    Private Sub txtSearchKey_Change()
''            Dim mIndex As Long
''            Dim mStr As String
''            mStr = txtSearchKey.Text
''            If IsNumeric(mStr) Then
''                mIndex = SendMyMessage(lstFunction.hwnd, LB_FINDSTRING, -1, ByVal mStr)
''            Else
''                mIndex = SendMyMessage(lstHead.hwnd, LB_FINDSTRING, -1, ByVal mStr)
''            End If
''            If mIndex > -1 Then
''                lstFunction.ListIndex = mIndex
''            End If
'    End Sub
'    Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
'        '38 Up Arrow
''        If (KeyCode = 38 Or KeyCode = 40) Then
''            If KeyCode = 38 And lstFunction.ListIndex > 0 Then
''                lstFunction.ListIndex = lstFunction.ListIndex - 1
''            End If
''            '40 = Down Arrow
''            If KeyCode = 40 And lstFunction.ListIndex < (lstFunction.ListCount - 1) Then
''                lstFunction.ListIndex = lstFunction.ListIndex + 1
''                'Debug.Print lstAccountHeads.ListCount - 1, lstAccountHeads.ListIndex
''            End If
''        End If
'    End Sub
'
'    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''        If KeyCode = vbKeyEscape Then
''            Unload Me
''        ElseIf KeyCode = 13 Then
''            If lstFunction.ListIndex > -1 Then
''                gbSearchStr = lstFunction.Text
''                gbSearchID = lstFunction.ItemData(lstFunction.ListIndex)
''                Unload Me
''            End If
''        End If
'    End Sub
'
''    Private Sub lstFunction_DblClick()
''        Call Form_KeyDown(13, 0)
''    End Sub
'
Option Explicit
Dim mRow            As Integer
Dim mSelectedRow    As Integer
    Private Sub FillvsGrid(Rec As ADODB.Recordset)
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        While Not Rec.EOF
            vsGrid.AddItem ""
            If Not (IsNull(Rec!vchFunctionCode)) Then
                If Right(Rec!vchFunctionCode, 6) = "000000" Then
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HC0FFC0
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = 1
                    'Commented above line By Sinoj
                End If
            End If
            vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode)
            vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            Rec.MoveNext
        Wend
    End Sub

    Private Sub cmdSearch_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mRow = 0
        If IsNumeric(txtSearchKey.Text) Then
            mSql = "Select intFunctionID, vchFunction, vchFunctionCode From faFunctions Where vchFunctionCode Like '" & val(txtSearchKey.Text) & "%'"
        Else
            mSql = "Select intFunctionID, vchFunction, vchFunctionCode From faFunctions Where vchFunction Like '%" & Trim(txtSearchKey.Text) & "%'"
        End If
        Rec.Open mSql, mCnn
        Call FillvsGrid(Rec)
        Rec.Close
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Unload Me
        ElseIf KeyCode = 13 Then
            If mRow <> 0 Then
                If (vsGrid.TextMatrix(mRow, 3) = "") Then
                    gbSearchStr = Trim(vsGrid.TextMatrix(mRow, 0)) + " " + Trim(vsGrid.TextMatrix(mRow, 1))
                    gbSearchID = val(vsGrid.TextMatrix(mRow, 2))
                    Unload Me
                End If
            End If
        End If
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb   As New clsDB
        Dim mSql    As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mRow = 0
        vsGrid.SelectionMode = flexSelectionByRow
        'vsGrid.AutoSearch = flexSearchFromCursor
        'txtSearchKey.SetFocus
        WindowsXPC1.InitSubClassing
        
        mSql = "Select intFunctionID, vchFunction, vchFunctionCode From faFunctions"
        Rec.Open mSql, mCnn
        Call FillvsGrid(Rec)
        Rec.Close
    End Sub
    
    Private Sub txtSearchKey_Change()
        Dim mCount      As Long
        Dim mLength     As Long
        Dim mStr1       As String
        Dim mStr2       As String

        mLength = Len(txtSearchKey.Text)
        mStr2 = LCase(txtSearchKey.Text)
        If mRow < vsGrid.Rows Then
            If vsGrid.TextMatrix(mRow, 3) = "" Then
                vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H80000005
                vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
            Else
                vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
                vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
            End If
        End If
        If mLength <> 0 Then
            For mCount = 1 To vsGrid.Rows - 1
                If IsNumeric(mStr2) Then
                    mStr1 = mID(LCase(vsGrid.TextMatrix(mCount, 0)), 1, mLength)
                Else
                    mStr1 = mID(LCase(vsGrid.TextMatrix(mCount, 1)), 1, mLength)
                End If
                If mStr1 = mStr2 Then
                    vsGrid.Cell(flexcpBackColor, mCount, 0, , 1) = &H8000000D
                    vsGrid.Cell(flexcpForeColor, mCount, 0, , 1) = vbWhite
                    Call vsGrid.ShowCell(mCount, 0)
                    'Call vsGrid.ShowCell(mCount, 1)
                    mRow = mCount
                    Exit For
                End If
            Next
        End If
        If Trim(txtSearchKey.Text) = "" Then
            mRow = 0
        End If
    End Sub

    Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
            '38 = Up Arrow ,  40 = Down Arrow
        If Not (mRow > vsGrid.Rows - 1) Then
            If (KeyCode = 40) Then
                If vsGrid.TextMatrix(mRow, 3) <> "" Then
                    vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
                    vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                    mRow = mRow + 1
                    vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                    vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                    Call vsGrid.ShowCell(mRow, 0)
                Else
'                    vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H80000005
                    If mRow <> vsGrid.Rows - 1 Then
'                        vsGrid.Cell(flexcpBackColor, mRow + 1, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = vbWhite
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                        mRow = mRow + 1
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                        Call vsGrid.ShowCell(mRow, 0)
                    End If
                End If
                'Call vsGrid.ShowCell(mRow, 0)
                'mRow = mRow + 1
            ElseIf (KeyCode = 38) Then
                If mRow <> 0 Then
                    If vsGrid.TextMatrix(mRow, 3) <> "" Then
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                        mRow = mRow - 1
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                        Call vsGrid.ShowCell(mRow, 0)
                    Else
'                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H80000005
'                        vsGrid.Cell(flexcpBackColor, mRow - 1, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = vbWhite
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                        mRow = mRow - 1
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                        Call vsGrid.ShowCell(mRow, 0)
                    End If
'                    Call VSGrid.ShowCell(mRow, 0)
'                    mRow = mRow - 1
                End If
            ElseIf (KeyCode = 13) Then
                gbSearchStr = Trim(vsGrid.TextMatrix(mRow, 0)) + " " + Trim(vsGrid.TextMatrix(mRow, 1))
                gbSearchID = val(vsGrid.TextMatrix(mRow, 2))
                gbSearchCode = Trim(vsGrid.TextMatrix(mRow, 0))
                Unload Me
            End If
            'Call vsGrid.ShowCell(mRow, 0)
        'Else
         '   mRow = 0
        End If
    End Sub

    Private Sub vsGrid_Click()
        If Not (mRow > vsGrid.Rows - 1) Then
            If vsGrid.TextMatrix(mRow, 3) <> "" Then
                vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
                vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
            Else
                vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H80000005
                vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
            End If
            mRow = vsGrid.Row
            vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
            vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
        End If
    End Sub

    Private Sub VSGrid_DblClick()
        gbSearchStr = ""
        gbSearchID = -1
        gbSearchCode = ""
        If vsGrid.Row <> 0 Then
            'If (VSGrid.TextMatrix(VSGrid.Row, 3) = "") Then
                gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) + " " + Trim(vsGrid.TextMatrix(vsGrid.Row, 1))
                gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 2))
                gbSearchCode = Trim(vsGrid.TextMatrix(vsGrid.Row, 0))
                Unload Me
            'End If
        End If
    End Sub
    
    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        Call txtSearchKey_KeyDown(KeyCode, Shift)
    End Sub
