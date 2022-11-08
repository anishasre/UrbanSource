VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchSEPanchayatAccountHeads 
   Caption         =   "Panchayat Account Heads"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   FillColor       =   &H80000006&
   ForeColor       =   &H00000000&
   Icon            =   "frmSearchSEPanchayatAccountHeads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   6930
      Top             =   5355
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdSearchPAccountHeads 
      Caption         =   "..."
      Height          =   315
      Left            =   6480
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   5085
      Width           =   375
   End
   Begin VB.TextBox txtPSearchKey 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1530
      TabIndex        =   0
      Top             =   5085
      Width           =   4905
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Width           =   8070
      _cx             =   14235
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
      GridColor       =   -2147483633
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
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchSEPanchayatAccountHeads.frx":1CCA
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   5130
      Width           =   1035
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Account Heads"
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
      TabIndex        =   2
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "frmSearchSEPanchayatAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRow            As Integer
Dim mSelectedRow    As Integer
Public intModeOfTransaction As Integer '   1= Receipt ; 2 = payment ;3=debit Receipt; 4=debit Payment
    Private Sub cmdSearchPAccountHeads_Click()
       Call fillgrid
    End Sub
    Private Sub Form_Load()
        vsGrid.SelectionMode = flexSelectionByRow
        WindowsXPC1.InitSubClassing
        Call fillgrid
        txtPSearchKey.Enabled = True
       
    End Sub
    Private Sub fillgrid()
        Dim mSql As String
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mRowCnt As Integer
        Dim mLoop As Integer
        Dim Rec     As New ADODB.Recordset
        Dim mArrIN As Variant
        Dim mPSearch  As Variant
        Dim mCnt        As Integer
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya

        If intModeOfTransaction = 1 Or intModeOfTransaction = 2 Or intModeOfTransaction = 3 Or intModeOfTransaction = 4 Then
        
        mSql = " SELECT  HeadID,HeadCode,Head, Level, tnyTypeID  FROM ("
        mSql = mSql + " Select intSEHeadID HeadID,vchSEHeadCode HeadCode,vchSEHead Head, 0 level, tnyTypeID   FROM faSEAccountHeads"
        mSql = mSql + " Union All"
        mSql = mSql + " SELECT  HeadID,HeadCode,Head, Level, A.tnyTypeID"
        mSql = mSql + " FROM ("
        mSql = mSql + " Select intSEMajorAccountHeadID HeadID,vchSEMajorAccountHeadCode HeadCode,vchSEMajorAccountHead Head, 6 level, tinType tnyTypeID FROM faSEMajorAccountHeads"
        mSql = mSql + " Union All"
        mSql = mSql + " Select intSESubMajorAccountHeadID HeadID,vchSESubMajorAccountHeadCode HeadCode,vchSESubMajorAccountHead Head, 5 level, tinType tnyTypeID  FROM faSESubMajorAccountHeads"
        mSql = mSql + " Union All"
        mSql = mSql + " Select intSEMinorAccountHeadID HeadID,vchSEMinorAccountHeadCode HeadCode,vchSEMinorAccountHead Head,4 level, tinType tnyTypeID  From faSEMinorAccountHeads"
        mSql = mSql + " Union All"
        mSql = mSql + " Select intSEDetailedAccountHeadID HeadID,vchSEDetailedAccountHeadCode HeadCode,vchSEDetailedAccountHead Head,2 level, tinType tnyTypeID  From faSEDetailedAccountHeads"
        mSql = mSql + " Union All"
        mSql = mSql + " Select intSESubAccountHeadID HeadID,vchSESubAccountHeadCode HeadCode,vchSESubAccountHead Head,3 level, tinType tnyTypeID  From  faSESubAccountHeads"
        mSql = mSql + " Union All"
        mSql = mSql + " Select intSEObjectAccountHeadID HeadID,vchSEObjectAccountHeadCode HeadCode,vchSEObjectAccountHead Head,1 level, tinType tnyTypeID  From faSEOBJAccountHeads"
        mSql = mSql + " ) A Full Outer  JOIN faSEAccountHeads ON faSEAccountHeads.vchSEHeadCode = A.HeadCode"
        mSql = mSql + " Where (vchSEHeadCode Is Null Or HeadCode Is Null)"
        mSql = mSql + " ) B"
 
        If intModeOfTransaction = 1 Then 'Receipts Part1
             mSql = mSql + " Where tnyTypeID = 1"
        ElseIf intModeOfTransaction = 2 Then 'Receipts Part2
             mSql = mSql + " Where tnyTypeID = 2"
        ElseIf intModeOfTransaction = 3 Then 'Payments Part1
            mSql = mSql + " Where tnyTypeID = 3"
        ElseIf intModeOfTransaction = 4 Then 'Payments Part1
            mSql = mSql + " Where tnyTypeID = 4"
        End If
        If IsNumeric(txtPSearchKey.Text) Then
           mSql = mSql + " and HeadCode like '" & Trim(txtPSearchKey.Text) & "%'"
        Else
            mSql = mSql + " and Head like '%" & Trim(txtPSearchKey.Text) & "%'"
        End If
        mSql = mSql + " ORDER BY HeadCode, Level Desc"
    ElseIf intModeOfTransaction = 5 Then
        
        mSql = " SELECT   vchSEHeadCode HeadCode, vchSEHead Head,0 Level, intSEHeadID HeadID FROM faSEAccountHeads WHERE tnyStatusID =1"
         If IsNumeric(txtPSearchKey.Text) Then
           mSql = mSql + " and vchSEHeadCode like '" & Trim(txtPSearchKey.Text) & "%'"
        Else
            mSql = mSql + " and vchSEHead like '%" & Trim(txtPSearchKey.Text) & "%'"
        End If
    End If
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        vsGrid.Clear 1, 1
        mRowCnt = 0
        vsGrid.Rows = 1
        While Not (Rec.EOF)
         
            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!HeadCode), "", Rec!HeadCode)
            vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!Head), "", Rec!Head)
            vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!Level), "", Rec!Level)
            If val(vsGrid.TextMatrix(mRowCnt, 2)) = 6 Then
                vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 3) = &H808080
            ElseIf val(vsGrid.TextMatrix(mRowCnt, 2)) = 5 Then
                vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 3) = &HC0C0C0
            ElseIf val(vsGrid.TextMatrix(mRowCnt, 2)) = 4 Then
                vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 3) = &HE0E0E0
            ElseIf val(vsGrid.TextMatrix(mRowCnt, 2)) = 3 Then
                vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 3) = &H8080FF
            ElseIf val(vsGrid.TextMatrix(mRowCnt, 2)) = 2 Then
                vsGrid.Cell(flexcpBackColor, mRowCnt, 1, mRowCnt, 3) = &HC0C0FF
'            ElseIf val(vsGrid.TextMatrix(mRowCnt, 2)) = 5 Then
'                vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 3) = &HC0E0FF
            End If
            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!HeadID), "", Rec!HeadID)
            
            vsGrid.Rows = vsGrid.Rows + 1
            mRowCnt = mRowCnt + 1
            Rec.MoveNext
        Wend
        Rec.Close
            For mCnt = 1 To vsGrid.Rows - 1
                
            Next
    End Sub
    Private Sub vsGrid_DblClick()
    If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
        If vsGrid.TextMatrix(vsGrid.Row, 2) = 0 Then
            vsGrid.Editable = flexEDKbdMouse
            gbSearchStr = ""
            gbSearchID = -1
            gbSearchCode = ""
            If vsGrid.Row <> -1 Then
                gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 1))
                gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 3))
                gbSearchCode = Trim(vsGrid.TextMatrix(vsGrid.Row, 0))
                'Me.Hide
                Unload Me
            End If
        Else
            vsGrid.Editable = flexEDNone
        End If
     End If
    End Sub
   Private Sub vsGrid_Click()
   
    If vsGrid.TextMatrix(vsGrid.Row, 2) = 0 Then
        vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 1) = &H80000005
        vsGrid.Cell(flexcpForeColor, vsGrid.Row, 0, , 1) = vbBlack
    End If
'        If Not (mRow > vsGrid.Rows - 1) Then
'            If vsGrid.TextMatrix(mRow, 4) <> "" Then
'                vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
'                vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
'            ElseIf vsGrid.TextMatrix(mRow, 2) = 6 Then
'                vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H80000005
'                vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
'            End If
'            mRow = vsGrid.Row
'            vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
'            vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
'        End If
    End Sub
    Public Property Let ModeOfTransaction(mData As Integer)
        intModeOfTransaction = mData
    End Property
    Private Function CheckModeOfTransaction() As Boolean
'       On Error GoTo Err:
'            If intModeOfTransaction = 1 Then 'Receipt Mode
'                Call fillgridR
'                CheckModeOfTransaction = True
'                Exit Function
'             ElseIf intModeOfTransaction = 2 Then
'                Call fillgridP
'                CheckModeOfTransaction = True
'                Exit Function
'            End If
'            CheckModeOfTransaction = False
'       Exit Function
'Err:
'       MsgBox (Error$)
    End Function
   Private Sub txtPSearchKey_Change()
        Dim mCount      As Long
        Dim mLength     As Long
        Dim mStr1       As String
        Dim mStr2       As String

        mLength = Len(txtPSearchKey.Text)
        mStr2 = LCase(txtPSearchKey.Text)
        If mRow < vsGrid.Rows Then
            If vsGrid.TextMatrix(mRow, 4) = "" Then
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
                    mRow = mCount
                    Exit For
                End If
            Next
        End If
        If Trim(txtPSearchKey.Text) = "" Then
            mRow = 0
        End If
    End Sub
    Private Sub txtPSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
             '38 = Up Arrow ,  40 = Down Arrow
        If Not (mRow > vsGrid.Rows - 1) Then
            If (KeyCode = 40) Then
                If vsGrid.TextMatrix(mRow, 4) <> "" Then
                    vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
                    vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                    mRow = mRow + 1
                    vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                    vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                    Call vsGrid.ShowCell(mRow, 0)
                Else
                    If mRow <> vsGrid.Rows - 1 Then
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = vbWhite
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                        mRow = mRow + 1
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                        Call vsGrid.ShowCell(mRow, 0)
                    End If
                End If
            ElseIf (KeyCode = 38) Then
                If mRow <> 0 Then
                    If vsGrid.TextMatrix(mRow, 4) <> "" Then
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &HC0FFC0
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                        mRow = mRow - 1
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                        Call vsGrid.ShowCell(mRow, 0)
                    Else
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = vbWhite
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbBlack
                        mRow = mRow - 1
                        vsGrid.Cell(flexcpBackColor, mRow, 0, , 1) = &H8000000D
                        vsGrid.Cell(flexcpForeColor, mRow, 0, , 1) = vbWhite
                        Call vsGrid.ShowCell(mRow, 0)
                    End If
                End If
            ElseIf (KeyCode = 13) Then
                gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 1))
                gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 3))
                gbSearchCode = Trim(vsGrid.TextMatrix(vsGrid.Row, 0))
                Unload Me
            End If
       End If
    End Sub
' Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'        Call txtPSearchKey_KeyDown(KeyCode, Shift)
'    End Sub
