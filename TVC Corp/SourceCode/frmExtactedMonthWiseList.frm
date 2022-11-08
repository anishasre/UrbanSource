VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmExtactedMonthWiseList 
   BorderStyle     =   0  'None
   Caption         =   "Extacted MonthWise List"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   Icon            =   "frmExtactedMonthWiseList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11250
      TabIndex        =   2
      Top             =   6615
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Height          =   960
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   13140
      TabIndex        =   1
      Top             =   45
      Width           =   13200
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   1020
      Width           =   14595
      _cx             =   25744
      _cy             =   9340
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   14
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmExtactedMonthWiseList.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
End
Attribute VB_Name = "frmExtactedMonthWiseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mYearID As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Left = 1850
    Me.Top = 2200
End Sub

Private Sub Form_Load()
    vsGrid.MergeRow(0) = True
    vsGrid.MergeCol(0) = True
    vsGrid.MergeCol(1) = True
    vsGrid.MergeCol(2) = True
    vsGrid.MergeCol(4) = True
    vsGrid.MergeCol(5) = True
    Call FillGrid
End Sub
Public Sub FillGrid()
    Call FillMonth
    Call VerifyStatus
End Sub
Private Sub FillMonth()
    Dim mCnt    As Integer
    Dim mCntM   As Integer
    mCnt = 2
    For mCntM = 4 To 12
        vsGrid.TextMatrix(mCnt, 0) = MonthName(mCntM)
        vsGrid.TextMatrix(mCnt, 6) = mCntM
        mCnt = mCnt + 1
    Next
    For mCntM = 1 To 3
        vsGrid.TextMatrix(mCnt, 0) = MonthName(mCntM)
        vsGrid.TextMatrix(mCnt, 6) = mCntM
        mCnt = mCnt + 1
    Next
End Sub

Private Sub vsGrid_DblClick()
     Dim mLoop As Integer
     
     If vsGrid.Row > 0 Then

        For mLoop = 2 To vsGrid.Row - 1
            If (vsGrid.Cell(flexcpChecked, mLoop, 1) = 2 Or vsGrid.Cell(flexcpChecked, mLoop, 2) = 2) And vsGrid.Row <> 1 Then
                MsgBox "Verify the Previous Year's Data"
                Exit Sub
            End If
        Next mLoop
        
        If vsGrid.Col = 1 Then
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 1) = 2 Then
                Call ExtarctMonthlyData(mYearID, val(vsGrid.TextMatrix(vsGrid.Row, 6)))
                frmExtractedCashBook.LoadMode = 2
                frmExtractedCashBook.YearID = mYearID
                frmExtractedCashBook.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                frmExtractedCashBook.Show vbModal
                Exit Sub
            Else
                frmExtractedCashBook.LoadMode = 2
                frmExtractedCashBook.YearID = mYearID
                frmExtractedCashBook.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                frmExtractedCashBook.cmdVerify.Enabled = False
                frmExtractedCashBook.Show vbModal
            End If
        End If
        If vsGrid.Col = 2 Then
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 1) = 1 Then
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 2) = 2 Then
                    frmExtractedBalanceSheet.LoadMode = 2
                    frmExtractedBalanceSheet.YearID = mYearID
                    frmExtractedBalanceSheet.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                    frmExtractedBalanceSheet.Show vbModal
                    Exit Sub
                Else
                    frmExtractedBalanceSheet.LoadMode = 2
                    frmExtractedBalanceSheet.YearID = mYearID
                    frmExtractedBalanceSheet.MonthID = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                    frmExtractedBalanceSheet.cmdVerify.Enabled = False
                    If val(vsGrid.TextMatrix(vsGrid.Row, 7)) = 2 Then
                        frmExtractedBalanceSheet.cmdPublish.Enabled = False
                        frmExtractedBalanceSheet.txtCouncilNo.Text = val(vsGrid.TextMatrix(vsGrid.Row, 3))
                        frmExtractedBalanceSheet.txtCouncilDate.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
                    Else
                        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                            frmExtractedBalanceSheet.cmdPublish.Enabled = True
                        Else
                            frmExtractedBalanceSheet.cmdPublish.Enabled = False
                        End If
                    End If
                    frmExtractedBalanceSheet.Show vbModal
                End If
            Else
                MsgBox "CASH AND BANK BALANCES NOT VERIFIED YET!!!!!!!!", vbInformation
                Exit Sub
            End If
        End If
     End If
End Sub
Private Sub ExtarctMonthlyData(mYearID As Integer, mMonthID As Integer)
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec  As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    Dim mExtractFlag As Boolean
    Dim mArrIn As Variant

    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mRowCnt = 1
    mSQL = " SELECT * FROM faDailyExtracts WHERE   intTypeID<>0 AND intFinancialYearID=" & mYearID & " AND intMonthID= " & mMonthID
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        mExtractFlag = True
    End If
    If mExtractFlag = False Then
        mArrIn = Array(mYearID)
        objDB.ExecuteSP "spDailyExtracts", mArrIn, , , mCnn, adCmdStoredProc
    End If
    Rec.Close
    mCnn.Close
End Sub
Public Property Let YearID(mData As Integer)
    mYearID = mData
End Property

Public Property Get YearID() As Integer
    YearID = mYearID
End Property
Private Sub VerifyStatus()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mRowCnt = 2
    mSQL = " SELECT * FROM faPostingIndex  "
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then

        vsGrid.Row = vsGrid.Row + 1
        While Not (Rec.EOF Or Rec.BOF)
            If mYearID = Rec!intFinYearID Then
                If val(vsGrid.TextMatrix(mRowCnt, 6)) = Rec!intMonthID Then
                    If Rec!tnyVerifyCash = 1 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 1) = vbChecked
                    End If
                    If Rec!tnyVerifyBS = 1 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 2) = vbChecked
                    End If
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                    If Rec!dtRefDate <> "" Then
                        vsGrid.TextMatrix(mRowCnt, 4) = DdMmmYy(IIf(IsNull(Rec!dtRefDate), "", Rec!dtRefDate))
                    Else
                        vsGrid.TextMatrix(mRowCnt, 4) = ""
                    End If
                    If Rec!tnyStage = 2 Then
                         vsGrid.Cell(flexcpChecked, mRowCnt, 5) = vbChecked
                    End If
                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!tnyStage), "", Rec!tnyStage)
                End If
                mRowCnt = mRowCnt + 1
            End If
        Rec.MoveNext
        
    Wend
    Rec.Close
    End If
End Sub
