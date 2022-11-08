VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmAllotmentClosingBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Allotment Closing Balance"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   1740
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   13740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6795
      TabIndex        =   3
      Top             =   6285
      Width           =   1965
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   13740
      TabIndex        =   2
      Top             =   0
      Width           =   13740
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Extract Balance of Appropriation Control Register"
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
         Height          =   420
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   13815
      End
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4875
      TabIndex        =   1
      Top             =   6285
      Width           =   1815
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5295
      Left            =   -60
      TabIndex        =   0
      Top             =   795
      Width           =   13725
      _cx             =   24209
      _cy             =   9340
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   200
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAllotmentClosingBalance.frx":0000
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
End
Attribute VB_Name = "frmAllotmentClosingBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Public RowCount As Variant
Private Sub FillGrid()
        Dim objdb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim mRow As Integer
        cmdApprove.Enabled = False
        
        vsGrid.Rows = 1
        vsGrid.TextMatrix(0, 3) = "Sl.No"
        vsGrid.TextMatrix(0, 4) = "Source"
        vsGrid.TextMatrix(0, 5) = "Category"
        vsGrid.TextMatrix(0, 6) = "Scheme"
        vsGrid.TextMatrix(0, 7) = "Amount"
     
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " SELECT faExtractAllotments.intID intID,tnyStatus,suSourceofFund.intSourceFundID as SourceID,suSourceofFund.vchSourceFundName as Source,faTransactionCategory.intCategoryID as CategoryID,"
        mSql = mSql + " faTransactionCategory.vchTransactionCategory as Category,fltAmount as Balance"
        mSql = mSql + " From faExtractAllotments"
        mSql = mSql + " INNER JOIN suSourceofFund ON faExtractAllotments.intSourceofFundID=suSourceofFund.intSourceFundID"
        mSql = mSql + " LEFT JOIN  faTransactionCategory ON faExtractAllotments.intCategoryID=faTransactionCategory.intCategoryID"
      '  mSql = mSql + " LEFT JOIN faDepSchPro ON  faExtractAllotments.intSchemeID=faDepSchPro.intID"
        mSql = mSql + " Where intFinancialYearID=" & gbFinancialYearID - 1
        Rec.Open mSql, mCnn
        mRow = 1
        If Not (Rec.BOF And Rec.EOF) Then
             While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!intID), "", Rec!intID)
                vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!SourceID), "", Rec!SourceID)
                vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!CategoryID), "", Rec!CategoryID)
                vsGrid.TextMatrix(mRow, 3) = mRow
                vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!Source), "", Rec!Source)
                vsGrid.TextMatrix(mRow, 5) = IIf(IsNull(Rec!Category), "NULL", Rec!Category)
                vsGrid.TextMatrix(mRow, 7) = IIf(IsNull(Rec!Balance), "", Rec!Balance)
                'If Rec!tnyStatus = 1 Then
                       'vsGrid.Cell(flexcpBackColor, mRow, 0, , 7) = &HC0FFC0
                'End If
                If Rec!tnyStatus <> 2 Then
                  cmdApprove.Enabled = True
                End If
                Rec.MoveNext
                mRow = mRow + 1
             Wend
             RowCount = mRow
             cmdExtract.Visible = False
             cmdApprove.Visible = True
        End If
End Sub
Private Sub cmdApprove_Click()
        Dim objdb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSql As String
        Dim i As Integer
        mRow = RowCount
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If GetLastPostingYear = True Then
            MsgBox "Transactions Locked for the Year!!!No more Transcations is possible", vbInformation
            cmdApprove.Enabled = False
        Else
            For i = 1 To RowCount - 1
              mSql = "Update faExtractAllotments Set fltAmount=" & val(vsGrid.Cell(flexcpText, i, 7)) & ", tnyStatus = 2 "
              mSql = mSql + " Where intID=" & val(vsGrid.Cell(flexcpText, i, 0)) & " And intSourceofFundID=" & val(vsGrid.Cell(flexcpText, i, 1))
              mSql = mSql + " And intCategoryID=" & val(vsGrid.Cell(flexcpText, i, 2)) & " And  tnyOpening=1 And intFinancialYearID=" & gbFinancialYearID - 1
              Rec.Open mSql, mCnn
            Next
            MsgBox "Saved Sucsessfully", vbInformation
            cmdApprove.Enabled = False
        
        End If
End Sub
Private Sub cmdExtract_Click()
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSql As String
        Dim mArrIn As Variant
        'mArrIn = Array(gbFinancialYearID, 0) 'MODIFIED BY SUNIL ON 12.06.2012
        mArrIn = Array(2014, 0)
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        objdb.ExecuteSP "spExtractAllotment", mArrIn, , , mCnn
        Call FillGrid
End Sub

Private Sub Form_Load()


    cmdExtract.Visible = True
    cmdApprove.Visible = False
    
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    
    'objdb.SetConnection mCnn
    'mSql = "DELETE FROM  faExtractAllotments WHERE intFinancialYearID = 2013"
    'mCnn.Execute mSql
    
    Call FillGrid
End Sub
 Private Function GetLastPostingYear() As Boolean
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mSql    As String
      

        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "SELECT  intFinYearID FROM faPostingIndex WHERE tnyStage=2 AND intFinYearID=" & gbFinancialYearID - 1
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            GetLastPostingYear = True
        Else
            GetLastPostingYear = False
        End If
        
        Rec.Close
        mCnn.Close
    End Function

