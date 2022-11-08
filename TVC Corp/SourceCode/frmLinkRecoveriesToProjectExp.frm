VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmLinkRecoveriesToProjectExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Link Recoveries To Project Expenditure"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSendSulekha 
      Caption         =   "SEND DATA TO SULEKHA"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11085
      TabIndex        =   2
      Top             =   6420
      Width           =   1350
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5310
      Left            =   15
      TabIndex        =   1
      Top             =   840
      Width           =   12780
      _cx             =   22542
      _cy             =   9366
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLinkRecoveriesToProjectExp.frx":0000
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   12810
      TabIndex        =   0
      Top             =   0
      Width           =   12840
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   10530
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   135
         Width           =   1875
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9135
         TabIndex        =   4
         Top             =   180
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmLinkRecoveriesToProjectExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mDate As Date
Dim mYearID As Integer

    Private Sub cmbYear_Change()
        If cmbYear.ListIndex > -1 Then
            mYearID = cmbYear.ItemData(cmbYear.ListIndex)
            Call FillGrid
        End If
    End Sub

Private Sub cmbYear_Click()
    If cmbYear.ListIndex > -1 Then
            mYearID = cmbYear.ItemData(cmbYear.ListIndex)
            Call FillGrid
        End If
End Sub

Private Sub cmdSendSulekha_Click()
    Dim mCnn        As New ADODB.Connection
    Dim objDB       As New clsDB
    Dim Rec         As New ADODB.Recordset
    Dim mSQL        As String
    Dim mCnt        As Integer
    Dim mCheck      As Boolean
    Dim arrInput    As Variant
    Dim arrOutPut   As Variant
    Dim objVrSub    As uVoucherSub
    Dim mCount      As Integer
    mCheck = False
    For mCnt = 0 To vsGrid.Rows - 1
        If vsGrid.TextMatrix(mCnt, 0) <> "" Then
            If vsGrid.Cell(flexcpChecked, mCnt, 8) = vbChecked Then
                mCheck = True
                Exit For
            End If
        End If
    Next
    If mCheck Then
        For mCnt = 0 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCnt, 9) <> "" Then
                If vsGrid.Cell(flexcpChecked, mCnt, 8) = vbChecked Then
                    If objDB.SetConnection(mCnn) Then
                        mSQL = "SELECT intVoucherID FROM faVoucherSub WHERE intVoucherID = " & val(vsGrid.TextMatrix(mCnt, 9)) & "  "
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mCount = 1
                        Else
                            mCount = 0
                        End If
                        Rec.Close
                        If mCount = 1 Then
                            mSQL = " UPDATE faVoucherSub SET "
                            mSQL = mSQL + " decProjectID = " & val(vsGrid.TextMatrix(mCnt, 10)) & ", intSourceOfFundID =" & val(vsGrid.TextMatrix(mCnt, 11)) & ",  intTypeID =8"
                            mSQL = mSQL + " WHERE intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 9))
                            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                        Else
                            With objVrSub
                                .intVoucherID = val(vsGrid.TextMatrix(mCnt, 9))
                                .decProjectID = val(vsGrid.TextMatrix(mCnt, 10))
                                .intSourceOfFundID = val(vsGrid.TextMatrix(mCnt, 11))
                                .intCategoryID = Null
                                .intSectorID = Null
                                .intAllotmentID = Null
                                .intAgreementID = Null
                                .intCashBookID = Null
                                .intImplementingOfficerID = Null
                                .intCreditorTypeID = Null
                                .intCreditorsID = Null
                                .intTypeID = 8                 'To Identify Journals with recoveries to project expenditure
                                .intLocalBodyID = gbLocalBodyID
                                
                                arrInput = Array(.intVoucherID, _
                                                .intLocalBodyID, _
                                                .decProjectID, _
                                                .intSourceOfFundID, _
                                                .intCategoryID, _
                                                .intSectorID, _
                                                .intAllotmentID, _
                                                .intAgreementID, _
                                                .intCashBookID, _
                                                .intImplementingOfficerID, _
                                                .intCreditorTypeID, _
                                                .intCreditorsID, _
                                                .intTypeID)
                                objDB.ExecuteSP "spSaveVoucherSub", arrInput, , , mCnn
                                
                            End With
                      End If
                  End If
                MsgBox "Successfully Added", vbApplicationModal
                Call UpdateDetailsToSulekha(val(vsGrid.TextMatrix(mCnt, 9)), mCnt)
                End If
            End If
        Next
    Else
        MsgBox "Please Select a row", vbApplicationModal
    End If
End Sub
Private Sub GetVoucherDetails(mVoucherID As Long)
  
    Dim mCnn As New ADODB.Connection
    Dim objDB   As New clsDB
    Dim mSQL As String
    Dim Rec As New ADODB.Recordset

    If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
        mSQL = "Select * from faVouchers Where intVoucherID=" & mVoucherID & "  "
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
           mDate = DdMmmYy(Rec!dtDate)
        End If
    End If
    Rec.Close
    mCnn.Close
End Sub
Private Sub UpdateDetailsToSulekha(mVoucherID As Long, mCnt As Integer)
     Dim mCnnSulekha   As New ADODB.Connection
     Dim objDB   As New clsDB
     Dim mSQL As String
     Dim arrInput As Variant
     Dim mCount As Integer
     Dim Rec  As New ADODB.Recordset
     
     GetVoucherDetails (mVoucherID)
     If mVoucherID > 0 Then
         If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
            mSQL = "Select * from ExpenseDetails where intVoucherID = " & mVoucherID & "  "
            Rec.Open mSQL, mCnnSulekha
            If Not (Rec.EOF And Rec.BOF) Then
                mCount = 1
            Else
                mCount = 0
            End If
            Rec.Close
            If mCount = 0 Then
                arrInput = Array(gbLBID, _
                                mYearID, _
                                val(vsGrid.TextMatrix(mCnt, 10)), _
                                -1, val(vsGrid.TextMatrix(mCnt, 11)), _
                                val(vsGrid.TextMatrix(mCnt, 4)), _
                                mVoucherID, mDate)
        
                objDB.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
            End If
         Else
            MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
            Exit Sub
         End If
         mCnnSulekha.Close
     End If
End Sub
Private Sub Form_Load()
    mYearID = gbFinancialYearID
    Call FillYear
    Call FillGrid
End Sub
Private Function CheckDataInSulekha(mVoucherID As Variant) As Boolean
    Dim mCnnSulekha   As New ADODB.Connection
    Dim mCnn As New ADODB.Connection
    Dim objDB   As New clsDB
    Dim mSQL As String
    Dim arrInput As Variant
    Dim Rec As New ADODB.Recordset
    Dim RecFin As New ADODB.Recordset
    
    If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "Select * from ExpenseDetails where intVoucherID = " & mVoucherID & "  "
            Rec.Open mSQL, mCnnSulekha
            If Not (Rec.EOF And Rec.BOF) Then
                CheckDataInSulekha = True
                mSQL = "Select * from faPayOrder Where intVoucherID=" & mVoucherID & "  "
                RecFin.Open mSQL, mCnn
                If Not (RecFin.EOF And RecFin.BOF) Then
                    If Rec!tnyCancelation <> RecFin!tnyCancelled Then
                        mSQL = ""
                        mSQL = " UPDATE ExpenseDetails SET tnyCancelation = " & RecFin!tnyCancelled & " , "
                        mSQL = mSQL + " tnyTransfer = 0 "
                        mSQL = mSQL + " WHERE intVoucherID=" & mVoucherID & "  "
                        objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                    End If
                End If
                RecFin.Close
            Else
                CheckDataInSulekha = False
            End If
            Rec.Close
        Else
            MsgBox "Connection to Saankhya Database doesnot exist", vbInformation, "Saankhya"
            Exit Function
        End If
    Else
        MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
        Exit Function
    End If
    mCnnSulekha.Close
    mCnn.Close
End Function
Private Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    Dim mLoop   As Integer
    On Error GoTo Err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSQL = " SELECT intKeyID2, JVAmt, PAmt , JVAmt-PAmt DrAmt,JVVrID,JVVrNo,POVrID,POVrNo," & vbCrLf
    mSQL = mSQL + " faPayOrder.numProjectNo,chvProjectnameEnglish,intSourceOfFundID,vchSourceFundName, " & vbCrLf
    mSQL = mSQL + " faPayOrder.tnyCancelled From " & vbCrLf
    mSQL = mSQL + "         (" & vbCrLf
    mSQL = mSQL + "         SELECT intKeyID2, Sum(JVAmt) JVAmt, SUM( PAmt) PAmt,Sum(IsNull(JVVrID,0)) JVVrID,Sum(IsNull(JVVrNo,0)) JVVrNo," & vbCrLf
    mSQL = mSQL + "             Sum(IsNull(POVrID,0)) POVrID,Sum(IsNull(POVrNo,0)) POVrNo FROM" & vbCrLf
    mSQL = mSQL + "         (" & vbCrLf
    mSQL = mSQL + "         SELECT  intKeyID2, fltAmount JVAmt, 0 PAmt,A.intVoucherID JVVrID,intVoucherNo JVVrNo,NULL POVrID,NULL POVrNo"
    mSQL = mSQL + "         From"
    mSQL = mSQL + "            ("
    mSQL = mSQL + "             SELECT  Max(intVoucherID) intVoucherID" & vbCrLf
    mSQL = mSQL + "             From faVouchers" & vbCrLf
    mSQL = mSQL + "             Where IsNull(intKeyID2, 0) > 0 And tnyVoucherTypeID = 40" & vbCrLf
    mSQL = mSQL + "             Group by intKeyID2" & vbCrLf
    mSQL = mSQL + "             ) A" & vbCrLf
    mSQL = mSQL + "             INNER JOIN faVouchers ON faVouchers.intVoucherID = A.intVoucherID" & vbCrLf
         
    'mSQL = mSQL + "         SELECT intKeyID2, fltAmount JVAmt, 0 PAmt,intVoucherID JVVrID,intVoucherNo JVVrNo,NULL POVrID,NULL POVrNo From faVouchers WHERE ISNULL(intKeyID2,0) > 0 AND tnyVoucherTypeID = 40" & vbCrLf
    
    mSQL = mSQL + "         Union" & vbCrLf
    mSQL = mSQL + "         SELECT intKeyID2, 0 JVAmt, fltAmount PAmt,NULL JVVrID,NULL JVVrNo,intVoucherID POVrID,intVoucherNo POVrNo From faVouchers WHERE ISNULL(intKeyID2,0) > 0 AND tnyVoucherTypeID = 20" & vbCrLf
    mSQL = mSQL + "         ) A" & vbCrLf
    mSQL = mSQL + "         Group by intKeyID2" & vbCrLf
    mSQL = mSQL + "         Having Sum(JVAmt) <> Sum(PAmt)" & vbCrLf
    mSQL = mSQL + "         ) B" & vbCrLf
    mSQL = mSQL + " INNER JOIN faPayOrder ON faPayOrder.vchPayOrderNo = B.intKeyID2" & vbCrLf
    mSQL = mSQL + " INNER JOIN suProjectDetails ON faPayOrder.numProjectNo=suProjectDetails.decProjectID" & vbCrLf
    mSQL = mSQL + " INNER JOIN suSourceOfFund ON faPayOrder.intSourceOfFundID=suSourceOfFund.intSourceFundID" & vbCrLf
    mSQL = mSQL + " Where  IsNull(faPayOrder.numProjectNo, 0) <> 0" & vbCrLf
    mSQL = mSQL + " And B.JVAmt  > 0 And faPayOrder.intFinancialYearID=" & mYearID & vbCrLf
    'faPayOrder.tnyCancelled <> 1 And
    
    Rec.CursorLocation = adUseClient
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!JVVrNo), "", Rec!JVVrNo)
        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!POVrNo), "", Rec!POVrNo)
        vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!JVAmt), "", Rec!JVAmt)
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!DrAmt), "", Rec!DrAmt)
        vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
        vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
        If Rec!tnyCancelled = 1 Then
            For mLoop = 0 To vsGrid.Cols - 1
                vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HC0E0FF
            Next mLoop
            vsGrid.TextMatrix(mRowCnt, 7) = "CANCELLED"
            'vsGrid.Editable = flexEDNone
            'Call vsGrid_BeforeEdit(mRowCnt, 7)
            'vsGrid.Cell(flexcpData, mRowCnt) = flexEDNone
        Else
            vsGrid.TextMatrix(mRowCnt, 7) = "APPROVED"
        End If
        If CheckDataInSulekha(IIf(IsNull(Rec!JVVrID), 0, Rec!JVVrID)) = True Then
            vsGrid.TextMatrix(mRowCnt, 8) = vbChecked
        Else
            vsGrid.TextMatrix(mRowCnt, 8) = vbUnchecked
        End If
        vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!JVVrID), "", Rec!JVVrID)
        vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!numProjectNo), "", Rec!numProjectNo)
        vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intSourceOfFundID), "", Rec!intSourceOfFundID)
        vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!POVrID), "", Rec!POVrID)
        vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!tnyCancelled), 0, Rec!tnyCancelled)
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    Rec.Close
    Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsGrid.TextMatrix(Row, 13) = 1 Then
        Cancel = True
    End If
End Sub

     Private Sub FillYear()
        PopulateList cmbYear, "Select Cast(intFinancialYearID as varchar(4)) + '-' + Right(Cast(intFinancialYearID+1 as varchar(4)),2),intFinancialYearID  From faFinancialYear WHERE intFinancialYearID > 2011", , , , True
        cmbYear.ListIndex = cmbYear.ListCount - 1
        vsGrid.SelectionMode = flexSelectionByRow
    End Sub
