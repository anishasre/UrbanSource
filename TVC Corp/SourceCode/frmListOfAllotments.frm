VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfAllotments 
   BackColor       =   &H00F4FAFA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Allotments"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13665
   ForeColor       =   &H00404040&
   Icon            =   "frmListOfAllotments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4860
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   13620
      _cx             =   24024
      _cy             =   8572
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      BackColorFixed  =   16055034
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   14540253
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
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfAllotments.frx":1CCA
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F4FAFA&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   13605
      TabIndex        =   1
      Top             =   6555
      Width           =   13665
      Begin VB.CommandButton cmdTreasuryBill 
         Caption         =   "&Treasury Bill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   11040
         TabIndex        =   16
         Top             =   30
         Width           =   1905
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   13260
         Top             =   270
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13665
      TabIndex        =   0
      Top             =   0
      Width           =   13665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4FAFA&
      Height          =   1020
      Left            =   15
      TabIndex        =   4
      Top             =   5505
      Width           =   13635
      Begin VB.TextBox txtAllotmentNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1365
         TabIndex        =   15
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   11910
         TabIndex        =   13
         Top             =   300
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dptFromDate 
         Height          =   315
         Left            =   9630
         TabIndex        =   11
         Top             =   375
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60489729
         CurrentDate     =   40134
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10125
         TabIndex        =   10
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   9
         Top             =   375
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearchSource 
         Caption         =   "..."
         Height          =   255
         Left            =   7470
         TabIndex        =   7
         Top             =   375
         Width           =   255
      End
      Begin VB.TextBox txtSource 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   11370
         TabIndex        =   12
         Top             =   375
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60489729
         CurrentDate     =   40134
      End
      Begin VB.Label lblAllotmentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allotment No"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   14
         Top             =   405
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   9960
         X2              =   10035
         Y1              =   540
         Y2              =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7905
         TabIndex        =   8
         Top             =   405
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2775
         TabIndex        =   5
         Top             =   390
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmListOfAllotments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private mLoadMode As Integer
    Private strAuthorityOrAllotment As Variant
    Private mPreviousYearMode As Integer
    
    'faAllotmentLetters.tnyStatus   8-Cancel,9-PDE Authority/Allotment Letters
    
    '*********************************************************************************************'
    '               Form to list all the Letter of Authority/Allotments                           '
    '*********************************************************************************************'
    
    Private Sub cmdNew_Click()
        Dim mCnn                    As New ADODB.Connection
        Dim objDB                   As New clsDB
        Dim Rec                     As New ADODB.Recordset
        Dim mSQL                    As String
        Dim mAuthorityOrAllotment   As Integer
        Dim mAryIn                  As Variant
        
        
        Dim mExtractedStatus As Integer
        Dim mMsg As String
        
        mExtractedStatus = GetStatusFlag
        If mExtractedStatus <> 2 Then
            cmdNew.Enabled = False
            
            mMsg = ""
            mMsg = mMsg + " Closing Balance Of Source Of Fund is " + vbCrLf
            mMsg = mMsg + " Either Not Brought Down  Or Approved " + vbCrLf
            mMsg = mMsg + " (Utility>>Annual Financial Statements-Finalization>>)"
            MsgBox mMsg, vbInformation
            Exit Sub
        End If
        
        
        
        
        
        '*********************************************************************************************'
        '               Procedure to Create a new Letter of Authority/Allotment'
        '*********************************************************************************************'
        On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
           
        If AuthorityOrAllotment = "Authority" Then
            mAuthorityOrAllotment = 1
        End If
        If mAuthorityOrAllotment = 1 Then
            'mAryIn = Array(mAuthorityOrAllotment, LoadMode, CheckDateInMMM(txtFromDate.Text), CheckDateInMMM(txtToDate.Text))
            Rec.CursorLocation = adUseClient
            'Set Rec = objDB.ExecuteSP("spSelectAllotmentLetters", mAryIn, , , mCnn, adCmdStoredProc)
            mSQL = "Select *,faAllotmentLetters.tnyStatus as Status,  faAllotmentLetters.fltAmount As Amount"
            mSQL = mSQL + " From faAllotmentLetters"
            mSQL = mSQL + " Left Join suSourceOfFund On suSourceOfFund.intSourceFundID = faAllotmentLetters.intSourceOfFundID"
            mSQL = mSQL + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faAllotmentLetters.intCategoryID"
            mSQL = mSQL + " Left Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
            mSQL = mSQL + " Left Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
            mSQL = mSQL + " Where intSourceOfFundID In(1,2,4,16,17,25,26,27,28,10,11,12,13,14,29,30,41)"
            mSQL = mSQL + " And tnyGroupID =" & LoadMode
            mSQL = mSQL + " And faAllotmentLetters.tnyStatus <> 9 And faAllotmentLetters.tnyStatus <> 8"
            If mPreviousYearMode = 1 Then
                mSQL = mSQL + " And faAllotmentLetters.intFinancialYearID =" & gbFinancialYearID - 1
            Else
                mSQL = mSQL + " And faAllotmentLetters.intFinancialYearID =" & gbFinancialYearID
            End If
            mSQL = mSQL + " AND ISNULL(faIDemandTBL.tnyStatus,0) <> 9 "
            
            'Dim mCount As Integer
            'mCount = 0
            
            
'''''''
'''''''            'C H A N G E D  B Y A I B Y  O N [28-MAR-2013]
'''''''            mSql = ""
'''''''            mSql = mSql + " Select faIDemandTbl.tnyStatus,faAllotmentLetters.tnyStatus,faAllotmentLetters.* From faAllotmentLetters"
'''''''            mSql = mSql + " Left Join faIDemandTbl On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID"
'''''''            mSql = mSql + " Left Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
'''''''            mSql = mSql + " WHERE NOT(ISNULL(faIDemandTbl.tnyStatus,0) NOT in ( 0) AND ISNull(faVouchers.tnyCancelFlag,0) <> 1 AND ISNULL(faVouchers.tnyReversed,0) <> 1)"
'''''''            mSql = mSql + " AND NOT ISNULL(faAllotmentLetters.tnyStatus,0) IN ( 8,9)"
'''''''            mSql = mSql + " AND intSourceOfFundID <> 3 AND tnyOpening <> 1"
'''''''            mSql = mSql + " AND faAllotmentLetters.intFinancialYearID = " & gbFinancialYearID
'''''''            mSql = mSql + " Order By intAllotmentID"
'''''''
'''''''
            Rec.Open mSQL, mCnn
'''''''            If Not (Rec.BOF And Rec.EOF) Then
'''''''                MsgBox "Please issue the Receipt for previous Letter of Authority", vbInformation
'''''''                Exit Sub
'''''''            End If
'''''''
            ' BLOCKED BY AIBY NEW CODE ADDED ABOVE
                        While Not Rec.EOF
                            If mPreviousYearMode = 1 Then
                                If IsNull(Rec!intVoucherNo) Then
                                    MsgBox "Please issue the Receipt for previous Letter of Authority", vbInformation
                                    Exit Sub
                                End If
                                'mCount = mCount + 1
                                End If
                            Rec.MoveNext
                        Wend
            
        End If
        
        'frmAllotmentLetter.PreviousYearMode = 1
        
        frmAllotmentLetter.LoadMode = LoadMode
        frmAllotmentLetter.AllotmentID = ""
        frmAllotmentLetter.AuthorityOrAllotment = AuthorityOrAllotment
        If AuthorityOrAllotment = "Authority" Then
            frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Receipt of A/C/D Fund in the Treasury Account"
        ElseIf AuthorityOrAllotment = "Allotment" Then
            frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Allotment of B Fund in the Consolidated Fund in the Treasury"
        ElseIf AuthorityOrAllotment = "OpeningAuthority" Then
            'BLOCK [1]
            'NOTE:- CHECKING Source of Fund Extraction is done or Not
            '       If done, no Opening Letter Of Authority/Allotment can ce done

                
                mMsg = ""
                mMsg = mMsg + "Previous year's Source wise transactions are all closed by Secretary" & vbCrLf
                mMsg = mMsg + "by brought down Source wise balances to new financial year by declaring the Source wise balances are correct." & vbCrLf
                mMsg = mMsg + "" & vbCrLf
                mMsg = mMsg + "Further changes in previous year's Opening source wise transaction will" & vbCrLf
                mMsg = mMsg + "make difference in Current year's Source wise allocations, thus this functionality is no more permitted " & vbCrLf
                
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                   MsgBox mMsg, vbInformation
                   Exit Sub
                End If
            'END OF BLOCK[1]
            frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Opening Letter of Authority Details"
            frmAllotmentLetter.txtAllotmentDate.Visible = True
            frmAllotmentLetter.txtAllotmentDate.Locked = False
            frmAllotmentLetter.txtAllotmentDate.Text = ""
        ElseIf AuthorityOrAllotment = "OpeningAllotment" Then
              'BLOCK [1]
            'NOTE:- CHECKING Source of Fund Extraction is done or Not
            '       If done, no Opening Letter Of Authority/Allotment can ce done

                
                mMsg = ""
                mMsg = mMsg + "Previous year's Source wise transactions are all closed by Secretary" & vbCrLf
                mMsg = mMsg + "by brought down Source wise balances to new financial year by declaring the Source wise balances are correct." & vbCrLf
                mMsg = mMsg + "" & vbCrLf
                mMsg = mMsg + "Further changes in previous year's Opening source wise transaction will" & vbCrLf
                mMsg = mMsg + "make difference in Current year's Source wise allocations, thus this functionality is no more permitted " & vbCrLf
                
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                   MsgBox mMsg, vbInformation
                   Exit Sub
                End If
            'END OF BLOCK[1]
            
            
            frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Opening Letter of Allotment Details "
            frmAllotmentLetter.txtAllotmentDate.Visible = True
            frmAllotmentLetter.txtAllotmentDate.Locked = False
            frmAllotmentLetter.txtAllotmentDate.Text = ""
        End If
        
        frmAllotmentLetter.Show vbModal
        Call FillGrid
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
     Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mTrAccHeadId As Integer
        
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                GetStatusFlag = -1
            End If
            Rec.Close
        End If
    End Function
    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub

    Private Sub cmdSearchSource_Click()
        '*********************************************************************************************'
        '                           Procedure to search Transaction Type                              '
        '*********************************************************************************************'

        On Error GoTo Err:
            If AuthorityOrAllotment = "Authority" Then
                'frmSearchMasters.SQLQry = "Select intTransactionTypeID,vchTransactionType From faTransactionType Where intTransactionTypeID In(108,109,110,125,126,155) Order By vchTransactionType"
                If gbLBPanchayat = 1 Then
                        frmSearchMasters.SQLQry = "Select intTransactionTypeID,vchTransactionType From faTransactionType Where intTransactionTypeID In(108,109,110,111,125,126,155,168,169,170,171,119,120,121,122,123,174) Order By vchTransactionType"
                    Else
                        frmSearchMasters.SQLQry = "Select intTransactionTypeID,vchTransactionType From faTransactionType Where intTransactionTypeID In(108,109,110,111,125,126,155,168,169,170,171,174) Order By vchTransactionType"
                End If
            ElseIf AuthorityOrAllotment = "Allotment" Then
                frmSearchMasters.SQLQry = "Select intTransactionTypeID,vchTransactionType From faTransactionType Where intTransactionTypeID In(112) Order By vchTransactionType"
            ElseIf AuthorityOrAllotment = "OpeningAuthority" Then
                frmSearchMasters.SQLQry = "Select intTransactionTypeID,vchTransactionType From faTransactionType Where intTransactionTypeID In(108,109,110,125,126,155) Order By vchTransactionType"
            ElseIf AuthorityOrAllotment = "OpeningAllotment" Then
                frmSearchMasters.SQLQry = "Select intTransactionTypeID,vchTransactionType From faTransactionType Where intTransactionTypeID In(112) Order By vchTransactionType"
            End If
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Show vbModal
            If gbSearchID > 0 Then
                txtSource.Text = gbSearchStr
                txtSource.Tag = gbSearchID
            Else
                txtSource.Text = ""
                txtSource.Tag = ""
            End If
            gbSearchID = -1
            gbSearchStr = ""
            
            'Call FillGrid
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdTreasuryBill_Click()
        If val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 4 Then
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = 1 Then
                frmViewAllotmentLetter.ArrayIn = Array(val(vsGrid.TextMatrix(vsGrid.Row, 8)))
                frmViewAllotmentLetter.Mode = 8
                frmViewAllotmentLetter.Show vbModal
            Else
                MsgBox "Letter Of Authority Not Approved", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "Treasury Bill Only For OWN FUND Sources", vbInformation
            Exit Sub
        End If
    End Sub

'''    Private Sub DTPicker1_CloseUp()
'''        txtFromDate.Text = CheckDateInMMM(DTPicker1.value)
'''        txtFromDate.SetFocus
'''    End Sub
'''    Private Sub DTPicker2_CloseUp()
'''        txtToDate.Text = CheckDateInMMM(DTPicker2.value)
'''    End Sub
    Private Sub dptFromDate_CloseUp()
        If CDate(dptFromDate.value) Then
        If CDate(dtpToDate.value) Then
                If CDate(dptFromDate.value) > CDate(dtpToDate.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dptFromDate.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtFromDate.Text = CheckDateInMMM(dptFromDate.value)
        End If
        txtFromDate.Text = CheckDateInMMM(dptFromDate.value)
        Call txtFromDate_LostFocus
    End Sub
    Private Sub dtpToDate_CloseUp()
        If CDate(dtpToDate.value) Then
            If CDate(dptFromDate.value) Then
                If CDate(dptFromDate.value) > CDate(dtpToDate.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpToDate.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtToDate.Text = CheckDateInMMM(dtpToDate.value)
        End If
      txtToDate.Text = CheckDateInMMM(dtpToDate.value)
    End Sub

    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
        
    End Sub
     Private Sub txtFromDate_LostFocus()
        Dim mDate As Date
        If Not IsDate(txtFromDate.Text) Then
            txtFromDate.Text = DdMmmYy(gbStartingDate)
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
        
        If CDate(txtFromDate.Text) < CDate(gbStartingDate) Then
            If CDate(txtFromDate.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                txtFromDate.Text = DateAdd("yyyy", -1, gbStartingDate)
                txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
            End If
            txtToDate.Text = DateAdd("yyyy", -1, gbEndingDate)
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
            mPreviousYearMode = 1
        Else
            mDate = CDate(txtFromDate)
            If Not (mDate >= gbStartingDate And mDate <= gbTransactionDate) Then
                txtFromDate.Text = DdMmmYy(gbStartingDate)
            End If
            If IsDate(txtToDate) Then
                mDate = CDate(txtToDate)
            Else
                mDate = gbTransactionDate
            End If
            If Not (mDate >= gbStartingDate And mDate <= gbTransactionDate) Then
                txtToDate.Text = DdMmmYy(gbEndingDate)
            End If
            mPreviousYearMode = 0
        End If
        Call FillGrid
    End Sub
    Private Sub txtToDate_LostFocus()
        Dim mDate As Date
        If Not IsDate(txtToDate.Text) Then
            txtToDate.Text = DdMmmYy(gbTransactionDate)
        Else
            txtToDate.Text = CheckDateInMMM(Trim(txtToDate))
        End If
        
        If CDate(txtToDate.Text) < CDate(gbStartingDate) Then
            If CDate(txtToDate.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                txtToDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
                txtFromDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbStartingDate))
            End If
            If CDate(txtFromDate.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                txtFromDate.Text = DateAdd("yyyy", -1, gbStartingDate)
                txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
            End If
            mPreviousYearMode = 1
        Else
            If IsDate(txtFromDate) Then
                mDate = CDate(txtFromDate)
            Else
                mDate = gbStartingDate
            End If
            If Not (mDate >= gbStartingDate And mDate <= gbTransactionDate) Then
                txtFromDate.Text = DdMmmYy(gbStartingDate)
            End If
            mDate = CDate(txtToDate)
            If Not (mDate >= gbStartingDate And mDate <= gbTransactionDate) Then
                txtToDate.Text = DdMmmYy(gbEndingDate)
            End If
            mPreviousYearMode = 0
        End If
        
        Call FillGrid
        
        
    End Sub

    Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyDelete Then
            txtSource.Text = ""
            txtSource.Tag = ""
        End If
    End Sub
    Private Sub txtToDate_GotFocus()
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate)
    End Sub
'    Private Sub txtToDate_LostFocus()
'        If Not IsDate(txtToDate.Text) Then
'            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
'        End If
''        if txtToDate.Text > < CDate("01/Jan/1900")
'        If txtToDate.Text > CDate(gbEndingDate) Then
'            MsgBox "Please Enter a  Date within this Financial Year", vbInformation
'            txtToDate.Text = ""
'            txtToDate.SetFocus
'            Exit Sub
'        End If
'        If CDate(txtToDate.Text) Then
'            If CDate(txtFromDate.Text) Then
'                If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
'                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
'                    txtToDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
'                    Exit Sub
'                End If
'            Else
'                MsgBox "Please Enter From date", vbInformation
'                Exit Sub
'            End If
'        Else
'            txtToDate.Text = CheckDateInMMM(txtFromDate.Text)
'        End If
'
'        'Call FillGrid
'    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Call FillGrid
    End Sub

    Private Sub Form_Load()
        vsGrid.Cell(flexcpFontName, 0) = "Verdana"
        WindowsXPC1.InitSubClassing
        dtpToDate.value = Date
        dptFromDate.value = Date
        'txtToDate.Text = CheckDateInMMM(Date)
        'txtFromDate.Text = DateAdd("m", -1, CheckDateInMMM(txtToDate.Text))
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            cmdNew.Enabled = True
        Else
            cmdNew.Enabled = False
        End If
        If AuthorityOrAllotment = "Authority" Then
            frmListOfAllotments.Caption = "Letter of Authority List"
        ElseIf AuthorityOrAllotment = "Allotment" Then
            frmListOfAllotments.Caption = "Letter of Allotment List"
        ElseIf AuthorityOrAllotment = "OpeningAuthority" Then
            frmListOfAllotments.Caption = "List Of Opening Letter of Authority"
        ElseIf AuthorityOrAllotment = "OpeningAllotment" Then
            frmListOfAllotments.Caption = "List Of Opening Letter of Allotment"
        End If
        Call FillGrid
        
    End Sub
    
    Private Function FillGrid() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDB As New clsDB
            Dim mSQL As String
            Dim mRowCnt As Integer
            Dim mStr As String
            Dim aryIn As Variant
            Dim mAuthorityOrAllotment As Integer
            '*********************************************************************************************'
            '               Function to List all the Letter of Authority/Allotments                       '
            '*********************************************************************************************'
            
            If objDB.SetConnection(mCnn) Then
                If AuthorityOrAllotment = "Authority" Then
                    mAuthorityOrAllotment = 1
                ElseIf AuthorityOrAllotment = "Allotment" Then
                    mAuthorityOrAllotment = 2
                ElseIf AuthorityOrAllotment = "OpeningAuthority" Then
                    mAuthorityOrAllotment = 3
                ElseIf AuthorityOrAllotment = "OpeningAllotment" Then
                    mAuthorityOrAllotment = 4
                End If
                'aryIn = Array(mAuthorityOrAllotment, LoadMode, CheckDateInMMM(txtFromDate.Text), CheckDateInMMM(txtToDate.Text))
                Rec.CursorLocation = adUseClient
                'Set Rec = objDB.ExecuteSP("spSelectAllotmentLetters", aryIn, , , mCnn, adCmdStoredProc)
                 mSQL = "Select *,faAllotmentLetters.tnyStatus as Status,  faAllotmentLetters.fltAmount As Amount,faReverseEntry.tnyStatus[ReverseStatus] "
                mSQL = mSQL + " ,faVouchers.tnyCancelFlag CancelFlag,faVouchers.tnyStatus VrStatus,faIDemandTBL.tnyStatus DemandStatus,faAllotmentLetters.intSourceOfFundID"
                mSQL = mSQL + " From faAllotmentLetters"
                If mPreviousYearMode = 1 Then
                    mSQL = mSQL + " LEFT JOIN faPendingTaskRequest ON faPendingTaskRequest.intKeyID = faAllotmentLetters.intAllotmentID AND faPendingTaskRequest.intTaskID = 1 AND NOT faPendingTaskRequest.tnyStatus IN (0,4) "
                End If
                mSQL = mSQL + " Left Join suSourceOfFund On suSourceOfFund.intSourceFundID = faAllotmentLetters.intSourceOfFundID"
                mSQL = mSQL + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faAllotmentLetters.intCategoryID"
                mSQL = mSQL + " Left Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID "
                mSQL = mSQL + "     And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                mSQL = mSQL + "     And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
                mSQL = mSQL + " Left Join faVouchers On faVouchers.intKeyID2 = faIDemandTBL.numDemandID And tnyVoucherTypeID = 10"
                mSQL = mSQL + " Left Join faReverseEntryChild On faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                mSQL = mSQL + " Left Join faReverseEntry On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID And faReverseEntry.tnyStatus = 2"
                If mAuthorityOrAllotment = 1 Then
                    mSQL = mSQL + " Where faAllotmentLetters.intSourceOfFundID In(1,2,4,16,17,21,25,26,27,28,10,11,12,13,14,29,30,41)"
                ElseIf mAuthorityOrAllotment = 2 Then
                    mSQL = mSQL + " Where faAllotmentLetters.intSourceOfFundID In(3,19)"
                ElseIf mAuthorityOrAllotment = 3 Then
                    mSQL = mSQL + "Where faAllotmentLetters.intSourceOfFundID In(1,2,4,16,17,21,25,26,27,28,10,11,12,13,14,29,30,41) and tnyOpening=1"
                ElseIf mAuthorityOrAllotment = 4 Then
                    mSQL = mSQL + "Where faAllotmentLetters.intSourceOfFundID In(3) and tnyOpening=1"
                End If
                mSQL = mSQL + " And tnyGroupID =" & LoadMode
                mSQL = mSQL + " And faAllotmentLetters.tnyStatus <> 9 And faAllotmentLetters.tnyStatus <> 8"
                If Trim(txtAllotmentNo.Text) <> "" Then
                    mSQL = mSQL + " And vchAllotmentNo Like '" & txtAllotmentNo & "%'"
                End If
                If txtFromDate.Text <> "" And txtToDate.Text <> "" Then
                    mSQL = mSQL + " And dtAllotmentDate between '" & txtFromDate & "' and '" & txtToDate & "'"
                End If
                If txtSource.Tag <> "" Then
                    mSQL = mSQL + " And faAllotmentLetters.intTransactionTypeID = " & txtSource.Tag
                End If
                
                If mPreviousYearMode = 1 Then
                    mSQL = mSQL + " And faAllotmentLetters.intFinancialYearID=" & gbFinancialYearID - 1 & " "
                Else
                    mSQL = mSQL + " And faAllotmentLetters.intFinancialYearID=" & gbFinancialYearID & " "
                End If
                
                mSQL = mSQL + " Order by dtAllotmentDate"
                Rec.Open mSQL, mCnn
                vsGrid.Rows = 2
                mRowCnt = 1
                vsGrid.Clear 1, 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intInstalmentNo), "", Rec!intInstalmentNo)
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
                    
                    If IsNull(Rec!intVoucherNo) Then
                        If IsNull(Rec!vchDemandNo) Then
                            vsGrid.TextMatrix(mRowCnt, 6) = "No-Demand/Receipt"
                        ElseIf Rec!vchDemandNo > 0 Then
                            If Rec!DemandStatus = 9 Then
                                mStr = "Demand Cancelled"
                                vsGrid.TextMatrix(mRowCnt, 6) = mStr
                            Else
                                mStr = "Demand-" + CStr(Rec!vchDemandNo)
                                vsGrid.TextMatrix(mRowCnt, 6) = mStr
                            End If
                        End If
                        If Rec!tnyOpening = 1 And mAuthorityOrAllotment = 3 Then
                            vsGrid.TextMatrix(mRowCnt, 6) = "Opening Letter Of Authority"
                        End If
                        If Rec!tnyOpening = 1 And mAuthorityOrAllotment = 4 Then
                            vsGrid.TextMatrix(mRowCnt, 6) = "Opening Letter Of Allotment"
                        End If
                    Else
                        If Rec!tnyOpening = 1 And mAuthorityOrAllotment = 3 Then
                            vsGrid.TextMatrix(mRowCnt, 6) = "Opening Letter Of Authority"
                        ElseIf Rec!tnyOpening = 1 And mAuthorityOrAllotment = 4 Then
                            vsGrid.TextMatrix(mRowCnt, 6) = "Opening Letter Of Allotment"
                        ElseIf Rec!intVoucherNo > 0 Then
                            If (IIf(IsNull(Rec!VrStatus), 0, Rec!VrStatus) = 4 And IIf(IsNull(Rec!CancelFlag), 0, Rec!CancelFlag) = 1) Then
                                mStr = "Receipt Cancelled"
                                vsGrid.TextMatrix(mRowCnt, 6) = mStr
                                vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbUnchecked
                            ElseIf (Rec!ReverseStatus = 2) Then
                                mStr = "Receipt Reversed"
                                vsGrid.TextMatrix(mRowCnt, 6) = mStr
                                vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbUnchecked
                            Else
                                mStr = "Receipt-" + CStr(Rec!intVoucherNo)
                                vsGrid.TextMatrix(mRowCnt, 6) = mStr
                                vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbChecked
                            End If
                        End If
                    End If
                    
                    If Rec!Status = 0 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 7) = vbUnchecked
                    ElseIf Rec!Status = 1 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 7) = vbChecked
                    ElseIf Rec!Status = 2 Then
                        vsGrid.Cell(flexcpChecked, mRowCnt, 7) = vbChecked
                        vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbChecked
                    End If
                    If (IIf(IsNull(Rec!VrStatus), 0, Rec!VrStatus) = 4 And IIf(IsNull(Rec!CancelFlag), 0, Rec!CancelFlag) = 1) Or (IIf(IsNull(Rec!ReverseStatus), 0, Rec!ReverseStatus) = 2) Then
                        vsGrid.Cell(flexcpBackColor, mRowCnt, 0, , 9) = &HC0E0FF
                    End If
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                    vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!intSourceOfFundID), "", Rec!intSourceOfFundID)
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Finance does not Exist, Please contact your System Administrator", vbInformation
            End If
            
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Public Property Let LoadMode(mData As Integer)
        mLoadMode = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = mLoadMode
    End Property
    
     Public Property Let AuthorityOrAllotment(mData As Variant)
        strAuthorityOrAllotment = mData
    End Property
    
    Public Property Get AuthorityOrAllotment() As Variant
        AuthorityOrAllotment = strAuthorityOrAllotment
    End Property
    
    Private Sub vsGrid_DblClick()
        On Error GoTo Err:
            If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then Exit Sub
            If ChkFinancialYear = False Then
                MsgBox "Financial Year Not Matching ", vbInformation
                Exit Sub
            Else
                If Me.LoadMode = 50 Then
                    Dim mMsg As String
                    Dim mExtractedStatus As Integer
                    mMsg = ""
                    mMsg = mMsg + "Previous year's Source wise transactions are all closed by Secretary" & vbCrLf
                    mMsg = mMsg + "by brought down Source wise balances to new financial year by declaring the Source wise balances are correct." & vbCrLf
                    mMsg = mMsg + "" & vbCrLf
                    mMsg = mMsg + "Further changes in previous year's Opening source wise transaction will" & vbCrLf
                    mMsg = mMsg + "make difference in Current year's Source wise allocations, thus this functionality is no more permitted " & vbCrLf
                    
                    mExtractedStatus = GetStatusFlag
                    If mExtractedStatus = 2 Then
                       MsgBox mMsg, vbInformation
                       Exit Sub
                    End If
                End If
                
                frmAllotmentLetter.PreviousYearMode = mPreviousYearMode
                frmAllotmentLetter.LoadMode = Me.LoadMode
                frmAllotmentLetter.AuthorityOrAllotment = Me.AuthorityOrAllotment
                frmAllotmentLetter.AllotmentID = vsGrid.TextMatrix(vsGrid.Row, 8)
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = 2 Then  '2
                    frmAllotmentLetter.ApproveStatus = 0
                Else
                    frmAllotmentLetter.ApproveStatus = 1
                End If
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 9) = 2 Then
                    If vsGrid.TextMatrix(vsGrid.Row, 6) = "Receipt Cancelled" Or vsGrid.TextMatrix(vsGrid.Row, 6) = "Receipt Reversed" Then
                        frmAllotmentLetter.cmdRegenerateDemand.Visible = True
                        frmAllotmentLetter.cmdRegenerateDemand.Enabled = True
                        frmAllotmentLetter.txtAllotmentDate.Tag = ""
                    End If
                End If
                frmAllotmentLetter.txtAllotmentNo.Enabled = False
                frmAllotmentLetter.Show vbModal
                Call FillGrid
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    Private Function ChkFinancialYear() As Boolean
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mFlag   As Boolean
        Dim mFinancialYearId As Variant
        Dim mYearID As Integer
        
        If objDB.SetConnection(mCnn) Then
                mSQL = "select * from faAllotmentLetters where intAllotmentId =" & vsGrid.TextMatrix(vsGrid.Row, 8) & ""
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mFinancialYearId = IIf(IsNull(Rec!intFinancialYearID), Null, Rec!intFinancialYearID)
                    If mPreviousYearMode = 1 Then
                        mYearID = gbFinancialYearID - 1
                    Else
                        mYearID = gbFinancialYearID
                    End If
                    If mFinancialYearId = mYearID Then
                        mFlag = True
                    Else
                        mFlag = False
                    End If
                    ChkFinancialYear = mFlag
                End If
        End If
    End Function
