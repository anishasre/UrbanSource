VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPDEAllotments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDE Allotments"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11565
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   450
      Left            =   7755
      TabIndex        =   13
      Top             =   6930
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      Height          =   420
      Left            =   9855
      TabIndex        =   6
      Top             =   6150
      Width           =   870
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6600
      MaxLength       =   8
      TabIndex        =   4
      Top             =   6195
      Width           =   1830
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3960
      TabIndex        =   3
      Top             =   6195
      Width           =   1830
   End
   Begin VB.TextBox txtAuthorityAllotmentNum 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   2
      Top             =   6195
      Width           =   1830
   End
   Begin VB.CommandButton cmdApprove 
      Appearance      =   0  'Flat
      Caption         =   "Close the Register"
      Enabled         =   0   'False
      Height          =   525
      Left            =   9510
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6885
      Width           =   1845
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8655
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6150
      Width           =   870
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmPDEAllotments.frx":0000
      Left            =   975
      List            =   "frmPDEAllotments.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   870
      Width           =   2160
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   -3630
      Top             =   7425
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   4455
      Left            =   135
      TabIndex        =   0
      Top             =   1395
      Width           =   11265
      _cx             =   19870
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPDEAllotments.frx":0004
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11565
      TabIndex        =   8
      Top             =   0
      Width           =   11565
   End
   Begin VB.Label Label4 
      Caption         =   "Amount :"
      Height          =   270
      Left            =   5880
      TabIndex        =   12
      Top             =   6195
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   6195
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "Authority Number :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   10
      Top             =   6195
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   390
      TabIndex        =   9
      Top             =   870
      Width           =   555
   End
End
Attribute VB_Name = "frmPDEAllotments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCheckDemand As Variant
 Private strAuthorityOrAllotment As Variant
    Private Sub cmbType_Click()
        If cmbType.ListIndex = 0 Then
            Label2.Caption = "Authority Number:"
        ElseIf cmbType.ListIndex = 1 Then
            Label2.Caption = "Allotment Number"
        End If
        Call FillGrid
        Call FormInitialize
    End Sub

    Private Sub cmdApprove_Click()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL    As String
        
        If MsgBox("Do you want to close this register?No more Entries can be Made", vbYesNo, "Saankhya") = vbYes Then
            If cmbType.ListIndex = 0 Then
                mSQL = "Update faAllotmentRegister set tnyStatus=1 where tnyStatus=0 and tnyTypeID=1"
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If cmbType.ListIndex = 1 Then
                mSQL = "Update faAllotmentRegister set tnyStatus=1 where tnyStatus=0 and tnyTypeID=2"
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            cmdSave.Enabled = False
            cmdApprove.Enabled = False
'            cmdReject.Enabled = False
            VSGrid.Editable = flexEDNone
            MsgBox "Approved!! ", vbInformation
        End If
    End Sub

    Private Sub cmdClear_Click()
        Call FormInitialize
    End Sub

'''    Private Sub cmdReject_Click()   '0-Letter of Authority;1-Letter of Allotment
'''        frmReject.Mode = 12
'''        frmReject.RequestTypeID = cmbType.ListIndex
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''        cmdApprove.Enabled = False
'''    End Sub

    Private Sub cmdSave_Click()
    
         Dim mCnn    As New ADODB.Connection
         Dim objDB   As New clsDB
         Dim mintID  As Variant
         Dim mStatus As Variant
         Dim mArrIn  As Variant
         Dim marrOut As Variant
     
         If Trim(txtAuthorityAllotmentNum.Text) = "" Then
             MsgBox "Enter the Authority/Allotment Number", vbInformation
             Exit Sub
         End If
         If Trim(txtDate.Text) = "" Then
             MsgBox "Enter the Date", vbInformation
             Exit Sub
         End If
         If Trim(txtAmount.Text) = "" Or Trim(txtAmount.Text) <= 0 Then
             MsgBox "Enter the Amount", vbInformation
             Exit Sub
         End If
        
         If cmbType.ListIndex < 0 Then
            MsgBox "Enter the Type", vbInformation
            Exit Sub
         End If
         If cmbType.ListIndex = 0 Then
            mStatus = 1
         ElseIf cmbType.ListIndex = 1 Then
            mStatus = 2
         End If
         If objDB.SetConnection(mCnn) Then
         mintID = IIf(txtAuthorityAllotmentNum.Tag = "", -1, val(txtAuthorityAllotmentNum.Tag))
         mArrIn = Array(mintID, mStatus, _
                            Trim(txtAuthorityAllotmentNum.Text), _
                            Null, _
                            Null, _
                            Trim(txtAmount.Text), _
                            Trim(txtDate.Text), _
                            gbUserID, _
                            0 _
                         )
         
         
         objDB.ExecuteSP "spSaveAllotmentRegister", mArrIn, marrOut, , mCnn, adCmdStoredProc
         MsgBox "Saved Successfully!", vbInformation, "Saankhya"
         Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
          End If
          Call FormInitialize
          Call FillGrid
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FormInitialize
        txtDate.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
        
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then   'gbUserTypeID = 4 Then   'Accounts Officer
            cmdSave.Enabled = False
            cmdApprove.Enabled = True
'            cmdReject.Enabled = True
            VSGrid.Editable = flexEDNone
        Else
            cmdApprove.Enabled = False
'            cmdReject.Enabled = False
            cmdSave.Enabled = True
            'VSGrid.Editable = flexEDKbdMouse
             VSGrid.Editable = flexEDNone
        End If
        'Call fillGrid
        frmAllotmentLetter.PDEMode = 1
        cmbType.AddItem "Letter of Authority"
        cmbType.ItemData(cmbType.NewIndex) = 0
        cmbType.AddItem "Letter of Allotment"
        cmbType.ItemData(cmbType.NewIndex) = 1
        
    End Sub
    Private Sub FormInitialize()
       txtAuthorityAllotmentNum.Text = ""
       txtDate.Text = ""
       txtAmount.Text = ""
       cmdSave.Caption = "Add"
       txtAuthorityAllotmentNum.Tag = ""
   End Sub
     Private Sub Form_Activate()
        Me.Top = 500
        Me.Left = (frmMenu.Width - Me.Width) / 2
     End Sub
    Private Sub FillGrid()
    
    Dim mSQL    As String
    Dim objDB   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim mRowCnt As Integer
    Dim mRecCnt As Integer
    Dim mChk    As Integer
   
    If objDB.SetConnection(mCnn) Then
'''        mSql = "SELECT   * From faAllotmentRegister"
        mSQL = " SELECT faAllotmentRegister.*, suSourceOfFund.vchSourceFundName , faTransactionCategory.vchTransactionCategory, faVouchers.intVoucherNo "
        mSQL = mSQL + " FROM faAllotmentRegister LEFT  JOIN faVouchers ON faAllotmentRegister.intVoucherID = faVouchers.intVoucherID LEFT JOIN"
        mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo LEFT JOIN"
        mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID Left JOIN"
        mSQL = mSQL + " faTransactionCategory ON faAllotmentLetters.intCategoryID = faTransactionCategory.intCategoryID "
        If cmbType.ListIndex = 0 Then
            mSQL = mSQL + " where faAllotmentRegister.tnyTypeID=1"
        ElseIf cmbType.ListIndex = 1 Then
            mSQL = mSQL + " where faAllotmentRegister.tnyTypeID=2"
        End If
        Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        mRecCnt = 1
        VSGrid.Clear 1, 1
        VSGrid.Rows = 1
        While Not (Rec.EOF Or Rec.BOF)
            VSGrid.Rows = VSGrid.Rows + 1
            VSGrid.TextMatrix(mRowCnt, 0) = mRecCnt
            VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
            VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtEntryDate), "", CheckDateInMMM(Rec!dtEntryDate))
            VSGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!tnyTypeID), "", Rec!tnyTypeID)
            VSGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            If VSGrid.TextMatrix(mRowCnt, 10) = 0 Then
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then                 'gbUserTypeID = 3 Then
                    cmdSave.Enabled = True
                    cmdApprove.Enabled = False
                ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then           'gbUserTypeID = 4 Then
                    cmdSave.Enabled = False
                    cmdApprove.Enabled = True
                End If
            End If
            If VSGrid.TextMatrix(mRowCnt, 10) = 1 Then
                cmdSave.Enabled = False
                cmdApprove.Enabled = False
            End If
            If VSGrid.TextMatrix(mRowCnt, 10) = 2 Or VSGrid.TextMatrix(mRowCnt, 10) = 3 Or VSGrid.TextMatrix(mRowCnt, 10) = 4 Or VSGrid.TextMatrix(mRowCnt, 10) = 5 Then
               VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
               VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
               cmdApprove.Enabled = False
               cmdSave.Enabled = False
            End If
            If VSGrid.TextMatrix(mRowCnt, 10) = 2 Or VSGrid.TextMatrix(mRowCnt, 10) = 3 Or VSGrid.TextMatrix(mRowCnt, 10) = 4 Or VSGrid.TextMatrix(mRowCnt, 10) = 5 Then
               'If gbUserTypeID = 3 Then
                    VSGrid.Editable = flexEDNone
                    cmdApprove.Enabled = False
                    cmdSave.Enabled = False
               'End If
            End If
            If Rec!tnyStatus >= 3 Then
                VSGrid.TextMatrix(mRowCnt, 6) = vbChecked
            End If
            If Rec!tnyStatus >= 4 Then
                VSGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            End If
            If Rec!tnyStatus = 5 Then
                VSGrid.TextMatrix(mRowCnt, 8) = vbChecked
            End If
            Rec.MoveNext
            
            mRowCnt = mRowCnt + 1
            mRecCnt = mRecCnt + 1
        Wend
        Rec.Close
    End If
    End Sub
    Private Sub txtAuthorityAllotmentNum_LostFocus()
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
            
            If objDB.SetConnection(mCnn) Then
                If txtAuthorityAllotmentNum.Text <> "" Then
                    mSQL = " select * from faAllotmentRegister "
                    mSQL = mSQL + " where vchAllotmentNo='" & txtAuthorityAllotmentNum.Text & " ' "
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        MsgBox " Allotment Number Already Entered", vbInformation
                        Exit Sub
                    End If
                    Rec.Close
                    mSQL = " select * from faAllotmentLetters "
                    mSQL = mSQL + " where vchAllotmentNo='" & txtAuthorityAllotmentNum.Text & " ' "
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        MsgBox " Allotment Number Already Entered", vbInformation
                        Exit Sub
                    End If
                    Rec.Close
                End If
               
            End If
    End Sub

    Private Sub txtDate_LostFocus()
       If txtDate.Text <> "" Then
            If CheckDateInMMM(txtDate.Text) >= gbTransactionDate Then
                txtDate.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
            Else
                txtDate.Text = Format(CheckDateInMMM(txtDate.Text), "dd/mmm/yyyy")
            End If
       End If
    End Sub
   Private Sub VSGrid_Click()
     Dim mSQL    As String
     Dim objDB   As New clsDB
     Dim mCnn    As New ADODB.Connection
     Dim Rec     As New ADODB.Recordset
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then                 'gbUserTypeID = 3 Then
            If VSGrid.Row > 0 Then
                txtAuthorityAllotmentNum.Text = VSGrid.TextMatrix(VSGrid.Row, 1)
                txtDate.Text = VSGrid.TextMatrix(VSGrid.Row, 2)
                txtAmount.Text = VSGrid.TextMatrix(VSGrid.Row, 5)
                cmdSave.Caption = "Update"
                If objDB.SetConnection(mCnn) Then
                    mSQL = " SELECT intID From faAllotmentRegister WHERE vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "' "
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        If Rec!intID <> "" Then
                             txtAuthorityAllotmentNum.Tag = (Rec!intID)
                        End If
                    End If
                    Rec.Close
                End If
            End If
        End If
        
'''        If VSGrid.TextMatrix(VSGrid.Row, 6) = vbChecked Then
'''         cmdSave.Enabled = False
'''        End If
    End Sub
    Private Sub VSGrid_DblClick()
    Dim mCnn As New ADODB.Connection
    Dim Rec  As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mSQL As String
        mCheckDemand = 1
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then                 'gbUserTypeID = 3 Then
         If VSGrid.Row > 0 Then
        
            If VSGrid.TextMatrix(VSGrid.Row, 10) = 1 Or VSGrid.TextMatrix(VSGrid.Row, 10) = 2 Then
                    If VSGrid.TextMatrix(VSGrid.Row, 9) = 1 Then
                        frmAllotmentLetter.LoadMode = 10
                        frmAllotmentLetter.AuthorityOrAllotment = "Authority"
                        frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Receipt of A/B/C Fund in the Treasury Account"
                        frmAllotmentLetter.txtAllotmentNo = VSGrid.TextMatrix(VSGrid.Row, 1)
                        frmAllotmentLetter.txtAllotmentNo.Enabled = False
                        frmAllotmentLetter.txtAllotmentDate = VSGrid.TextMatrix(VSGrid.Row, 2)
                        frmAllotmentLetter.txtAmountInFigures = VSGrid.TextMatrix(VSGrid.Row, 5)
                        frmAllotmentLetter.txtAmountInFigures.Enabled = False
                        frmAllotmentLetter.CheckDemand = 1
                        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                        mSQL = "SELECT faAllotmentRegister.tnyTypeID, faAllotmentRegister.vchAllotmentNo, faAllotmentRegister.dtLetterDate, faAllotmentRegister.intVoucherID, "
                        mSQL = mSQL + " faAllotmentRegister.fltAmount, faAllotmentLetters.intAllotmentID,faAllotmentRegister.dtEntryDate, faAllotmentRegister.tnyStatus, faTransactionType.vchTransactionType,"
                        mSQL = mSQL + " suSourceOfFund.vchSourceFundName, B.vchAccountHead AS AccountHead, faFunctionaries.vchFunctionary, faFunctions.vchFunction,"
                        mSQL = mSQL + " faTransactionCategory.vchTransactionCategory,B.vchAccountHeadCode "
                        mSQL = mSQL + " FROM faAllotmentRegister INNER JOIN"
                        mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo INNER JOIN"
                        mSQL = mSQL + " faTransactionType ON faAllotmentLetters.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
                        mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID INNER JOIN"
                        mSQL = mSQL + " faFunctionaries ON faAllotmentLetters.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN"
                        mSQL = mSQL + " faFunctions ON faAllotmentLetters.intFunctionID = faFunctions.intFunctionID INNER JOIN"
                        mSQL = mSQL + " faTransactionCategory ON faAllotmentLetters.intCategoryID = faTransactionCategory.intCategoryID LEFT OUTER JOIN"
                        mSQL = mSQL + " faAccountHeads B ON B.intAccountHeadID = faAllotmentLetters.intGrossAccountHeadID"
                        mSQL = mSQL + " WHERE faAllotmentRegister.vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "'"
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            frmAllotmentLetter.cmbTransactionTypes = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                            frmAllotmentLetter.txtAllotmentNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                            frmAllotmentLetter.cmbSource = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                            frmAllotmentLetter.cmbCategory = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                            frmAllotmentLetter.txtFunction = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                            frmAllotmentLetter.txtFunctionary = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                            frmAllotmentLetter.txtAccountHead = IIf(IsNull(Rec!AccountHead), "", Rec!AccountHead)
                            frmAllotmentLetter.txtAccountHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                            frmAllotmentLetter.txtAmountInFigures = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                        End If
                        Rec.Close
                        frmAllotmentLetter.cmdNew.Enabled = False
                        frmAllotmentLetter.Show vbModal
                    ElseIf VSGrid.TextMatrix(VSGrid.Row, 9) = 2 Then
                        frmAllotmentLetter.LoadMode = 10
                        frmAllotmentLetter.AuthorityOrAllotment = "Allotment"
                        frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Allotment of B Fund in the Consolidated Fund in the Treasury"
                        frmAllotmentLetter.txtAllotmentNo = VSGrid.TextMatrix(VSGrid.Row, 1)
                        frmAllotmentLetter.txtAllotmentNo.Enabled = False
                        frmAllotmentLetter.txtAllotmentDate = VSGrid.TextMatrix(VSGrid.Row, 2)
                        frmAllotmentLetter.txtAmountInFigures = VSGrid.TextMatrix(VSGrid.Row, 5)
                        frmAllotmentLetter.txtAmountInFigures.Enabled = False
                        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                        mSQL = "SELECT faAllotmentLetters.vchTreasuryCode,faAllotmentLetters.vchTreasuryName,faAllotmentRegister.tnyTypeID, faAllotmentRegister.vchAllotmentNo, faAllotmentRegister.dtLetterDate, faAllotmentRegister.intVoucherID, "
                        mSQL = mSQL + " faAllotmentRegister.fltAmount, faAllotmentLetters.intAllotmentID,faAllotmentRegister.dtEntryDate, faAllotmentRegister.tnyStatus, faTransactionType.vchTransactionType,"
                        mSQL = mSQL + " suSourceOfFund.vchSourceFundName , A.vchAccountHead Scheme,B.vchAccountHead AccountHead,B.intAccountHeadID,  faFunctionaries.vchFunctionary, faFunctions.vchFunction, B.vchAccountHeadCode,faAllotmentLetters.intSchemeID"
                        mSQL = mSQL + " FROM faAllotmentRegister INNER JOIN"
                        mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo INNER JOIN"
                        mSQL = mSQL + " faTransactionType ON faAllotmentLetters.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
                        mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID INNER JOIN"
                        mSQL = mSQL + " faAccountHeads A ON faAllotmentLetters.intSchemeID = A.intAccountHeadID INNER JOIN"
                        mSQL = mSQL + " faFunctionaries ON faAllotmentLetters.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN"
                        mSQL = mSQL + " faFunctions ON faAllotmentLetters.intFunctionID = faFunctions.intFunctionID LEFT JOIN faAccountHeads B On B.intAccountHEadID = faAllotmentLetters.intGrossAccountHeadID"
                        mSQL = mSQL + " WHERE faAllotmentRegister.vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "'"
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            frmAllotmentLetter.cmbTransactionTypes = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                            frmAllotmentLetter.txtAllotmentNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                            frmAllotmentLetter.cmbSource = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                            frmAllotmentLetter.txtScheme = IIf(IsNull(Rec!Scheme), "", Rec!Scheme)
                            frmAllotmentLetter.txtScheme.Tag = IIf(IsNull(Rec!intSchemeID), "", Rec!intSchemeID)
                            frmAllotmentLetter.txtFunction = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                            frmAllotmentLetter.txtFunctionary = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                            frmAllotmentLetter.txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                            frmAllotmentLetter.txtAccountHead = IIf(IsNull(Rec!AccountHead), "", Rec!AccountHead)
                            frmAllotmentLetter.txtAccountHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                            frmAllotmentLetter.txtAmountInFigures = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            frmAllotmentLetter.txtTreasuryCode = IIf(IsNull(Rec!vchTreasuryCode), "", Rec!vchTreasuryCode)
                            frmAllotmentLetter.txtNameOfTreasury = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
                        End If
                        Rec.Close
                        frmAllotmentLetter.cmdNew.Enabled = False
                        frmAllotmentLetter.Show vbModal
                    End If
                ElseIf VSGrid.TextMatrix(VSGrid.Row, 10) = 3 Then
                    If VSGrid.TextMatrix(VSGrid.Row, 9) = 1 Then
                         frmAllotmentLinkWithVoucher.txtAllotmentNo.Text = VSGrid.TextMatrix(VSGrid.Row, 1)
                         objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                         mSQL = "SELECT faAllotmentRegister.tnyTypeID, faAllotmentRegister.vchAllotmentNo, faAllotmentRegister.dtLetterDate, faAllotmentRegister.intVoucherID, "
                         mSQL = mSQL + " faAllotmentRegister.fltAmount, faAllotmentLetters.intAllotmentID,faAllotmentRegister.dtEntryDate, faAllotmentRegister.tnyStatus, faTransactionType.vchTransactionType,"
                         mSQL = mSQL + " suSourceOfFund.vchSourceFundName, B.vchAccountHead AS AccountHead, faFunctionaries.vchFunctionary, faFunctions.vchFunction,"
                         mSQL = mSQL + " faTransactionCategory.vchTransactionCategory,B.vchAccountHeadCode "
                         mSQL = mSQL + " FROM faAllotmentRegister INNER JOIN"
                         mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo INNER JOIN"
                         mSQL = mSQL + " faTransactionType ON faAllotmentLetters.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
                         mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID INNER JOIN"
                         mSQL = mSQL + " faFunctionaries ON faAllotmentLetters.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN"
                         mSQL = mSQL + " faFunctions ON faAllotmentLetters.intFunctionID = faFunctions.intFunctionID INNER JOIN"
                         mSQL = mSQL + " faTransactionCategory ON faAllotmentLetters.intCategoryID = faTransactionCategory.intCategoryID LEFT OUTER JOIN"
                         mSQL = mSQL + " faAccountHeads B ON B.intAccountHeadID = faAllotmentLetters.intGrossAccountHeadID"
                         mSQL = mSQL + " WHERE faAllotmentRegister.vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "'"
                         Rec.Open mSQL, mCnn
                         If Not (Rec.EOF And Rec.BOF) Then
                            frmAllotmentLinkWithVoucher.txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                            frmAllotmentLinkWithVoucher.txtTransactionType.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                            frmAllotmentLinkWithVoucher.txtSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                            frmAllotmentLinkWithVoucher.txtCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                            frmAllotmentLinkWithVoucher.txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                            frmAllotmentLinkWithVoucher.txtFunctionaries.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                            frmAllotmentLinkWithVoucher.txtAccountHead.Text = IIf(IsNull(Rec!AccountHead), "", Rec!AccountHead)
                            'frmAllotmentLinkWithVoucher = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                            frmAllotmentLinkWithVoucher.txtAmount.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                         End If
                         Rec.Close
                         frmAllotmentLinkWithVoucher.Show vbModal
                         
                    End If
    '''            Else
    '''                MsgBox "The Register is not closed yet", vbInformation
    '''                Exit Sub
                End If
         End If
         '*********
'''         Else
'''            MsgBox "Pending to Enter data", vbInformation, "Saankhya"
         
        End If
   '--------------------------------------------Accounts Officer--------------------------------------------------------------------'
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then                         'gbUserTypeID = 4 Then
            If VSGrid.Row > 0 Then
                If VSGrid.TextMatrix(VSGrid.Row, 10) = 2 Then
                    If VSGrid.TextMatrix(VSGrid.Row, 9) = 1 Then
                        frmAllotmentLetter.LoadMode = 10
                        frmAllotmentLetter.AuthorityOrAllotment = "Authority"
                        frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Receipt of A/B/C Fund in the Treasury Account"
                        frmAllotmentLetter.txtAllotmentNo = VSGrid.TextMatrix(VSGrid.Row, 1)
                        frmAllotmentLetter.txtAllotmentNo.Enabled = False
                        frmAllotmentLetter.txtAllotmentDate = VSGrid.TextMatrix(VSGrid.Row, 2)
                        frmAllotmentLetter.txtAmountInFigures = VSGrid.TextMatrix(VSGrid.Row, 5)
                        frmAllotmentLetter.txtAmountInFigures.Enabled = False
                        frmAllotmentLetter.CheckDemand = 1
                        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                        mSQL = "SELECT faAllotmentRegister.tnyTypeID, faAllotmentRegister.vchAllotmentNo, faAllotmentRegister.dtLetterDate, faAllotmentRegister.intVoucherID, "
                        mSQL = mSQL + " faAllotmentRegister.fltAmount, faAllotmentLetters.intAllotmentID,faAllotmentRegister.dtEntryDate, faAllotmentRegister.tnyStatus, faTransactionType.vchTransactionType,"
                        mSQL = mSQL + " suSourceOfFund.vchSourceFundName, B.vchAccountHead AS AccountHead, faFunctionaries.vchFunctionary, faFunctions.vchFunction,"
                        mSQL = mSQL + " faTransactionCategory.vchTransactionCategory,B.vchAccountHeadCode "
                        mSQL = mSQL + " FROM faAllotmentRegister INNER JOIN"
                        mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo INNER JOIN"
                        mSQL = mSQL + " faTransactionType ON faAllotmentLetters.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
                        mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID INNER JOIN"
                        mSQL = mSQL + " faFunctionaries ON faAllotmentLetters.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN"
                        mSQL = mSQL + " faFunctions ON faAllotmentLetters.intFunctionID = faFunctions.intFunctionID INNER JOIN"
                        mSQL = mSQL + " faTransactionCategory ON faAllotmentLetters.intCategoryID = faTransactionCategory.intCategoryID LEFT OUTER JOIN"
                        mSQL = mSQL + " faAccountHeads B ON B.intAccountHeadID = faAllotmentLetters.intGrossAccountHeadID"
                        mSQL = mSQL + " WHERE faAllotmentRegister.vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "'"
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            frmAllotmentLetter.cmbTransactionTypes = Rec!vchTransactionType
                            frmAllotmentLetter.txtAllotmentNo.Tag = Rec!intAllotmentID
                            frmAllotmentLetter.cmbSource = Rec!vchSourceFundName
                            frmAllotmentLetter.cmbCategory = IIf(IsNull(Rec!vchTransactionCategory), 0, Rec!vchTransactionCategory)
                            frmAllotmentLetter.txtFunction = IIf(IsNull(Rec!vchFunction), 0, Rec!vchFunction)
                            frmAllotmentLetter.txtFunctionary = IIf(IsNull(Rec!vchFunctionary), 0, Rec!vchFunctionary)
                            frmAllotmentLetter.txtAccountHead = IIf(IsNull(Rec!AccountHead), 0, Rec!AccountHead)
                            frmAllotmentLetter.txtAccountHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), 0, Rec!vchAccountHeadCode)
                            frmAllotmentLetter.txtAmountInFigures = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                            
                        End If
                        Rec.Close
                        frmAllotmentLetter.cmdNew.Enabled = False
                        frmAllotmentLetter.cmdSave.Enabled = False
                        frmAllotmentLetter.cmdApprove.Enabled = True
                        frmAllotmentLetter.Show vbModal
                    
                    ElseIf VSGrid.TextMatrix(VSGrid.Row, 9) = 2 Then
                        frmAllotmentLetter.LoadMode = 10
                        frmAllotmentLetter.AuthorityOrAllotment = "Allotment"
                        frmAllotmentLetter.lblDescription.Caption = "Use this form to Record Allotment of B Fund in the Consolidated Fund in the Treasury"
                        frmAllotmentLetter.txtAllotmentNo = VSGrid.TextMatrix(VSGrid.Row, 1)
                        frmAllotmentLetter.txtAllotmentNo.Enabled = False
                        frmAllotmentLetter.txtAllotmentDate = VSGrid.TextMatrix(VSGrid.Row, 2)
                        frmAllotmentLetter.txtAmountInFigures = VSGrid.TextMatrix(VSGrid.Row, 5)
                        frmAllotmentLetter.txtAmountInFigures.Enabled = False
                        frmAllotmentLetter.CheckDemand = 1
                        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                        mSQL = "SELECT faAllotmentLetters.vchTreasuryCode,faAllotmentLetters.vchTreasuryName,faAllotmentRegister.tnyTypeID, faAllotmentRegister.vchAllotmentNo, faAllotmentRegister.dtLetterDate, faAllotmentRegister.intVoucherID, "
                        mSQL = mSQL + " faAllotmentRegister.fltAmount, faAllotmentLetters.intAllotmentID,faAllotmentRegister.dtEntryDate, faAllotmentRegister.tnyStatus, faTransactionType.vchTransactionType,"
                        mSQL = mSQL + " suSourceOfFund.vchSourceFundName , A.vchAccountHead Scheme,B.vchAccountHead AccountHead,B.intAccountHeadID, faFunctionaries.vchFunctionary, faFunctions.vchFunction, B.vchAccountHeadCode,faAllotmentLetters.intSchemeID"
                        mSQL = mSQL + " FROM faAllotmentRegister INNER JOIN"
                        mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo INNER JOIN"
                        mSQL = mSQL + " faTransactionType ON faAllotmentLetters.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
                        mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID INNER JOIN"
                        mSQL = mSQL + " faAccountHeads A ON faAllotmentLetters.intSchemeID = A.intAccountHeadID INNER JOIN"
                        mSQL = mSQL + " faFunctionaries ON faAllotmentLetters.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN"
                        mSQL = mSQL + " faFunctions ON faAllotmentLetters.intFunctionID = faFunctions.intFunctionID LEFT JOIN faAccountHeads B On B.intAccountHEadID = faAllotmentLetters.intGrossAccountHeadID"
                        mSQL = mSQL + " WHERE faAllotmentRegister.vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "'"
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            frmAllotmentLetter.cmbTransactionTypes = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                            frmAllotmentLetter.txtAllotmentNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                            frmAllotmentLetter.cmbSource = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                            frmAllotmentLetter.txtScheme = IIf(IsNull(Rec!Scheme), "", Rec!Scheme)
                            frmAllotmentLetter.txtScheme.Tag = IIf(IsNull(Rec!intSchemeID), "", Rec!intSchemeID)
                            frmAllotmentLetter.txtFunction = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                            frmAllotmentLetter.txtFunctionary = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                            frmAllotmentLetter.txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                            frmAllotmentLetter.txtAccountHead = IIf(IsNull(Rec!AccountHead), "", Rec!AccountHead)
                            frmAllotmentLetter.txtAccountHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                            frmAllotmentLetter.txtAmountInFigures = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            frmAllotmentLetter.txtNameOfTreasury = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
                        End If
                        Rec.Close
                        frmAllotmentLetter.cmdNew.Enabled = False
                        frmAllotmentLetter.cmdSave.Enabled = False
                        frmAllotmentLetter.cmdApprove.Enabled = True
                        frmAllotmentLetter.Show vbModal
                    End If
                ElseIf VSGrid.TextMatrix(VSGrid.Row, 10) = 4 Then
                    If VSGrid.TextMatrix(VSGrid.Row, 9) = 1 Then
                        Call ShowAllotmentDetails
                        '''frmAllotmentLinkWithVoucher.Show vbModal
                        
                    End If
    '''            Else
    '''                MsgBox "The details are not entered", vbInformation
    '''                Exit Sub
                End If
            End If
            '**********
'''            Else
'''                MsgBox "Pending for Approval", vbInformation, "Saankhya"
            
            
        End If
    '------------------------------------------------------------------------------------------------------------------------------
    Call FillGrid
    End Sub
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
         If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
    Private Sub txtAuthorityAllotmentNum_KeyPress(KeyAscii As Integer)
'         If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
'                KeyAscii = 0
'        End If
    End Sub
    Private Sub ShowAllotmentDetails()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSQL As String
        
        frmAllotmentLinkWithVoucher.cmdApprove.Visible = True
        frmAllotmentLinkWithVoucher.txtAllotmentNo.Text = VSGrid.TextMatrix(VSGrid.Row, 1)
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = "SELECT faAllotmentRegister.tnyTypeID, faAllotmentRegister.vchAllotmentNo, faAllotmentRegister.dtLetterDate, faAllotmentRegister.intVoucherID, "
        mSQL = mSQL + " faAllotmentRegister.fltAmount, faAllotmentLetters.intAllotmentID,faAllotmentRegister.dtEntryDate, faAllotmentRegister.tnyStatus, faTransactionType.vchTransactionType,"
        mSQL = mSQL + " suSourceOfFund.vchSourceFundName, B.vchAccountHead AS AccountHead, faFunctionaries.vchFunctionary, faFunctions.vchFunction,"
        mSQL = mSQL + " faTransactionCategory.vchTransactionCategory,B.vchAccountHeadCode "
        mSQL = mSQL + " FROM faAllotmentRegister INNER JOIN"
        mSQL = mSQL + " faAllotmentLetters ON faAllotmentRegister.vchAllotmentNo = faAllotmentLetters.vchAllotmentNo INNER JOIN"
        mSQL = mSQL + " faTransactionType ON faAllotmentLetters.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
        mSQL = mSQL + " suSourceOfFund ON faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID INNER JOIN"
        mSQL = mSQL + " faFunctionaries ON faAllotmentLetters.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN"
        mSQL = mSQL + " faFunctions ON faAllotmentLetters.intFunctionID = faFunctions.intFunctionID INNER JOIN"
        mSQL = mSQL + " faTransactionCategory ON faAllotmentLetters.intCategoryID = faTransactionCategory.intCategoryID LEFT OUTER JOIN"
        mSQL = mSQL + " faAccountHeads B ON B.intAccountHeadID = faAllotmentLetters.intGrossAccountHeadID"
        mSQL = mSQL + " WHERE faAllotmentRegister.vchAllotmentNo = '" & VSGrid.TextMatrix(VSGrid.Row, 1) & "'"
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            frmAllotmentLinkWithVoucher.txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            frmAllotmentLinkWithVoucher.txtTransactionType.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
            frmAllotmentLinkWithVoucher.txtSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            frmAllotmentLinkWithVoucher.txtCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
            frmAllotmentLinkWithVoucher.txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            frmAllotmentLinkWithVoucher.txtFunctionaries.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            frmAllotmentLinkWithVoucher.txtAccountHead.Text = IIf(IsNull(Rec!AccountHead), "", Rec!AccountHead)
            'frmAllotmentLinkWithVoucher = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
            frmAllotmentLinkWithVoucher.txtAmount.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        End If
        Rec.Close
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT faVouchers.intVoucherID, faTransactionType.vchTransactionType, faVouchers.intVoucherNo, faVouchers.vchInstrumentNo, faVouchers.fltAmount, "
            mSQL = mSQL + " faVouchers.vchBank , faInstrumentTypes.vchInstrumentType, faVouchers.dtDate FROM faVouchers INNER JOIN"
            mSQL = mSQL + " faTransactionType ON faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID INNER JOIN"
            mSQL = mSQL + " faInstrumentTypes ON faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
            mSQL = mSQL + " Where faVouchers.intVoucherNo = " & VSGrid.TextMatrix(VSGrid.Row, 7) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmAllotmentLinkWithVoucher.txtVoucherNoList.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                frmAllotmentLinkWithVoucher.txtVoucherDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                frmAllotmentLinkWithVoucher.TxtInstrumentType.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                frmAllotmentLinkWithVoucher.txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                frmAllotmentLinkWithVoucher.txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                frmAllotmentLinkWithVoucher.txtVoucherTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                frmAllotmentLinkWithVoucher.txtVoucherAmount.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            End If
            Rec.Close
        End If
        frmAllotmentLinkWithVoucher.Show vbModal
    End Sub
