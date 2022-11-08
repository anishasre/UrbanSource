VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmOBCashBook 
   BorderStyle     =   0  'None
   Caption         =   "OpeningCashBook"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14325
   Icon            =   "frmOBCashBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4545
      TabIndex        =   7
      Top             =   6210
      Width           =   825
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4515
      Left            =   270
      TabIndex        =   6
      Top             =   1080
      Width           =   9645
      _cx             =   17013
      _cy             =   7964
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
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOBCashBook.frx":1CCA
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
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      TabIndex        =   2
      Top             =   5580
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5445
      TabIndex        =   1
      Top             =   6210
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3645
      TabIndex        =   0
      Top             =   6210
      Width           =   825
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10125
      TabIndex        =   18
      Top             =   2115
      Width           =   195
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10170
      TabIndex        =   17
      Top             =   1350
      Width           =   195
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "If a Bank is not defined,define it (Go Administration->BankAccounts)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   750
      Left            =   10395
      TabIndex        =   16
      Top             =   2070
      Width           =   3570
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify the correctness of the Cash/Bank/Treasury Balances"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   825
      Left            =   10395
      TabIndex        =   15
      Top             =   1260
      Width           =   3795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10305
      TabIndex        =   14
      Top             =   5535
      Width           =   330
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10305
      TabIndex        =   13
      Top             =   5265
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank defined"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10665
      TabIndex        =   12
      Top             =   5265
      Width           =   2580
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank not Dfined"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10665
      TabIndex        =   11
      Top             =   5580
      Width           =   2580
   End
   Begin VB.Label lblOB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4830
      Left            =   9990
      TabIndex        =   10
      Top             =   1080
      Width           =   4245
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance Sheet Amount of Cash/bank/treasury  "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   45
      TabIndex        =   9
      Top             =   585
      Width           =   10500
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   45
      Picture         =   "frmOBCashBook.frx":1DB0
      Stretch         =   -1  'True
      Top             =   540
      Width           =   14310
   End
   Begin VB.Label lblOpening 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance Already Verified"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   450
      TabIndex        =   8
      Top             =   5625
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Cash book As On"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   405
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7140
      TabIndex        =   3
      Top             =   5625
      Width           =   975
   End
End
Attribute VB_Name = "frmOBCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private mvarOpCLType As Integer '0-Opening,1-Closing
    Private Sub cmdClose_Click()
        If MsgBox("You haven't finished the Opening Cash Book Wizard, are you sure you want to quit?   ", vbQuestion + vbYesNo, "Close Wizard") = vbYes Then
            Unload Me
            frmOpeningWizard.cmdCancel_Click
        End If
    End Sub

    Private Sub cmdNext_Click()
        Me.Hide
        Unload Me
        frmOpeningWizard.cmdNext_Click
    End Sub
 
    Private Sub cmdSave_Click()
        Dim mAccID          As Integer
        Dim mAccCode        As String
        Dim mAmtOp            As Double
        Dim mAmtCL            As Double
        Dim mCnt            As Integer
        Dim mOBCashBookID   As Integer
        Dim mSql            As String
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim objAcc          As New clsAccounts
        Dim arrIn           As Variant
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        For mCnt = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCnt, 0) <> "" Then
                If vsGrid.TextMatrix(mCnt, 4) = "r" Then
                    MsgBox "Bank/Treasury not Defined for the Account " & vbNewLine & vsGrid.TextMatrix(mCnt, 2), vbApplicationModal
                    Exit Sub
                End If
            End If
        Next
        mSql = "Delete FROM faOBCashBook "
        mCnn.Execute (mSql)
        For mCnt = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCnt, 0) <> "" Then
                mOBCashBookID = -1
                mAccID = vsGrid.TextMatrix(mCnt, 0)
                objAcc.SetAccountID (mAccID)
                mAccCode = objAcc.AccountCode
                mAmtOp = val(vsGrid.TextMatrix(mCnt, 3))
                mAmtCL = val(vsGrid.TextMatrix(mCnt, 6))
                arrIn = Array(mOBCashBookID, mAccID, mAccCode, mAmtOp, mAmtCL)
                objdb.ExecuteSP "spSaveOBCashBook", arrIn, , , mCnn, adCmdStoredProc
            End If
        Next
        MsgBox "Successfully Verified..", vbApplicationModal
        cmdSave.Enabled = False
        cmdNext.Enabled = True
        frmOpeningWizard.chkOB.value = vbChecked
        Call FillOBcashBook
        Me.Hide
        frmOpeningWizard.FrameNo = 2
        frmOpeningWizard.cmdNext_Click
        Unload Me
    End Sub
    Private Sub Form_Load()
        Call FillOBcashBook
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSave.Enabled = False
            cmdNext.Enabled = True
        End If
    End Sub

    Private Sub FillOBcashBook()
        Dim mSql             As String
        Dim objdb            As New clsDB
        Dim mCnn             As New ADODB.Connection
        Dim Rec              As New ADODB.Recordset
        Dim mCnt             As Integer
        Dim mVStatus         As Boolean
        Dim mBank           As Boolean
        Dim mOBVerified     As Double
        mVStatus = True
        mBank = True
        mSql = "Select dtRPOpeningDate From faConfig"
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            txtDate.Text = IIf(IsNull(Rec!dtRPOpeningDate), "", Rec!dtRPOpeningDate)
        End If
        Rec.Close
        mSql = " Select faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode,"
        mSql = mSql + " faAccountHeads.vchAccountHead,"
        mSql = mSql + " Case When faVoucherChild.tnyDebitOrCredit=0 Then faVoucherChild.fltAmount*-1 "
        mSql = mSql + " When faVoucherChild.tnyDebitOrCredit=1 Then faVoucherChild.fltAmount End fltAmount,"
        mSql = mSql + " Case When faAccountHeads.intGroupID=2 then("
        mSql = mSql + " case When isNull(faBanks.intAccountHeadID,0)=0 then 'r' Else 'a' End ) Else 'a' End BankStatus"
        mSql = mSql + " From faVouchers"
        mSql = mSql + " INNER JOIN faVoucherChild ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
        mSql = mSql + " INNER JOIN faAccountHeads ON faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
        mSql = mSql + " LEFT JOIN faBanks ON faBanks.intAccountHeadID=faAccountHeads.intAccountHeadID"
        mSql = mSql + " Where faVouchers.intTransactionTypeID =3000 AND faAccountHeads.intGroupID in (1,2)"
        mSql = mSql + " ORDER BY faAccountHeads.intAccountHeadID DESC"
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        vsGrid.Rows = 1
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.Rows = Rec.RecordCount + 1
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColSel = 4
            vsGrid.RowSel = vsGrid.Rows - 1
            mSql = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = mSql
            vsGrid.Cell(flexcpFontName, 1, 4, vsGrid.Rows - 1, 4) = "Webdings"
        End If
        Rec.Close
        Call Calculate
        
        
        For mCnt = 1 To vsGrid.Rows - 1
           If vsGrid.TextMatrix(mCnt, 4) = "r" Then
                vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, vsGrid.Cols - 1) = &H8080FF
           End If
        Next
        
        If vsGrid.FindRow("r", , 4, 1, 1) > 0 Then
'            vsGrid.Cell(flexcpBackColor, vsGrid.FindRow("r", , 4, 1, 1), 0, vsGrid.FindRow("r", , 4, 1, 1), vsGrid.Cols - 1) = &H8080FF
            cmdNext.Enabled = False
            mBank = False
        End If
        If mBank = False Then
            lblOpening.Visible = True
            lblOpening.Caption = "Bank/trasurey Not Defined.."
            cmdSave.Enabled = False
            cmdNext.Enabled = False
            Exit Sub
        End If
        
        mSql = ""
        mSql = "Select * From faOBCashBook " 'Where isNull(fltOpening,0)<>0 "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                Dim mAccID As Integer
                mAccID = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
                If vsGrid.FindRow(mAccID, , 0, 1, 1) > 0 Then
                    vsGrid.TextMatrix(vsGrid.FindRow(mAccID, , 0, 1, 1), 5) = IIf(IsNull(Rec!fltOpening), 0, Rec!fltOpening)
                    vsGrid.TextMatrix(vsGrid.FindRow(mAccID, , 0, 1, 1), 6) = IIf(IsNull(Rec!fltClosing), 0, Rec!fltClosing)
                End If
                Rec.MoveNext
            Wend
            frmOpeningWizard.chkOB.value = vbChecked
            cmdSave.Enabled = False
        End If
        Rec.Close
        
        mSql = "SELECT SUM(fltOpening) AS OB FROM  faOBCashBook"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mOBVerified = IIf(IsNull(Rec!OB), 0, Rec!OB)
            If mOBVerified <> val(txtTotal.Text) Then
                lblOpening.Visible = True
                lblOpening.Caption = "Amount Modified in Opening balance... Please Do Undo voucher and Verify"
                lblOpening.WordWrap = True
                cmdSave.Enabled = True
                Exit Sub
            End If
        End If
        Rec.Close
        For mCnt = 1 To vsGrid.Rows - 1
           If val(vsGrid.TextMatrix(mCnt, 3)) <> val(vsGrid.TextMatrix(mCnt, 5)) Then
               mVStatus = False
               If frmOpeningWizard.mFreeze <> 1 Then
                cmdSave.Enabled = True
               End If
               Exit For
           End If
        Next
        If mVStatus = False Then
            lblOpening.Visible = True
            lblOpening.Caption = "Please Verify Opening Balance.."
            If frmOpeningWizard.mFreeze <> 1 Then
                cmdSave.Enabled = True
            End If
            cmdNext.Enabled = False
            Exit Sub
        Else
            cmdNext.Enabled = True
            lblOpening.Visible = True
            'If frmOpeningWizard.mFreeze <> 1 Then
                cmdSave.Enabled = False
            'End If
        End If
    End Sub
    Private Sub txtTotal_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub
    
    Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
        If Shift = vbCtrlMask And (Chr(KeyCode) = "v" Or Chr(KeyCode) = "V") Then
            KeyCode = 0
        End If
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
        End If
    End Sub
    Private Sub txtTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then
             txtTotal.Locked = True
        End If
    End Sub
    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Col = 4 Then
            vsGrid.Editable = flexEDKbdMouse
            If vsGrid.TextMatrix(Row, 4) = "r" Then
              MsgBox "Please Go Administration->BankAccounts .. then Define Account Head"
              Exit Sub
            End If
        End If
    End Sub
    Private Sub vsGrid_Click()
        If vsGrid.Col = 4 Then
            vsGrid.Editable = flexEDKbdMouse
        End If
    End Sub

    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Calculate
    End Sub
    Private Sub Calculate()
        Dim mCnt    As Integer
        Dim mTotal  As Double
        mTotal = 0
        For mCnt = 1 To vsGrid.Rows - 1
            mTotal = mTotal + val(vsGrid.TextMatrix(mCnt, 3))
        Next
        txtTotal.Text = Format(mTotal, "#.00")
    End Sub

    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        KeyAscii = 0
    End Sub
     Public Property Let OpCLType(mData As Integer)
        mvarOpCLType = mData
    End Property
