VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmChequeRegister 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHEQUE REGISTER"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSubmitted 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "Submitted"
      Height          =   285
      Left            =   10170
      TabIndex        =   13
      Top             =   180
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000009&
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
      Height          =   405
      Left            =   5910
      TabIndex        =   12
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4410
      TabIndex        =   11
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H80000009&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8610
      TabIndex        =   10
      Top             =   960
      Width           =   1035
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4200
      Left            =   30
      TabIndex        =   9
      Top             =   1560
      Width           =   11205
      _cx             =   19764
      _cy             =   7408
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
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChequeRegister.frx":0000
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   90
      Picture         =   "frmChequeRegister.frx":01AA
      ScaleHeight     =   1305
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   105
      Width           =   1440
   End
   Begin VB.CommandButton cmdSearchAccountHeadCode 
      Caption         =   "..."
      Height          =   285
      Left            =   8100
      TabIndex        =   7
      Top             =   1020
      Width           =   330
   End
   Begin VB.TextBox txtAccountHead 
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3420
      TabIndex        =   6
      Top             =   1035
      Width           =   4680
   End
   Begin VB.TextBox txtToDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5055
      TabIndex        =   4
      Top             =   675
      Width           =   1245
   End
   Begin VB.OptionButton optIssued 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Issued"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5205
      TabIndex        =   1
      Top             =   210
      Width           =   1290
   End
   Begin VB.OptionButton optReceived 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Received"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   0
      Top             =   210
      Width           =   1125
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3420
      TabIndex        =   3
      Top             =   675
      Width           =   1290
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reconciled"
      Height          =   195
      Left            =   10620
      TabIndex        =   17
      Top             =   6165
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Submitted"
      Height          =   195
      Left            =   10620
      TabIndex        =   16
      Top             =   5895
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   11700
      TabIndex        =   15
      Top             =   5895
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00446FD5&
      Height          =   195
      Left            =   11700
      TabIndex        =   14
      Top             =   6165
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head"
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
      Left            =   2340
      TabIndex        =   5
      Top             =   1095
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   3015
      TabIndex        =   2
      Top             =   720
      Width           =   360
   End
End
Attribute VB_Name = "frmChequeRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
        vsGrid.Clear 1
    End Sub
    Private Sub FillGrid()
        Dim mCon            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim objdb           As New clsDB
        Dim mSql            As String
        Dim mVoucherType    As Integer
        Dim mFromDate       As String
        Dim mToDate         As String
        Dim mAccountID      As Integer
        vsGrid.Clear 1
        vsGrid.Rows = 1
        If optReceived.value = True Then
            mVoucherType = 10
        ElseIf optIssued.value = True Then
            mVoucherType = 20
        Else
            mVoucherType = 10
        End If
        
        If txtFrom.Text = "" Then
            MsgBox "Enter date"
        Else
            mFromDate = txtFrom.Text
        End If
        mToDate = txtToDate.Text
        If Trim(txtAccountHead.Tag) <> "" Then
            mAccountID = txtAccountHead.Tag
        Else
            mAccountID = 0
        End If
        
        objdb.CreateNewConnection mCon, enuSourceString.Saankhya
        If chkSubmitted.value Then
            mSql = "Select intVoucherNo,dtDate,vchInstrumentNo,fltAmount,tnyReconciled, vchInstrumentType,dtInstrumentDate, vchBank, vchBankPlace, faVouchers.intVoucherID, dtDate,vchName ,dtChequeRealiseDate"
            mSql = mSql + " From faVouchers INNER JOIN faInstrumentTypes  ON faVouchers.intInstrumentTypeID =faInstrumentTypes.intInstrumentTypeID "
            mSql = mSql + " left Join faVoucherAddress On faVoucherAddress.intVoucherID=faVouchers.intVoucherID"
            mSql = mSql + " Where faVouchers.intInstrumentTypeID <> 1  And tnyVoucherTypeID = " & mVoucherType
            mSql = mSql + " AND dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "' AND intKeyID1=" & mAccountID & "  AND dtChequeRealiseDate is not Null"
            mSql = mSql + " Union All"
            mSql = mSql + " Select intVoucherNo,dtDate,vchInstrumentNo,fltAmount,tnyReconciled, vchInstrumentType,dtInstrumentDate, vchBank,"
            mSql = mSql + " vchBankPlace , faVouchers.intVoucherID, dtDate, vchName, dtChequeRealiseDate"
            mSql = mSql + " From faVouchers INNER JOIN faInstrumentTypes"
            mSql = mSql + " ON faVouchers.intInstrumentTypeID =faInstrumentTypes.intInstrumentTypeID  left Join faVoucherAddress"
            mSql = mSql + " On faVoucherAddress.intVoucherID=faVouchers.intVoucherID"
            mSql = mSql + " Where faVouchers.intVoucherID"
            mSql = mSql + " in (Select intVoucherID From faVoucherChild Where intAccountHeadID=" & mAccountID & "  "
            If mVoucherType = 20 Then
                mSql = mSql + " And tnyDebitOrCredit=1"
            End If
            mSql = mSql + " And tnyVoucherTypeID = 30 )"
            mSql = mSql + " And faVouchers.intInstrumentTypeID <> 1  "
            mSql = mSql + " AND dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "' AND dtChequeRealiseDate is not Null"
            mSql = mSql + " Order By dtDate"
            vsGrid.ColHidden(7) = False
            'vsGrid.Editable = flexEDNone
        Else
            mSql = "Select intVoucherNo,dtDate,vchInstrumentNo,fltAmount,tnyReconciled, vchInstrumentType,dtInstrumentDate, vchBank, vchBankPlace, faVouchers.intVoucherID, dtDate,vchName,dtChequeRealiseDate "
            mSql = mSql + " From faVouchers INNER JOIN faInstrumentTypes  ON faVouchers.intInstrumentTypeID =faInstrumentTypes.intInstrumentTypeID "
            mSql = mSql + " left Join faVoucherAddress On faVoucherAddress.intVoucherID=faVouchers.intVoucherID"
            mSql = mSql + " Where faVouchers.intInstrumentTypeID <> 1  And tnyVoucherTypeID = " & mVoucherType
            mSql = mSql + " AND dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "' AND intKeyID1=" & mAccountID & "  AND dtChequeRealiseDate is Null"
             mSql = mSql + " Union All"
            mSql = mSql + " Select intVoucherNo,dtDate,vchInstrumentNo,fltAmount,tnyReconciled, vchInstrumentType,dtInstrumentDate, vchBank,"
            mSql = mSql + " vchBankPlace , faVouchers.intVoucherID, dtDate, vchName, dtChequeRealiseDate"
            mSql = mSql + " From faVouchers INNER JOIN faInstrumentTypes"
            mSql = mSql + " ON faVouchers.intInstrumentTypeID =faInstrumentTypes.intInstrumentTypeID  left Join faVoucherAddress"
            mSql = mSql + " On faVoucherAddress.intVoucherID=faVouchers.intVoucherID"
            mSql = mSql + " Where faVouchers.intVoucherID"
            mSql = mSql + " in (Select intVoucherID From faVoucherChild Where intAccountHeadID=" & mAccountID & "  "
            If mVoucherType = 20 Then
                mSql = mSql + " And tnyDebitOrCredit=1"
            End If
            mSql = mSql + " And tnyVoucherTypeID = 30 )"
            mSql = mSql + " And faVouchers.intInstrumentTypeID <> 1  "
            mSql = mSql + " AND dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'  AND dtChequeRealiseDate is Null"
            mSql = mSql + " Order By dtDate"
            vsGrid.ColHidden(7) = True
            vsGrid.Editable = flexEDKbdMouse
        End If
        Set Rec = objdb.ExecuteSP(mSql, , , , mCon, adCmdText)
        If Not (Rec.EOF Or Rec.BOF) Then
            While Not (Rec.EOF Or Rec.BOF)
                 vsGrid.AddItem IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                 If IIf(IsNull(Rec!dtChequeRealiseDate), 0, Rec!dtChequeRealiseDate) <> 0 Then
                    'vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Cols - 1) = &HD2AE9E
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, vsGrid.Cols - 1) = &HC0C0C0
                 End If
                 If (IIf(IsNull(Rec!tnyReconciled), 0, Rec!tnyReconciled)) > 0 Then
                     vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = "Reconciled"
    '                vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = "Submitted to bank"
                 End If
                 If (IIf(IsNull(Rec!dtChequeRealiseDate), 0, Rec!dtChequeRealiseDate)) <> 0 Then
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = "Submitted to bank"
                 End If
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                 
    ''             If (IIf(IsNull(Rec!tnyReconciled), "", Rec!tnyReconciled)) = 3 Then
    ''                vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 9) = flexcpChecked
                 If (IIf(IsNull(Rec!tnyReconciled), 0, Rec!tnyReconciled)) > 0 Then
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 9) = flexcpChecked
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 9, vsGrid.Rows - 1, 9) = &H446FD5
                 Else
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 9) = Checked
                 End If
                 
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 10) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 11) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                 vsGrid.TextMatrix(vsGrid.Rows - 1, 12) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                 Rec.MoveNext
             Wend
        Else
            MsgBox "No Record  Exists", vbApplicationModal
        End If
    End Sub
'''''    Private Sub FillGrid()s
'''''        Dim mCon As New ADODB.Connection
'''''        Dim Rec As New ADODB.Recordset
'''''        Dim objDB As New clsDB
'''''        Dim mSQL As String
'''''        Dim mInstrumentNo As String
'''''        vsGrid.Rows = 1
'''''        With vsGrid
'''''           .Rows = 1
'''''           .OutlineBar = flexOutlineBarComplete
'''''           .Editable = flexEDKbdMouse
'''''           .ColAlignment(-1) = flexAlignLeftCenter
'''''
'''''           objDB.CreateNewConnection mCon, enuSourceString.Saankhya
'''''           mSQL = "Select vchInstrumentNo, dtInstrumentDate, vchBank, vchBankPlace, intVoucherNo, dtDate From faVouchers Where tnyVoucherTypeID = 10 And intInstrumentTypeID = 5 Order by vchInstrumentNo"
'''''           Set Rec = objDB.ExecuteSP(mSQL, , , , mCon, adCmdText)
'''''           While Not (Rec.EOF Or Rec.BOF)
'''''                If mInstrumentNo <> Rec!vchInstrumentNo Then
'''''                    .AddItem IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''                    .Cell(flexcpBackColor, .Rows - 1, 0) = &HE0E0E0
'''''                    '.Cell(flexcpChecked, .Rows - 1, 1) = vbChecked
'''''                    .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''                   ' .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rec!dtInstrumentDate), "", DdMmmYy(Rec!dtInstrumentDate))
'''''                    .IsSubtotal(.Rows - 1) = True
'''''                    .RowOutlineLevel(.Rows - 1) = 0
'''''                    mInstrumentNo = Rec!vchInstrumentNo
'''''               Else
'''''                    .AddItem IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''                    .Cell(flexcpBackColor, .Rows - 1, 0) = &HE0E0E0
'''''                    '.Cell(flexcpChecked, .Rows - 1, 1) = vbChecked
'''''                    .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''                    '.TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rec!dtInstrumentDate), "", DdMmmYy(Rec!dtInstrumentDate))
'''''                    .IsSubtotal(.Rows - 1) = True
'''''                    .RowOutlineLevel(.Rows - 1) = 1
'''''
'''''               End If
'''''               Rec.MoveNext
'''''           Wend
'''''           .AutoSizeMouse = True
'''''        End With
'''''    End Sub
    
    Private Sub chkSubmitted_Click()
        Call FillGrid
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdSave_Click()
        Dim mCon        As ADODB.Connection
        Dim Rec         As ADODB.Recordset
        Dim objdb       As New clsDB
        Dim mCnt        As Integer
        Dim arrInput    As Variant
        Dim mBank       As String
        Dim mSql        As String
        If chkSubmitted.value Then
            MsgBox "Already Submitted To Bank"
            Exit Sub
        End If
        
        FileInitialize
                
       ' gbFileNO = FreeFile
       ' Updates tnyReconciled as 3 indicates Cheques submitted to the Bank
        objdb.CreateNewConnection mCon, enuSourceString.Saankhya
       
        
        '-------To Display Cheque Register details
        
        Print #gbFileNO,
        Print #gbFileNO, "-----------------------------------------------------------------------------"
        Print #gbFileNO, "_____________________________  CHEQUE REGISTER  _____________________________"
        
        mBank = mID(txtAccountHead.Text, 10, Len(txtAccountHead.Text)) & " (" & Left(txtAccountHead.Text, 9) & ") "
        If optReceived.value = True Then
            Print #gbFileNO, "Cheque Received By"; mBank
        Else
            Print #gbFileNO, "Cheque Issued By"; mBank
        End If
        
        Print #gbFileNO, "-----------------------------------------------------------------------------"
        Print #gbFileNO, "Receipt date:-From "; txtFrom.Text; " To"; txtToDate.Text
        Print #gbFileNO, "-----------------------------------------------------------------------------"
        Print #gbFileNO, "Receipt No"; "  "; "Cheque No"; "  "; "Cheque Date"; "  "; "Bank Name/Place"; " PartyName   "; "   "; PadL("Amount", 9)
        'Print #gbFileNO, PadR("ReceiptNo", 9);
        Print #gbFileNO, "-----------------------------------------------------------------------------"
        
        Print #gbFileNO,
        If chkSubmitted.value Then
            For mCnt = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, mCnt, 9) = flexChecked Then
''                    mSql = "UPDATE faVouchers SET tnyReconciled = 0 WHERE intVoucherID=" & vsGrid.TextMatrix(mCnt, 10)
''                    objDb.ExecuteSP mSql, , , , mCon, adCmdText
                Else
                    arrInput = Array(vsGrid.TextMatrix(mCnt, 10))
                    Print #gbFileNO, PadR(vsGrid.TextMatrix(mCnt, 0), 5); PadR(vsGrid.TextMatrix(mCnt, 1), 10); PadR(vsGrid.TextMatrix(mCnt, 2), 10); "  ";
                    Print #gbFileNO, PadR(vsGrid.TextMatrix(mCnt, 4) & "    /    " & vsGrid.TextMatrix(mCnt, 5), 27); PadR(vsGrid.TextMatrix(mCnt, 3), 10); PadL(Format(val(vsGrid.TextMatrix(mCnt, 8)), "0.00"), 13)
                End If
            Next
        Else
            For mCnt = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, mCnt, 9) = flexChecked Then
                    Print #gbFileNO, PadR(vsGrid.TextMatrix(mCnt, 0), 12); PadR(vsGrid.TextMatrix(mCnt, 1), 10); PadR(vsGrid.TextMatrix(mCnt, 2), 10); "  ";
                    Print #gbFileNO, PadR(vsGrid.TextMatrix(mCnt, 4) & "/" & vsGrid.TextMatrix(mCnt, 5), 27); PadR(vsGrid.TextMatrix(mCnt, 3), 10); PadL(Format(val(vsGrid.TextMatrix(mCnt, 8)), "0.00"), 13)
                    mSql = "UPDATE faVouchers SET dtChequeRealiseDate = '" & Format(gbTransactionDate, "dd/mmm/yy") & "' WHERE intVoucherID=" & vsGrid.TextMatrix(mCnt, 10)
                    objdb.ExecuteSP mSql, , , , mCon, adCmdText
                End If
            Next
        End If
        Close #gbFileNO
        ShellPad
        
    End Sub

    Private Sub cmdSearch_Click()
        If txtAccountHead.Text = "" Then
            MsgBox "Select the Bank", vbApplicationModal
            Exit Sub
        End If
        Call FillGrid
    End Sub

    Private Sub cmdSearchAccountHeadCode_Click()
        Dim mSql As String
        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = " & faBank
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
        txtAccountHead.SetFocus
    End Sub

    Private Sub Form_Load()
           Call FormInitialize
    End Sub

    Private Sub txtAccountHead_GotFocus()
        If gbSearchStr <> "" Then
            txtAccountHead.Text = gbSearchStr
            txtAccountHead.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
        txtAccountHead.SelStart = 0
        txtAccountHead.SelLength = Len(txtAccountHead)
    End Sub
    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
        End If
    End Sub
    Private Sub txtFrom_LostFocus()
        If Not IsDate(txtFrom.Text) Then
            txtFrom.Text = DdMmmYy(gbTransactionDate)
        Else
            txtFrom.Text = CheckDateInMMM(Trim(txtFrom))
        End If
        If Not IsDate(txtToDate) Then
            txtToDate.Text = CheckDateInMMM(Trim(txtFrom))
        End If
        If CDate(txtFrom.Text) Then
            If CDate(txtToDate.Text) Then
                If CDate(txtFrom.Text) > CDate(txtToDate.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtFrom.Text = CheckDateInMMM(txtFrom.Text)
        End If
    
    End Sub
    Private Sub txtToDate_LostFocus()
        If Not IsDate(txtToDate.Text) Then
            txtToDate.Text = DdMmmYy(gbTransactionDate)
        Else
            txtToDate.Text = CheckDateInMMM(Trim(txtToDate))
        End If
        If CDate(txtToDate.Text) Then
            If CDate(txtFrom.Text) Then
                If CDate(txtFrom.Text) > CDate(txtToDate.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtToDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtToDate.Text = CheckDateInMMM(txtFrom.Text)
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Col = 9 Then
            If vsGrid.TextMatrix(Row, 7) = "Reconciled" Then
                Cancel = True
            End If
        End If
    End Sub

