VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfContraEntries 
   BackColor       =   &H00EDF7F7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                 L i s t   o f   C o n t r a   E n t r i e s"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListOfContraEntries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkNotApproved 
      BackColor       =   &H00EDF7F7&
      Caption         =   "Not Approved"
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   10170
      TabIndex        =   23
      Top             =   6435
      Width           =   1590
   End
   Begin VB.CheckBox chkOthers 
      BackColor       =   &H00EDF7F7&
      Caption         =   "All the Others"
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   10170
      TabIndex        =   22
      Top             =   7110
      Width           =   1590
   End
   Begin VB.CheckBox chkApproved 
      BackColor       =   &H00EDF7F7&
      Caption         =   "Approved"
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   10170
      TabIndex        =   21
      Top             =   6750
      Width           =   1590
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11835
      TabIndex        =   20
      Top             =   6165
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search criteria"
      Height          =   1635
      Left            =   45
      TabIndex        =   5
      Top             =   6075
      Width           =   10005
      Begin VB.TextBox txtDescription 
         Height          =   405
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   4200
      End
      Begin VB.ComboBox cmbTransactionTypes 
         Height          =   360
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   270
         Width           =   4200
      End
      Begin VB.ComboBox cmbInstrumentTypes 
         Height          =   360
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   675
         Width           =   4200
      End
      Begin VB.TextBox txtVoucherNo 
         Height          =   360
         Left            =   1305
         TabIndex        =   10
         Top             =   675
         Width           =   1410
      End
      Begin VB.TextBox txtAmount 
         Height          =   360
         Left            =   1305
         TabIndex        =   9
         Top             =   1080
         Width           =   1410
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   330
         Left            =   1305
         TabIndex        =   6
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   17498113
         CurrentDate     =   40435
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   330
         Left            =   2745
         TabIndex        =   7
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   17498113
         CurrentDate     =   40435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#: Temperary Voucher Numbers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   2835
         TabIndex        =   24
         Top             =   1395
         Width           =   2400
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   300
         Left            =   2610
         TabIndex        =   19
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   240
         Left            =   4500
         TabIndex        =   18
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         Height          =   240
         Left            =   4140
         TabIndex        =   16
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Type"
         Height          =   240
         Left            =   4305
         TabIndex        =   15
         Top             =   765
         Width           =   1425
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "# Voucher No"
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   765
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   240
         Left            =   495
         TabIndex        =   11
         Top             =   1125
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   240
         Left            =   720
         TabIndex        =   8
         Top             =   315
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11835
      TabIndex        =   3
      Top             =   7380
      Width           =   1230
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11835
      TabIndex        =   2
      Top             =   6975
      Width           =   1230
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11835
      TabIndex        =   1
      Top             =   6570
      Width           =   1230
   End
   Begin WinXPC_Engine.WindowsXPC winXpc 
      Left            =   180
      Top             =   7830
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5730
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   13065
      _cx             =   23045
      _cy             =   10107
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorFixed  =   15595511
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483626
      ForeColorSel    =   4194368
      BackColorBkg    =   -2147483639
      BackColorAlternate=   -2147483633
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   19
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmListOfContraEntries.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   13065
   End
End
Attribute VB_Name = "frmListOfContraEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public mGridRow As Integer
    Private Sub chkApproved_Click()
        If chkApproved.value = 1 Then
            chkOthers.value = 0
            chkNotApproved.value = 0
        End If
    End Sub

    Private Sub chkNotApproved_Click()
        If chkNotApproved.value = 1 Then
            chkApproved.value = 0
        End If
    End Sub

    Private Sub chkOthers_Click()
        If chkOthers.value = 1 Then
            chkApproved.value = 0
        End If
    End Sub

    Private Sub cmdClear_Click()
        Call InitForm
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        Call SeatwiseContraFormListing
    End Sub

    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
        Call FillGrid
    End Sub

    Private Sub Form_Load()
        winXPC.InitIDESubClassing
        Call PopulateCombos
        Call InitForm
        Call FillGrid
        Call SeatGroupSettings
    End Sub
    
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, ".")
    End Sub

    Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, "#-")
    End Sub
    
    Private Sub InitForm()
        chkOthers.value = 1
        'chkNotApproved.value = 0
        
        dtpFromDate.value = gbStartingDate
        dtpToDate.value = gbTransactionDate
        
        cmbTransactionTypes.ListIndex = -1
        cmbInstrumentTypes.ListIndex = -1
        
        txtVoucherNo.Text = ""
        txtAmount.Text = ""
        txtDescription.Text = ""
    End Sub

    Private Sub PopulateCombos()
        Dim mSql As String
        
        mSql = "SELECT vchTransactionType,intTransactionTypeID FROM faTransactionType WHERE intGroupID = 30"
        PopulateList cmbTransactionTypes, mSql, , True, , True, enuSourceString.Saankhya
        
        mSql = "SELECT  vchInstrumentType,intInstrumentTypeID FROM faInstrumentTypes"
        PopulateList cmbInstrumentTypes, mSql, , True, , True, enuSourceString.Saankhya
    End Sub

    Private Sub FillGrid()      ' In Decending Order col 8 Type (>0) ----Voucher/Approved Type (0) ----Demand Not Approved
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        
        If chkApproved.value <> 1 And chkOthers.value <> 1 And chkNotApproved.value <> 1 Then
            MsgBox "Please Check to View", vbInformation
            chkNotApproved.SetFocus
            Exit Sub
        End If
        '''Connection'''
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        '''Query Generation
        mSql = ""
        If chkOthers.value = 1 Then
            mSql = "SELECT      Cast(faVouchers.intVoucherNo as varchar(20)) intVoucherNo,faVouchers.dtDate,faVouchers.vchInstrumentNo,vchInstrumentType,faVouchers.intTransactionTypeID,vchTransactionType,faVouchers.dtInstrumentDate,faVouchers.fltAmount,faVouchers.vchDescription,99 tnyStatus,intKeyID2" & vbNewLine
            mSql = mSql + "From faVouchers" & vbNewLine
            mSql = mSql + "LEFT JOIN   faTransactionType ON faTransactionType.intTransactionTypeID = faVouchers.intTransactionTypeID" & vbNewLine
            mSql = mSql + "LEFT JOIN   faInstrumentTypes ON faInstrumentTypes.intInstrumentTypeID = faVouchers.intInstrumentTypeID" & vbNewLine
            mSql = mSql + "Where faVouchers.tnyVoucherTypeID = 30" & Crieterias(1) & vbNewLine
            If chkNotApproved.value = 1 Then
                mSql = mSql + "Union All" & vbNewLine
            End If
        End If
        If chkNotApproved.value = 1 Then
            mSql = mSql + "Select * From(" & vbNewLine
            mSql = mSql + "SELECT      isNull(Cast(faVouchers.intVoucherNo as varchar(20)),faIDemandTBL.vchDemandNo) intVoucherNo,isNull(faVouchers.dtDate,faIDemandTBL.dtDemandDate) dtDate,faIDemandTBL.vchInstrumentNo,vchInstrumentType,faIDemandTBL.intTransactionTypeID,vchTransactionType,faIDemandTBL.dtInstrumentDate,faVouchers.fltAmount,faIDemandTBL.vchRemarks vchDescription,faIDemandTBL.tnyStatus,faVouchers.intKeyID2" & vbNewLine
            mSql = mSql + "From faIDemandTBL" & vbNewLine
            mSql = mSql + "Left Join faVouchers On faVouchers.intVoucherID = faIDemandTBL.intVoucherID" & vbNewLine
            mSql = mSql + "LEFT JOIN   faTransactionType ON faTransactionType.intTransactionTypeID = faIDemandTBL.intTransactionTypeID" & vbNewLine
            mSql = mSql + "LEFT JOIN   faInstrumentTypes ON faInstrumentTypes.intInstrumentTypeID = faIDemandTBL.intInstrumentTypeID" & vbNewLine
            mSql = mSql + "Where faIDemandTBL.tnyDemandType = 30 And faIDemandTBL.tnyExtModuleID in (25,50) And faIDemandTBL.tnyStatus = 0" & Crieterias(2) & ") A" & vbNewLine
        End If
        If chkApproved.value = 1 Then
            mSql = mSql + "Select * From(" & vbNewLine
            mSql = mSql + "SELECT      isNull(Cast(faVouchers.intVoucherNo as varchar(20)),faIDemandTBL.vchDemandNo) intVoucherNo,isNull(dtDate,faIDemandTBL.dtDemandDate) dtDate,faIDemandTBL.vchInstrumentNo,vchInstrumentType,faIDemandTBL.intTransactionTypeID,vchTransactionType,faIDemandTBL.dtInstrumentDate,faVouchers.fltAmount,faIDemandTBL.vchRemarks vchDescription,faIDemandTBL.tnyStatus,faVouchers.intKeyID2" & vbNewLine
            mSql = mSql + "From faIDemandTBL" & vbNewLine
            mSql = mSql + "Left Join faVouchers On faVouchers.intVoucherID = faIDemandTBL.intVoucherID" & vbNewLine
            mSql = mSql + "LEFT JOIN   faTransactionType ON faTransactionType.intTransactionTypeID = faIDemandTBL.intTransactionTypeID" & vbNewLine
            mSql = mSql + "LEFT JOIN   faInstrumentTypes ON faInstrumentTypes.intInstrumentTypeID = faIDemandTBL.intInstrumentTypeID" & vbNewLine
            mSql = mSql + "Where faIDemandTBL.tnyDemandType = 30 And faIDemandTBL.tnyExtModuleID in (25,50) And faIDemandTBL.tnyStatus > 0" & Crieterias(2) & ") B" & vbNewLine
        End If
        mSql = mSql + "Order By    intVoucherNo Desc,dtDate Desc"
        
        
        Rec.Open mSql, mCnn
        vsGrid.Rows = 1
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                With vsGrid
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    .TextMatrix(.Rows - 1, 4) = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    .TextMatrix(.Rows - 1, 5) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'                    If Rec!tnyStatus > 1 Then
'                        .Cell(flexcpBackColor, .Rows - 1, 7) = vbBlue
'                    Else
'                        .Cell(flexcpBackColor, .Rows - 1, 7) = vbWhite
'                    End If
                    .TextMatrix(.Rows - 1, 7) = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
                    .TextMatrix(.Rows - 1, 8) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                    .TextMatrix(.Rows - 1, 9) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)       '   not Approved/Approved Or Direct Contra
                    .TextMatrix(.Rows - 1, 10) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                End With
                Rec.MoveNext
            Wend
        End If
        '''Bolding the VoucherNo Column
        vsGrid.Cell(flexcpForeColor, 0, 0, vsGrid.Rows - 1) = vbBlue
        vsGrid.Cell(flexcpForeColor, 0, 7, vsGrid.Rows - 1) = vbBlue
        '''Fixing a Minimum Rows
        If vsGrid.Rows < 19 Then
            vsGrid.Rows = 19
        End If
    End Sub

    Private Function Crieterias(Table As Integer) As String
        Dim mSql As String
        If Table = 1 Then   '' Voucher Table
            mSql = " And dtDate Between '" & Format(dtpFromDate.value, "dd/MMM/yyyy") & "' And '" & Format(dtpToDate.value, "dd/MMM/yyyy") & "'"
            If Trim(txtVoucherNo.Text) <> "" Then
                mSql = mSql + " And faVouchers.intVoucherNo = '" & val(Trim(txtVoucherNo.Text)) & "'"
            End If
            If val(txtAmount.Text) > 0 Then
                mSql = mSql + " And faVouchers.fltAmount = '" & val(Trim(txtAmount.Text)) & "'"
            End If
            If Trim(txtDescription.Text) <> "" Then
                mSql = mSql + " And isNull(faVouchers.vchDescription,'') Like '%" & Trim(txtDescription.Text) & "%'"
            End If
            If cmbTransactionTypes.ListIndex > 0 Then
                mSql = mSql + " And faVouchers.intTransactionTypeID = '" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & "'"
            End If
            If cmbInstrumentTypes.ListIndex > 0 Then
                mSql = mSql + " And faVouchers.intInstrumentTypeID = '" & cmbInstrumentTypes.ItemData(cmbInstrumentTypes.ListIndex) & "'"
            End If
        Else                '' Demand Table
            mSql = " And dtDemandDate Between '" & Format(dtpFromDate.value, "dd/MMM/yyyy") & "' And '" & Format(dtpToDate.value, "dd/MMM/yyyy") & "'"
            If Trim(txtVoucherNo.Text) <> "" Then
                mSql = mSql + " And faIDemandTBL.vchDemandNo = '" & Trim(txtVoucherNo.Text) & "'"
            End If
            If Trim(txtDescription.Text) <> "" Then
                mSql = mSql + " And faIDemandTBL.vchRemarks = '%" & Trim(txtDescription.Text) & "%'"
            End If
            If cmbTransactionTypes.ListIndex > 0 Then
                mSql = mSql + " And faIDemandTBL.intTransactionTypeID = '" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & "'"
            End If
            If cmbInstrumentTypes.ListIndex > 0 Then
                mSql = mSql + " And faIDemandTBL.intInstrumentTypeID = '" & cmbInstrumentTypes.ItemData(cmbInstrumentTypes.ListIndex) & "'"
            End If
        End If
        Crieterias = mSql
    End Function

    Private Sub SeatGroupSettings()
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            cmdNew.Caption = "&New"
            cmdNew.Tag = 0          '   For Data Entry Users
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
            cmdNew.Caption = "&Verify"
            cmdNew.Tag = 1          '   For Approval Users
        Else
            cmdNew.Tag = -1
            cmdNew.Visible = False
        End If
    End Sub

    Private Sub SeatwiseContraFormListing()
        mGridRow = -1
        If cmdNew.Tag = 0 Then
            frmContraEntry.PreviousYearMode = 0
            frmContraEntry.Visible = True
            frmContraEntry.ZOrder (0)
            frmContraEntry.cmdNew.Enabled = True
            frmContraEntry.cmdSave.Caption = "&Save"
            frmContraEntry.cmdSave.Tag = 0
        ElseIf cmdNew.Tag = 1 Then
            If vsGrid.Row > 0 And Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) <> "" Then
                frmContraEntry.cmdNew.Enabled = False
                If vsGrid.TextMatrix(vsGrid.Row, 9) > 0 Then
                    frmContraEntry.cmdSave.Enabled = False
                    MsgBox "The Contra Entry already Approved", vbInformation
                    Exit Sub
                End If
                mGridRow = vsGrid.Row           ' To Update the row From Contra Entry Screen
                frmContraEntry.PreviousYearMode = 0
                frmContraEntry.Visible = True
                frmContraEntry.ZOrder (0)
                frmContraEntry.cmdSave.Caption = "&Approve"
                frmContraEntry.cmdSave.Tag = 1
'                frmContraEntry.cmdReject.Visible = True
                '   Filling the Contra Voucher in the Contra Voucher Load
                Call frmContraEntry.ListContraDemandOrVoucher(vsGrid.TextMatrix(vsGrid.Row, 0))
            Else
                MsgBox "Please Select a voucher number to Make Approval", vbInformation
            End If
        Else
            MsgBox "Invalid User", vbCritical
        End If
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 And vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
            If vsGrid.Col = 7 And val(vsGrid.TextMatrix(vsGrid.Row, 7)) > 0 Then
                Call frmJournalEntry.DisplayReceiptDetails(val(vsGrid.TextMatrix(vsGrid.Row, 7)))
                frmJournalEntry.cmdNew.Enabled = False
                frmJournalEntry.cmdSave.Enabled = False
            Else
                Call SeatwiseContraFormListing
                mGridRow = vsGrid.Row
                If cmdNew.Tag = 0 Then
                    Call frmContraEntry.ListContraDemandOrVoucher(vsGrid.TextMatrix(vsGrid.Row, 0))     ''' Filling Contra Voucher
                    frmContraEntry.copiedAmount = ""
                    If val(vsGrid.TextMatrix(vsGrid.Row, 9)) > 0 Then
                        If (val(vsGrid.TextMatrix(vsGrid.Row, 10)) = gbTransactionTypeContraRegularPension Or val(vsGrid.TextMatrix(vsGrid.Row, 10)) = gbTransactionTypeContraContingentPension) Then
                            frmContraEntry.cmdSave.Enabled = False
                        End If
                    Else
                        If val(vsGrid.TextMatrix(vsGrid.Row, 10)) = gbTransactiontypeDailyCollection Then
                            frmContraEntry.copiedAmount = 0   ' If Not Approved and User Type Clerk And JSK Daily Collection
                        End If
                    End If
                End If
            End If
        End If
    End Sub
'    Private Sub vsGrid_Click()
'        If vsGrid.Row > 0 Then
'            If (vsGrid.Col = 0 Or vsGrid.Col = 9) Then
'
'            End If
'        End If
'    End Sub
