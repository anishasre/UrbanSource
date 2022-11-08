VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmBank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Bank  Account"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDeactivate 
      Caption         =   "DeActivate"
      Height          =   375
      Left            =   6945
      TabIndex        =   43
      Top             =   7755
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   3270
      TabIndex        =   24
      Top             =   7755
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   5730
      TabIndex        =   25
      Top             =   7755
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4500
      TabIndex        =   23
      Top             =   7755
      Width           =   1215
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
      ScaleWidth      =   11220
      TabIndex        =   38
      Top             =   0
      Width           =   11220
   End
   Begin VB.Frame Frame1 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   11220
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   6735
         Left            =   11115
         TabIndex        =   46
         Top             =   135
         Visible         =   0   'False
         Width           =   1245
         _cx             =   2196
         _cy             =   11880
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBank.frx":0000
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
      Begin VB.ListBox lstBanks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5430
         ItemData        =   "frmBank.frx":003E
         Left            =   4815
         List            =   "frmBank.frx":0040
         TabIndex        =   27
         Top             =   1350
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.ComboBox cmbPensionType 
         Height          =   315
         Left            =   3960
         TabIndex        =   44
         Top             =   1620
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.ListBox lstMasters 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         ItemData        =   "frmBank.frx":0042
         Left            =   4065
         List            =   "frmBank.frx":0049
         TabIndex        =   37
         Top             =   570
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4035
         MaxLength       =   50
         TabIndex        =   2
         Top             =   210
         Width           =   2985
      End
      Begin VB.TextBox txtFundCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   1
         Top             =   210
         Width           =   1260
      End
      Begin VB.CommandButton cmdSearchFund 
         Caption         =   "..."
         Height          =   285
         Left            =   7065
         TabIndex        =   3
         Top             =   225
         Width           =   345
      End
      Begin VB.ComboBox cmbNatureofFunds 
         Height          =   315
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   3030
      End
      Begin VB.TextBox txtBankCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2640
         Width           =   2985
      End
      Begin VB.TextBox txtBranchCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2955
         Width           =   2985
      End
      Begin VB.OptionButton optCredit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6180
         TabIndex        =   22
         Top             =   5460
         Width           =   765
      End
      Begin VB.OptionButton optDebit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6180
         TabIndex        =   21
         Top             =   5220
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   7020
         TabIndex        =   10
         Top             =   2040
         Width           =   345
      End
      Begin VB.TextBox txtOpeningBalance 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   15
         TabIndex        =   20
         Top             =   5250
         Width           =   1995
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   19
         Top             =   4845
         Width           =   2985
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   15
         TabIndex        =   18
         Top             =   4530
         Width           =   2025
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   15
         TabIndex        =   17
         Top             =   4215
         Width           =   2025
      End
      Begin VB.TextBox txtAddress2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3900
         Width           =   2985
      End
      Begin VB.TextBox txtAddress1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3585
         Width           =   2985
      End
      Begin VB.TextBox txtAccountNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   16
         TabIndex        =   14
         Top             =   3270
         Width           =   2985
      End
      Begin VB.TextBox txtBranch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2325
         Width           =   2985
      End
      Begin VB.TextBox txtNameOfBank 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2010
         Width           =   2985
      End
      Begin VB.TextBox txtAccountHeadCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2715
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1290
         Width           =   1260
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3990
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1290
         Width           =   2985
      End
      Begin VB.ComboBox cmbBanks 
         Height          =   315
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   900
         Width           =   3030
      End
      Begin VB.CommandButton cmdSearchPrimaryAcHead 
         Caption         =   "..."
         Height          =   285
         Left            =   7020
         TabIndex        =   8
         Top             =   1290
         Width           =   345
      End
      Begin VB.Label Label15 
         Caption         =   "Pension Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2730
         TabIndex        =   45
         Top             =   1650
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Account Head"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1530
         TabIndex        =   42
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2295
         TabIndex        =   41
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3465
         TabIndex        =   40
         Top             =   945
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Fund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2700
         TabIndex        =   39
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bank Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3000
         TabIndex        =   29
         Top             =   2670
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2835
         TabIndex        =   30
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Passbook Opening Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1575
         TabIndex        =   36
         Top             =   5280
         Width           =   2310
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3420
         TabIndex        =   35
         Top             =   4890
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3615
         TabIndex        =   34
         Top             =   4560
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3360
         TabIndex        =   33
         Top             =   4245
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3210
         TabIndex        =   32
         Top             =   3615
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Account No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2985
         TabIndex        =   31
         Top             =   3330
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3330
         TabIndex        =   28
         Top             =   2355
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name of Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2745
         TabIndex        =   26
         Top             =   2040
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mEditFlag As Boolean
    Dim mAccountHeadID      As Variant
    Dim mTransactionID      As Variant
    Dim mRegularPensionID   As Variant
    Dim mContigentPensionID As Variant
    Dim mBankType           As Variant
    Dim mCount              As Integer
    Private Sub DisplayBank(mID As Double)
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim objAcc As New clsAccounts
        Dim objBk As New clsBank
        Dim objFund As New clsFund
        
        objBk.SetBankInfo mID
        If objBk.BankID > -1 Then
            mEditFlag = True
            objAcc.SetAccountID objBk.BankAccountHeadID
            objFund.SetFund objBk.FundID
            txtNameOfBank.Text = objBk.BankName
            txtNameOfBank.Tag = objBk.BankID
            txtBranch.Text = objBk.Branch
            txtBankCode.Text = objBk.BankCode
            txtBranchCode.Text = objBk.BranchCode
            txtAccountNo.Text = objBk.AccountNumber
            txtAccountHead.Text = objAcc.AccountHead
            txtAccountHead.Tag = objAcc.AccountHeadID
            txtAccountHeadCode.Text = objAcc.AccountCode
            txtFundCode.Text = objFund.FundCode
            txtFundCode.Tag = objFund.FundID
            txtFund.Text = objFund.FundName
            txtAddress1.Text = objBk.Address1
            txtAddress2.Text = objBk.Address2
            txtPhone.Text = objBk.Phone
            txtEmail.Text = objBk.Email
            txtFax.Text = objBk.Fax
            txtOpeningBalance.Text = Format(Abs(objBk.Opening), "0.00")
            If objBk.CrDrFlag = True Then 'If objBk.Opening < 0 Then
                optCredit.value = True
            Else
                optDebit.value = True
            End If
        End If
        Set objBk = Nothing
    End Sub
    
    Private Sub FormInitialize()
    
        mEditFlag = False
        txtNameOfBank.Text = ""
        txtNameOfBank.Tag = ""
        txtBranch.Text = ""
        txtBankCode.Text = ""
        txtBranchCode.Text = ""
        txtAccountNo.Text = ""
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
        txtAddress1.Text = ""
        txtAddress2.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        txtOpeningBalance.Text = ""
        txtAccountHeadCode.Text = ""
        txtFund.Text = ""
        txtFundCode.Text = ""
        txtFundCode.Tag = ""
        optCredit.value = False
        optDebit.value = True
        cmbBanks.ListIndex = -1
        cmbNatureofFunds.ListIndex = -1
        cmbPensionType.ListIndex = -1
    End Sub
    Private Sub cmbBanks_Click()
        If cmbBanks.ListIndex < 1 Then
            txtAccountHead.Text = ""
            txtAccountHeadCode.Text = ""
            txtAccountHead.Tag = ""
            Exit Sub
        End If
        If cmbBanks.ItemData(cmbBanks.ListIndex) <> mID(Trim(txtAccountHeadCode), 5, 2) Then
            txtAccountHead.Text = ""
            txtAccountHeadCode.Text = ""
            txtAccountHead.Tag = ""
        End If
        If cmbNatureofFunds.ListIndex > 0 Then
            If gbLBType = 3 Or gbLBType = 4 Then
                If cmbNatureofFunds.ItemData(cmbNatureofFunds.ListIndex) = gbSpecialFund And cmbBanks.ItemData(cmbBanks.ListIndex) = 50 Then
                    Label15.Visible = True
                    cmbPensionType.Visible = True
                Else
                    Label15.Visible = False
                    cmbPensionType.Visible = False
                End If
            End If
        End If
    End Sub

    Private Sub cmbNatureofFunds_Click()
        If cmbNatureofFunds.ListIndex < 1 Then
            txtAccountHead.Text = ""
            txtAccountHeadCode.Text = ""
            txtAccountHead.Tag = ""
            Exit Sub
        End If
        If cmbNatureofFunds.ItemData(cmbNatureofFunds.ListIndex) <> Left(Trim(txtAccountHeadCode), 4) Then
            txtAccountHead.Text = ""
            txtAccountHeadCode.Text = ""
            txtAccountHead.Tag = ""
        End If
    End Sub
    Private Sub cmbPensionType_Click()
        If cmbPensionType.ListIndex > 0 Then
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objdb As New clsDB
            Dim mSql As String
            objdb.SetConnection mCnn
            Rec.Open "SELECT intRegularTreasuryPensionAccountHeadID, intContingentTreasuryPensionAccountHeadID FROM  faConfig", mCnn
'            mRegularPensionID = IIf(IsNull(Rec!intRegularTreasuryPensionAccountHeadID), 0, Rec!intRegularTreasuryPensionAccountHeadID)
'            mContigentPensionID = IIf(IsNull(Rec!intContingentTreasuryPensionAccountHeadID), 0, Rec!intContingentTreasuryPensionAccountHeadID)
            'Note:- Check whether the any treasury account is defined for the selected Pension Type
            If (cmbPensionType.ItemData(cmbPensionType.ListIndex)) <> val(txtAccountHead.Tag) And _
                cmbPensionType.ItemData(cmbPensionType.ListIndex) <> 0 Then
                MsgBox "  There is already a Treasury AccountHead" & vbCrLf + "  defined for  " & cmbPensionType.Text & " " & vbCrLf + "   If you want to continue set the Previous Account Head to Null "
                
                cmbPensionType.Text = ""
            End If
                
            Rec.Close
        End If
    End Sub

    Private Sub cmdCancel_Click()
        Call FormInitialize
        Unload Me
    End Sub

    Private Sub cmdDeactivate_Click()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mCount As Variant
        'Dim objAcc As New clsAccounts
        'Dim mTransactionID As Variant
        'Dim mAccountHeadID As Variant
        
        
       objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
       
       If Trim(txtNameOfBank.Text) = "" Then
            MsgBox "Enter name of Bank", vbInformation
            txtNameOfBank.SetFocus
            Exit Sub
        End If
       
        'objAcc.SetAccountID (val(txtAccountHead.Tag))
            mAccountHeadID = val(txtAccountHead.Tag) 'objAcc.AccountHeadID
        If (mAccountHeadID <> Null) Then
            mSql = "Select Count(*) mCount From faTransactionChild   Inner Join"
            mSql = mSql + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID"
            mSql = mSql + " Where (tnyStatus Is Null Or tnyStatus <> 4) And intAccountHeadID = " & mAccountHeadID & " "
        Else
            mSql = " Select Count(*) mCount From faTransactionChild   Inner Join"
            mSql = mSql + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID"
            mSql = mSql + " Where (tnyStatus Is Null Or tnyStatus <> 4) And intByAccountHeadID = " & mAccountHeadID & " "
        End If
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            mCount = Rec!mCount
        End If
        Rec.Close
        If mCount = 0 Then
            mCnn.Execute "DELETE FROM faBanks WHERE intBankID = " & val(Trim(txtNameOfBank.Tag))
        Else
            mCnn.Execute "Update faAccountHeads Set vchAccountHead = vchAlias,tinHiddenFlag = 1 From faAccountHeads Inner Join faBanks On faAccountHeads.intAccountHeadID = faBanks.intAccountHeadID WHERE intBankID = " & Trim(txtNameOfBank.Tag) & ";"
            mCnn.Execute " Update faBanks Set vchBankName = vchAlias From faAccountHeads Inner Join faBanks On faAccountHeads.intAccountHeadID = faBanks.intAccountHeadID WHERE intBankID =" & Trim(txtNameOfBank.Tag) & " "
            
        End If
        MsgBox "Updated Succesfully"
        Call FormInitialize
    End Sub
    
'    Private Sub TransactionID()
'        Dim objDb As New clsDb
'        Dim mCnn As New ADODB.Connection
'        Dim Rec As New ADODB.Recordset
'        Dim mSql As String
'        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        mSql = " SELECT intTransactionID  FROM faTransactionTypeChild WHERE intAccountHeadID=" & mAccID
'        If Not (Rec.EOF And Rec.BOF) Then
'            mTransactionID = Rec!intTransactionID
'        End If
'        Rec.Close
'
'    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        txtNameOfBank.SetFocus
    End Sub

    Private Sub cmdSave_Click()
        Dim objdb               As New clsDB
        Dim objAcc              As New clsAccounts
        Dim objBank             As New clsBank
        Dim mCnn                As New ADODB.Connection
        Dim mPryAccountHeadID   As Long
        Dim mSecAccountHeadID   As Long
        Dim mAmt                As Double
        Dim arrInput            As Variant
        Dim intBankID           As Integer
        Dim intFundID           As Variant
        Dim objFund             As New clsFund
        Dim Rec                 As New ADODB.Recordset
        Dim mAccID              As Variant
        Dim mSql                As Variant
        '---------------------------------------------------'
        '  Validations                                      '
        '---------------------------------------------------'
        If Trim(txtNameOfBank.Text) = "" Then
            MsgBox "Enter name of Bank", vbInformation
            txtNameOfBank.SetFocus
            Exit Sub
        End If
        
        If Trim(txtAccountNo.Text) = "" Then
            MsgBox "Enter Account Number!", vbInformation
            txtAccountNo.SetFocus
            Exit Sub
        End If
        If Trim(txtAccountHead.Text) = "" Or Trim(txtAccountHeadCode) = "" Then
            MsgBox "Please specify the Account Head", vbInformation
            Exit Sub
        End If
        objAcc.SetAccountID (val(txtAccountHead.Tag))
            mAccID = objAcc.AccountHeadID
        If mAccID = -1 Then
            MsgBox "Select a Primary account head", vbInformation
            txtAccountHead.SetFocus
            Exit Sub
        Else
            mPryAccountHeadID = objAcc.AccountHeadID
        End If
        If Not (val(mID(Trim(txtAccountHeadCode.Text), 1, 4)) > 4501 And val(mID(Trim(txtAccountHeadCode.Text), 1, 3))) = 450 Then
            MsgBox "Please select a bank head", vbInformation
            cmdSearchPrimaryAcHead.SetFocus
            Exit Sub
        End If
        objFund.SetFund (val(txtFundCode.Tag))
        
        If optCredit.value Then
            mAmt = Abs(val(txtOpeningBalance)) * -1
        Else
            mAmt = Abs(val(txtOpeningBalance))
        End If
        
        objBank.SetBankInfoByAccID (mAccID)
        If objBank.BankID > 0 Then
            If (mEditFlag And objBank.BankID <> val(txtNameOfBank.Tag)) Or mEditFlag = False Then
                MsgBox "This Account Head already exists!!!please select another one"
                Exit Sub
            End If
        End If
        objAcc.FindAccountByHead (Trim(txtAccountHead.Text))
        If objAcc.AccountHeadID > 0 Then
            If Trim(txtAccountHeadCode.Text) <> objAcc.AccountCode Then
                MsgBox "Account head already Exists"
                Exit Sub
            End If
        End If
'        objDB.SetConnection mCnn
'        Rec.Open "Select * From faBanks Where vchBankName = '" & Trim(txtNameOfBank.Text) & "'", mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            If Rec!intAccountHeadID <> txtAccountHead.Tag Then
'                MsgBox "Please select another name for the Bank, as it is already assigned", vbInformation
'                txtNameOfBank.SetFocus
'                Exit Sub
'            End If
'        End If
'        Rec.Close
        '---------------------------------------------------'
        '  Updating                                         '
        '---------------------------------------------------'
        
        arrInput = Array(IIf(mEditFlag, val(txtNameOfBank.Tag), Null), _
                    mAccID, _
                    objFund.FundID, _
                    Trim(txtNameOfBank.Text), _
                    Trim(txtBranch), _
                    Trim(txtBankCode), _
                    Trim(txtBranchCode), _
                    Trim(txtAccountNo), _
                    Trim(txtAddress1), _
                    Trim(txtAddress2), _
                    Trim(txtPhone), _
                    Trim(txtFax), _
                    Trim(txtEmail), _
                    Format(mAmt, "0.00"), _
                    IIf(optDebit, 1, 0) _
                    )
        
        'mCnn.BeginTrans
        'On Error GoTo ErrRollBack:
        objdb.ExecuteSP "spSaveBank", arrInput, , , mCnn
        mCnn.Execute "Update faAccountHeads Set tinHiddenFlag = 0, vchAccountHead = '" & txtAccountHead.Text & "' Where vchAccountHeadCode = '" & Trim(txtAccountHeadCode.Text) & "'"


        '--------------------------------------------------------------------------------'
        'FOR SPECIAL FUND TREASURY ACCOUNTS LINKING TO faConfig Table
        '--------------------------------------------------------------------------------'
        'Note: Regulare Pension type is trying to add No Reg.Pension is defined in Config
        If cmbPensionType.ListIndex = 1 And mRegularPensionID = 0 Then
            If val(txtAccountHead.Tag) <> mContigentPensionID Then
               mSql = "update faConfig set intRegularTreasuryPensionAccountHeadID= " & txtAccountHead.Tag & " " 'Can update Regular Pension Field
               mCnn.Execute mSql
            End If
        End If
        
        'Note: Contingent Pension type is trying to add No Con.Pension is defined in Config
        If cmbPensionType.ListIndex = 2 And mContigentPensionID = 0 Then
            If val(txtAccountHead.Tag) <> mRegularPensionID Then
                 mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "    'Can update Contingent Pension Field
                mCnn.Execute mSql
            End If
        End If
        
        If val(txtAccountHead.Tag) = mRegularPensionID Then
            If cmbPensionType.ListIndex = 0 Then
               mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  0" 'Set NULL in Regular Pension Fund Field
                mCnn.Execute mSql
            End If
            
            If cmbPensionType.ListIndex = 2 Then
                If mContigentPensionID = 0 Then
                    mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  0;" 'Set NULL in Regular Pension Fund Field
                    mSql = mSql + "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " , intRegularTreasuryPensionAccountHeadID= 0 " 'Update Contingent Field
                    mCnn.Execute mSql
                End If
            End If
        End If
        
        If val(txtAccountHead.Tag) = mContigentPensionID Then
            If cmbPensionType.ListIndex = 0 Then
                 mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  0" 'SET Null in Contingent Pension Field
                 mCnn.Execute mSql
            End If
            If cmbPensionType.ListIndex = 1 Then
                If mRegularPensionID = 0 Then
                    mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  0;" 'SET Null in Contingent Pension Field
                    mSql = mSql + "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " ,intContingentTreasuryPensionAccountHeadID = 0 " 'Update Regular Pension Fund field
                    mCnn.Execute mSql
                End If
            End If
        End If
        
        '--------------------------------------------------------------------------------'
        
        
        
        
        
'If cmbNatureofFunds.ItemData(cmbNatureofFunds.ListIndex) = gbSpecialFund And cmbBanks.ItemData(cmbBanks.ListIndex) = 50 Then
''-------------------------------------------------------------------------------------------------------------------------'
''                                   Inserting for the first time                                                          '
''-------------------------------------------------------------------------------------------------------------------------'
'
'            If val(txtAccountHead.Tag) <> mRegularPensionID And val(txtAccountHead.Tag) <> mContigentPensionID Then
'                If cmbPensionType.ListIndex > 0 Then
'                    If (cmbPensionType.ListIndex) = 1 Then
'                        mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "
'                    ElseIf (cmbPensionType.ListIndex) = 2 Then
'                        mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "
'                    End If
'                End If
'         '-------------------------------------------------------------------------------------------------------------------------'
'        '                                      Editng                                                          '
'        '-------------------------------------------------------------------------------------------------------------------------'
'            ElseIf val(txtAccountHead.Tag) <> mRegularPensionID Then
'                    If cmbPensionType.ListIndex > 0 Then
'                        If cmbPensionType.ItemData(cmbPensionType.ListIndex) = mRegularPensionID Then
'                            mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "
'                        ElseIf cmbPensionType.ItemData(cmbPensionType.ListIndex) = mContigentPensionID Then
'                            mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "
'                        End If
'                    End If
'            ElseIf val(txtAccountHead.Tag) <> mContigentPensionID Then
'                    If cmbPensionType.ListIndex > 0 Then
'                        If cmbPensionType.ListIndex = 1 Then
'                            mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "
'                        ElseIf cmbPensionType.ListIndex = 2 Then
'                            mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " "
'
'                        End If
'                    End If
'            End If
'            If val(txtAccountHead.Tag) = mRegularPensionID Then
'                If cmbPensionType.ListIndex = 0 Then
'                    mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  0"
'                ElseIf cmbPensionType.ListIndex > 0 Then
'                    mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " , intRegularTreasuryPensionAccountHeadID= 0 "
'                End If
'            ElseIf val(txtAccountHead.Tag) = mContigentPensionID Then
'                If cmbPensionType.ListIndex = 0 Then
'                    mSql = "Update faConfig Set intContingentTreasuryPensionAccountHeadID =  0"
'                ElseIf cmbPensionType.ListIndex > 0 Then
'                    mSql = "Update faConfig Set intRegularTreasuryPensionAccountHeadID =  " & val(txtAccountHead.Tag) & " ,intContingentTreasuryPensionAccountHeadID = 0 "
'                End If
'            End If
'    'End If
'End If
        
        'mCnn.CommitTrans
        Call FormInitialize
        'Exit Sub

ErrRollBack:
        'mCnn.RollbackTrans
        Set mCnn = Nothing
    End Sub
    Private Sub cmdSearch_Click()
        Call PopulateList(lstBanks, "Select vchBankName, intBankID from faBanks Order By vchBankName", , , , True)
        lstBanks.Visible = True
        lstBanks.SetFocus
    End Sub
    Private Sub cmdSearchFund_Click()
        Dim mSql As String
        mSql = "Select vchFund, intFundID From faFunds Where tnyActiveFlag  = 1 Order By vchFund"
        Call PopulateList(lstMasters, mSql, , , , True)
        
        lstMasters.Left = 5000
        lstMasters.Width = 2500
        lstMasters.Height = 3000
        lstMasters.Visible = True
        lstMasters.SetFocus
    End Sub
    Private Sub cmdSearchPrimaryAcHead_Click()
        If cmbNatureofFunds.ListIndex < 1 Or cmbBanks.ListIndex < 1 Then
            MsgBox "Please Select Nature of Fund Or Bank", vbInformation
            Exit Sub
        End If
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads LEFT JOIN faMajorAccountHeads ON faAccountHeads.intMajorAccountHeadID =faMajorAccountHeads.intMajorAccountHeadID Where faAccountHeads.intGroupID =2 And Left(vchAccountHeadCode,6) = '" & CStr(cmbNatureofFunds.ItemData(cmbNatureofFunds.ListIndex)) + CStr(cmbBanks.ItemData(cmbBanks.ListIndex)) & "'"
        frmSearchAccountHeads.Show vbModal
        txtAccountHeadCode.SetFocus
    End Sub
        
    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = (frmMenu.Width - Me.Width) / 2
    End Sub

    Private Sub Form_Load()
    
    Dim mCnn As New ADODB.Connection
        Call FormInitialize
        cmbNatureofFunds.AddItem ("")
        cmbNatureofFunds.AddItem ("Own Fund")
        cmbNatureofFunds.ItemData(cmbNatureofFunds.NewIndex) = 4502
        cmbNatureofFunds.AddItem ("Special Fund")
        cmbNatureofFunds.ItemData(cmbNatureofFunds.NewIndex) = 4504
        cmbNatureofFunds.AddItem ("Grant Fund")
        cmbNatureofFunds.ItemData(cmbNatureofFunds.NewIndex) = 4506
        
        cmbBanks.AddItem ("")
        cmbBanks.AddItem ("Nationalised Banks")
        cmbBanks.ItemData(cmbBanks.NewIndex) = 10
        cmbBanks.AddItem ("Other Scheduled Banks")
        cmbBanks.ItemData(cmbBanks.NewIndex) = 20
        cmbBanks.AddItem ("Co-operative Banks")
        cmbBanks.ItemData(cmbBanks.NewIndex) = 30
        cmbBanks.AddItem ("Treasury")
        cmbBanks.ItemData(cmbBanks.NewIndex) = 50
        Call FillPensionTypeCombo
     
    End Sub
    Private Sub FillPensionTypeCombo()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        
        objdb.SetConnection mCnn
        Rec.Open "SELECT intRegularTreasuryPensionAccountHeadID, intContingentTreasuryPensionAccountHeadID FROM  faConfig", mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            'Note: Adding Blank Row
            cmbPensionType.AddItem ""
            cmbPensionType.ItemData(cmbPensionType.NewIndex) = 0
            'Note: Adding Regular Pesion
            cmbPensionType.AddItem ("Pension Fund for Regular Employees")
            cmbPensionType.ItemData(cmbPensionType.NewIndex) = IIf(IsNull(Rec!intRegularTreasuryPensionAccountHeadID), 0, Rec!intRegularTreasuryPensionAccountHeadID)
            'Note: Adding Contingent Pesion
            cmbPensionType.AddItem ("Pension Fund for Contigent Employees")
            cmbPensionType.ItemData(cmbPensionType.NewIndex) = IIf(IsNull(Rec!intContingentTreasuryPensionAccountHeadID), 0, Rec!intContingentTreasuryPensionAccountHeadID)
            
            mRegularPensionID = IIf(IsNull(Rec!intRegularTreasuryPensionAccountHeadID), 0, Rec!intRegularTreasuryPensionAccountHeadID)
            mContigentPensionID = IIf(IsNull(Rec!intContingentTreasuryPensionAccountHeadID), 0, Rec!intContingentTreasuryPensionAccountHeadID)
        End If
        Rec.Close
    End Sub
    
    Private Sub lstBanks_DblClick()
        Call lstBanks_KeyDown(13, 0)
    End Sub

    Private Sub lstBanks_KeyDown(KeyCode As Integer, Shift As Integer)
        mEditFlag = False
        If KeyCode = 13 Then
            Dim objBank As New clsBank
            Dim objAcc As New clsAccounts
            Dim objFund As New clsFund
            
            objBank.SetBankInfo (lstBanks.ItemData(lstBanks.ListIndex))
            If objBank.BankID > 0 Then
                mEditFlag = True
                txtNameOfBank.Text = objBank.BankName
                txtNameOfBank.Tag = objBank.BankID
                txtBranch.Text = objBank.Branch
                txtBankCode.Text = objBank.BankCode
                txtBranchCode.Text = objBank.BranchCode
                txtAccountNo.Text = objBank.AccountNumber
                objAcc.SetAccountID objBank.BankAccountHeadID
                If objAcc.AccountHeadID > 0 Then
                    txtAccountHead.Text = objAcc.AccountHead
                    txtAccountHead.Tag = objBank.BankAccountHeadID
                    txtAccountHeadCode.Text = objAcc.AccountCode
                Else
                    txtAccountHead.Text = ""
                    txtAccountHead.Tag = ""
                    txtAccountHeadCode.Text = ""
                End If
                objFund.SetFund objBank.FundID
                If objFund.FundID > 0 Then
                    txtFundCode.Text = objFund.FundCode
                    txtFundCode.Tag = objFund.FundID
                    txtFund.Text = objFund.FundName
                Else
                    txtFundCode.Text = ""
                    txtFundCode.Tag = ""
                    txtFund.Text = ""
                End If
                txtAddress1.Text = objBank.Address1
                txtAddress2.Text = objBank.Address2
                txtPhone.Text = objBank.Phone
                txtFax.Text = objBank.Fax
                txtEmail.Text = objBank.Email
'                txtOpeningBalance.Text = objBank.Opening
                txtOpeningBalance.Text = Format(Abs(objBank.Opening), "0.00")  'Added By Poornima on 12/10/2011
                If objBank.CrDrFlag = True Then
                    optDebit.value = True
                Else
                    optCredit.value = True
                End If
                '---------------------------------------------------------------------------------
                '               To fill Data on the Combo Boxes
                '---------------------------------------------------------------------------------
                Call gSubSetComboItem2(cmbNatureofFunds, mID(txtAccountHeadCode.Text, 1, 4))
                Call gSubSetComboItem2(cmbBanks, mID(txtAccountHeadCode, 5, 2))
                
               
                If cmbNatureofFunds.ItemData(cmbNatureofFunds.ListIndex) = gbSpecialFund And cmbBanks.ItemData(cmbBanks.ListIndex) = 50 Then
                    Label15.Visible = True
                    cmbPensionType.Visible = True
                Else
                    Label15.Visible = False
                    cmbPensionType.Visible = False
                End If
                If val(txtAccountHead.Tag) = mRegularPensionID Or mContigentPensionID Then
                    Call gSubSetComboItem2(cmbPensionType, txtAccountHead.Tag)
                End If
'                If mID(txtAccountHeadCode, 5, 2) <> "" Then
'                      For mCount = 0 To cmbPensionType.ListCount - 1
'                          If cmbPensionType.ItemData(mCount) = mID(txtAccountHeadCode, 5, 2) Then
'                              cmbPensionType.ListIndex = mCount
'                          End If
'                      Next mCount
'                End If
            End If
            lstBanks.Visible = False
        End If
    End Sub
            
    Private Sub lstBanks_LostFocus()
        lstBanks.Visible = False
    End Sub
    
    Private Sub optCredit_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub optDebit_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
          
    Private Sub txtAccountHead_GotFocus()
        
        txtAccountHead.SelStart = 0
        txtAccountHead.SelLength = Len(txtAccountHead)
        'Call DisplayBanksByAccount
    End Sub
    
    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
        
    Private Sub txtAccountHead_LostFocus()
        If val(txtAccountHead.Tag) < 1 Then
            txtAccountHead.Text = ""
        End If
    End Sub
    
Private Sub txtAccountHeadCode_GotFocus()
    If gbSearchStr <> "" Then
        Dim mStr As String
        txtAccountHeadCode.Text = Trim(Token(gbSearchStr, " "))
        txtAccountHead.Text = Trim(gbSearchStr)
        txtAccountHead.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        Call DisplayBanksByAccount
    End If
End Sub

    Private Sub txtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
         If Shift = vbCtrlMask And (Chr(KeyCode) = "v" Or Chr(KeyCode) = "V") Then
        txtAccountNo.Locked = True
    Else
        txtAccountNo.Locked = False
    End If
    End Sub

    Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtAccountNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
         If Button = vbRightButton Then
        txtAccountNo.Locked = True
    Else
        txtAccountNo.Locked = False
    End If
    End Sub

    Private Sub txtBranch_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtDetailedCode_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtEmail_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtFax_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtMajorHeadCode_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub

    Private Sub txtNameOfBank_GotFocus()
        If gbSearchStr <> "" Then
            txtNameOfBank.Text = gbSearchStr
            txtNameOfBank.Tag = gbSearchID
            Call DisplayBank(gbSearchID)
            gbSearchStr = ""
            gbSearchID = -1
        End If
        txtNameOfBank.SelStart = 0
        txtNameOfBank.SelLength = Len(txtNameOfBank.Text)
    End Sub

    Private Sub txtNameOfBank_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
        
    Private Sub txtNameOfBank_LostFocus()
'        If Val(txtNameOfBank.Tag) > 0 Then
'            Call DisplayBank(Val(txtNameOfBank.Tag))
'        End If
    End Sub
    
    Private Sub txtOpeningBalance_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtOpeningBalance_LostFocus()
        txtOpeningBalance.Text = Format(val(txtOpeningBalance.Text), "0.00")
        If val(txtOpeningBalance.Text) < 0 Then
            txtOpeningBalance.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub txtPhone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub lstMasters_DblClick()
        Dim objFund As New clsFund
        objFund.SetFund lstMasters.ItemData(lstMasters.ListIndex)
        If Not IsNull(objFund.FundID) Then
            txtFundCode.Text = objFund.FundCode
            txtFundCode.Tag = objFund.FundID
            txtFund.Text = objFund.FundName
            Set objFund = Nothing
        End If
        lstMasters.Visible = False
        txtFund.SetFocus
    End Sub

    Private Sub lstMasters_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call lstMasters_DblClick
        End If
    End Sub

    Private Sub lstMasters_LostFocus()
        lstMasters.Visible = False
    End Sub
    Private Sub InsertNewBanks()
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim objAcc As New clsAccounts
        Dim mNFund As Integer
        Dim mBank As Integer
        Dim mCount As Integer
        Dim mArIn As Variant
        Dim mHead As String
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connection to Saankhya Not Present", vbCritical
            Exit Sub
        End If
        For mNFund = 1 To cmbNatureofFunds.ListCount - 1
            For mBank = 1 To cmbBanks.ListCount - 1
                For mCount = 100 To 900 Step 100
                    mHead = CStr(cmbNatureofFunds.ItemData(mNFund)) + CStr(cmbBanks.ItemData(mBank)) + CStr(mCount)
                    mSql = "if not Exists(Select * From faAccountHeads Where vchAccountHeadCode  = '" & mHead & "') Begin " & _
                            "Declare @MaxID int; Select @MaxID = isNull(Max(intAccountHeadID),1) From faAccountHeads " & _
                            "Select @MaxID [intAccountHeadID],vchMinorAccountHead,intMinorAccountHeadID,intMajorAccountHeadID,tinType From faMinorAccountHeads Where left(vchMinorAccountHeadCode,6) = '" & CStr(cmbNatureofFunds.ItemData(mNFund)) + CStr(cmbBanks.ItemData(mBank)) & "' End Else Begin Select 0 [intAccountHeadID] End"
                    'Rec.Open mSQL, mCnn
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Rec!intAccountHeadID <> 0 Then
                        mArIn = Array(Null, mHead, Rec!vchMinorAccountHead + " _" + Left(CStr(mCount), 1), 0, _
                                    Rec!vchMinorAccountHead + " _" + Left(CStr(mCount), 1), Rec!intMinorAccountHeadID, _
                                    Rec!intMajorAccountHeadID, Null, Rec!tinType, gbLocalBodyID, gbFinancialYearID, 1, 0)
                        objdb.ExecuteSP "spSaveDetailedHead", mArIn, , , mCnn
                        mCnn.Execute "Update faAccountHeads Set tinHiddenFlag = 1,intGroupID = 2 Where vchAccountHeadCode = '" & mHead & "'"
                    End If
                    Rec.Close
                Next mCount
            Next mBank
        Next mNFund
    End Sub
    Private Sub DisplayBanksByAccount()
        Dim objBank As New clsBank
        Dim objAcc As New clsAccounts
        Dim objFund As New clsFund
        
        mEditFlag = False
        txtNameOfBank.Text = ""
        txtNameOfBank.Tag = ""
        txtBranch.Text = ""
        txtBankCode.Text = ""
        txtBranchCode.Text = ""
        txtAccountNo.Text = ""
        txtAddress1.Text = ""
        txtAddress2.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        txtOpeningBalance.Text = ""
        optCredit.value = False
        optDebit.value = True
        
        objBank.SetBankInfoByAccID (val(txtAccountHead.Tag))
        If objBank.BankID > 0 Then
            mEditFlag = True
            txtNameOfBank.Text = objBank.BankName
            txtNameOfBank.Tag = objBank.BankID
            txtBranch.Text = objBank.Branch
            txtBankCode.Text = objBank.BankCode
            txtBranchCode.Text = objBank.BranchCode
            txtAccountNo.Text = objBank.AccountNumber
            objAcc.SetAccountID objBank.BankAccountHeadID
            If objAcc.AccountHeadID > 0 Then
                txtAccountHead.Text = objAcc.AccountHead
                txtAccountHead.Tag = objBank.BankAccountHeadID
                txtAccountHeadCode.Text = objAcc.AccountCode
            Else
                txtAccountHead.Text = ""
                txtAccountHead.Tag = ""
                txtAccountHeadCode.Text = ""
            End If
            objFund.SetFund objBank.FundID
            If objFund.FundID > 0 Then
                txtFundCode.Text = objFund.FundCode
                txtFundCode.Tag = objFund.FundID
                txtFund.Text = objFund.FundName
            Else
                txtFundCode.Text = ""
                txtFundCode.Tag = ""
                txtFund.Text = ""
            End If
            txtAddress1.Text = objBank.Address1
            txtAddress2.Text = objBank.Address2
            txtPhone.Text = objBank.Phone
            txtFax.Text = objBank.Fax
            txtEmail.Text = objBank.Email
            txtOpeningBalance.Text = objBank.Opening
            If objBank.CrDrFlag Then
                optDebit.value = True
            Else
                optCredit.value = True
            End If
            
        End If
        lstBanks.Visible = False
    End Sub

