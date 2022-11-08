VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSubsidiaryAccount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subsidiary Accounts"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   11640
      Top             =   6450
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   465
      Left            =   4110
      TabIndex        =   12
      Top             =   5640
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   465
      Left            =   6630
      TabIndex        =   13
      Top             =   5640
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   465
      Left            =   5370
      TabIndex        =   11
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   90
      TabIndex        =   17
      Top             =   120
      Width           =   11625
      Begin VB.TextBox txtEditFlag 
         Height          =   285
         Left            =   11460
         TabIndex        =   30
         Top             =   4590
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ListBox lstHeads 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   6705
         TabIndex        =   14
         Top             =   1290
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.TextBox txtFlag 
         Height          =   285
         Left            =   11310
         TabIndex        =   29
         Top             =   4590
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txttemp 
         Height          =   285
         Left            =   11130
         TabIndex        =   28
         Top             =   4590
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtHeadID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4380
         TabIndex        =   27
         Top             =   1290
         Width           =   1035
      End
      Begin VB.CommandButton cmdSearchSub 
         Caption         =   "..."
         Height          =   315
         Left            =   6300
         TabIndex        =   3
         Top             =   1290
         Width           =   375
      End
      Begin VB.TextBox txtSubHead 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5100
         TabIndex        =   26
         Top             =   900
         Width           =   4575
      End
      Begin VB.TextBox txtMainCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4380
         TabIndex        =   25
         Top             =   900
         Width           =   345
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   3345
      End
      Begin VB.Frame frmeAddress 
         BackColor       =   &H00E0E0E0&
         Height          =   1395
         Left            =   3480
         TabIndex        =   22
         Top             =   2790
         Width           =   5385
         Begin VB.TextBox txtAddress1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   870
            MaxLength       =   100
            TabIndex        =   7
            Top             =   210
            Width           =   4245
         End
         Begin VB.TextBox txtAddress2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   870
            MaxLength       =   100
            TabIndex        =   8
            Top             =   540
            Width           =   4245
         End
         Begin VB.TextBox txtAddress3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   870
            MaxLength       =   100
            TabIndex        =   9
            Top             =   870
            Width           =   4245
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Height          =   225
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4380
         TabIndex        =   4
         Top             =   1695
         Width           =   4215
      End
      Begin VB.TextBox txtOpeningBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   4380
         TabIndex        =   5
         Top             =   2115
         Width           =   1995
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   315
         Left            =   9720
         TabIndex        =   2
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4740
         MaxLength       =   4
         TabIndex        =   1
         Top             =   900
         Width           =   345
      End
      Begin VB.Frame fraType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Type"
         Height          =   675
         Left            =   6570
         TabIndex        =   6
         Top             =   2100
         Width           =   2025
         Begin VB.OptionButton optCreditors 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   240
            TabIndex        =   16
            Top             =   420
            Width           =   735
         End
         Begin VB.OptionButton optDebtors 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Debt"
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
            Height          =   240
            Left            =   240
            TabIndex        =   15
            Top             =   210
            Width           =   705
         End
      End
      Begin VB.TextBox txtHead 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5430
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1290
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Category"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   1740
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opening Balance"
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
         Left            =   2940
         TabIndex        =   20
         Top             =   2145
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Code"
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
         Left            =   3870
         TabIndex        =   18
         Top             =   930
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subsidiary Account Head"
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
         Left            =   2280
         TabIndex        =   19
         Top             =   1290
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSubsidiaryAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '=================================================================================='
    '                              Subsidiary Accounts Form                            '
    '=================================================================================='
    '
    '----------------------------------------------------------------------------------'
    '       Coaded By               :       Cijith Sreedharan                          '
    '       Date                    :                                                  '
    '       Stored Procedure Used   :       spSaveSubsidiaryHead                                                                                  '
    '                                                                                  '
    '----------------------------------------------------------------------------------'
    
    Option Explicit
    Dim mEditFlag As Boolean
    Dim mSubID          As Variant
    Dim mNewSubID       As Variant
    Dim SubAccountCode  As Variant
    Dim SubHeadCode As Integer
    Private Function checkNumeric(mChar As Integer) As Integer
        If (mChar > 47 And mChar < 58) Or mChar = 8 Or mChar = 44 Or mChar = 46 Then
        Else
            mChar = 0
        End If
        checkNumeric = mChar
    End Function
    Private Sub FormClear()
        'cmbCategory.ListIndex = -1
        lstHeads.Visible = False
        txtMainCode.Text = ""
        txtCode.Text = ""
        txtHeadID.Text = ""
        txtHead.Text = ""
        txtName.Text = ""
        txtOpeningBalance.Text = ""
        optCreditors.Value = False
        optDebtors.Value = False
        txtAddress1.Text = ""
        txtAddress2.Text = ""
        txtAddress3.Text = ""
        txtEditFlag.Text = ""
        txtSubHead.Text = ""
    End Sub
    
    Private Sub FormInitialize()
        txtCode.Enabled = False
        txtCode.Text = ""
        txtCode.Tag = ""
        txtHead.Text = ""
        txtAddress1.Text = ""
        txtAddress2.Text = ""
        txtAddress3.Text = ""
        txtOpeningBalance.Text = ""
        optCreditors = False
        txtName.Text = ""
        optDebtors = False
        mEditFlag = False
    End Sub
    
    Private Sub cmbCategory_Click()
        If cmbCategory.ListIndex <> -1 Then
            txtMainCode.Text = cmbCategory.ItemData(cmbCategory.ListIndex) * 10
            lstHeads.Visible = False
        End If
    End Sub
    
    Private Sub cmdCancel_Click()
        'Call FormInitialize
        Call FormClear
    End Sub
    
    Private Sub cmdNew_Click()
        Dim objDB       As New clsDB
        Dim mCon        As ADODB.Connection
        Dim Rec         As ADODB.Recordset
        Dim rs          As New ADODB.Recordset
        Dim mHeadID     As Variant
        Dim mQry        As String
        Dim temp2       As Variant
        Dim temp3       As Integer
        
        txtHead.Visible = True
        txtName.Text = ""
        txtOpeningBalance.Text = ""
        txtAddress1.Text = ""
        txtAddress2.Text = ""
        txtAddress3.Text = ""
                   
        frmeAddress.Visible = True
        mEditFlag = False
        
        txtHeadID.Enabled = False
        SubHeadCode = (txtMainCode.Text) + (txtCode.Text)
        txtHeadID.Text = SubHeadCode
        
        Set Rec = New ADODB.Recordset
        objDB.SetConnection mCon
        
        mQry = "Select Isnull( Max( Cast( vchSubAccountCode as Int)) ,1) as No From faSubsidiaryAccounts where vchSubHeadCode =  " & SubHeadCode
        Rec.Open mQry, mCon
            txttemp.Text = mID(CStr(Rec!no), 5)
            temp3 = Len(txttemp.Text)
            If Rec(0) = 1 Then
                temp2 = 0
            Else
                temp2 = Rec(0)
            End If
            If Rec!no <> 1 Then
                txtHead.Text = Right(Val(temp2 + 1), temp3)
            Else
                txtHead.Text = 1
            End If
        Rec.Close
        txttemp.Text = SubHeadCode
        SubAccountCode = txttemp.Text + txtHead.Text

    End Sub
    
    Private Sub cmdSave_Click()

        Dim objDB               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim arrInput            As Variant
        Dim mCategoryID         As Integer
        Dim mAccountCode        As Variant
        Dim mAccountHead        As Variant
        Dim rs                  As New ADODB.Recordset
        Dim mQry                As String
        Dim Recs                As New ADODB.Recordset
        Dim mSubAccountID       As Integer
        Dim mCategory           As Integer
        
        
        '---------------------------------------------------'
        '  Validations                                      '
        '---------------------------------------------------'
            objDB.SetConnection mCnn
        
        '---------------------------------------------------'
        '  Getting Primary Key                                          '
        '---------------------------------------------------'
            mQry = "Select Isnull( Max(intSubAccountID) ,1) as No From faSubsidiaryAccounts "
            Recs.Open mQry, mCnn
            mSubAccountID = Recs!no + 1
            Recs.Close
               
        '---------------------------------------------------'
        '  Updating                                          '
        '---------------------------------------------------'
             
        If optCreditors.Value = True Then mCategory = 0
        If optDebtors.Value = True Then mCategory = 1

        arrInput = Array(txtEditFlag.Text, _
                     IIf(txtEditFlag.Text = "", Trim(SubAccountCode), txtHeadID.Text), _
                    Trim(txtName.Text), _
                    IIf(IsNull(txtAddress1.Text), Null, txtAddress1.Text), _
                    IIf(IsNull(txtAddress2.Text), Null, txtAddress2.Text), _
                    IIf(IsNull(txtAddress3.Text), Null, txtAddress3.Text), _
                    IIf(txtEditFlag.Text = "", Trim(SubHeadCode), (mID(CStr(txtHeadID.Text), 1, 4))), _
                    mCategory, _
                    IIf(IsNull(txtOpeningBalance.Text), Null, Trim(txtOpeningBalance.Text)) _
                    )

        objDB.ExecuteSP "spSaveSubsidiaryHead", arrInput, , , mCnn
        MsgBox "Saved Successfully", vbInformation
        Call FormClear
        cmbCategory.ListIndex = -1
        txtHead.Text = ""
        txtHeadID.Text = ""
        txtName.Text = ""
        txtMainCode.Text = ""
        txtSubHead.Text = ""
'        Unload Me
'        frmSubsidiaryAccount.Show
    End Sub
    
    Private Sub cmdSearch_Click()
        If cmbCategory.ListIndex <> -1 Then
            Dim mSQL As String
            mSQL = "Select (vchSubHeadCode + '      ' + vchSubHead) , numSubID from faSubsidaryAccountCategory where tinType = " & cmbCategory.ItemData(cmbCategory.ListIndex) & " Order By vchSubHeadCode"
            Call PopulateList(lstHeads, mSQL, , , , True)
            txtFlag.Text = 1
            lstHeads.Visible = True
            lstHeads.ZOrder 0
            lstHeads.SetFocus
        Else
            MsgBox "Please Select the Category from the Combo Box", vbCritical
        End If
    End Sub
    
    Private Sub cmdSearchSub_Click()
        Dim mSQL As String
        mSubID = txtMainCode.Text + txtCode.Text
        mSQL = "Select (vchSubAccountCode + '       ' + vchSubAccountHead) , intSubAccountID from faSubsidiaryAccounts where vchSubHeadCode = " & mSubID & " Order By vchSubAccountCode"
        Call PopulateList(lstHeads, mSQL, , , , True)
        txtFlag.Text = 2
        lstHeads.Visible = True
        lstHeads.SetFocus
    End Sub

    Private Sub Form_Activate()
'        frmSubsidiaryAccount.Top = 1700
'        frmSubsidiaryAccount.Left = (frmMenu.Width - Me.Width) / 2
        frmeAddress.Visible = False
        fillCategoryCombo
    End Sub
    
    Private Sub Form_Load()
       ' Call FormInitialize
        Call FormClear
        lstHeads.Visible = False
        WindowsXPC1.InitIDESubClassing
    End Sub
    
    Private Sub Frame1_Click()
        lstHeads.Visible = False
    End Sub
    

    Private Sub optSubAccount_Click()
        'If optSubAccount.Value = True Then
        'txtAccountCode.Visible = True
        'txtAccountHead.Visible = True
        'lblAccountCode.Visible = True
        'lblAccountHead.Visible = True
        'ElseIf optSubAccount.Value = False Then
        'HideforSubAccount
        'End If
    End Sub
    
    Private Sub optTag_Click()
        'HideforSubAccount
    End Sub
    Private Sub fillCategoryCombo()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim mSQL As String
        
        mSQL = "Select tinType,vchCategory from faSubsidaryAccountMajorCategory"
        
        objDB.SetConnection mCnn
        rs.Open mSQL, mCnn, adOpenDynamic, adLockPessimistic
        Do Until rs.EOF
            cmbCategory.AddItem rs(1)
            cmbCategory.ItemData(cmbCategory.NewIndex) = rs(0)
            rs.MoveNext
        Loop
        
    End Sub


    Private Sub txtAddress1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 39 Then KeyAscii = 0
    End Sub

    Private Sub txtAddress2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 39 Then KeyAscii = 0
    End Sub

    Private Sub txtAddress3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 39 Then KeyAscii = 0
    End Sub

    Private Sub txtCode_Click()
        lstHeads.Visible = False
    End Sub

    Private Sub txtCode_LostFocus()
        Dim objDB       As New clsDB
        Dim mCon        As ADODB.Connection
        Dim Rec         As ADODB.Recordset
        Dim mHeadID     As Variant
        
        Set Rec = New ADODB.Recordset
        txtCode.Text = txtCode.Text
        If txtCode.Text <> "" Then
           mEditFlag = True
           objDB.SetConnection mCon
           Rec.Open "Select * from faSubsidiaryAccounts where faSubsidiaryAccounts.vchSubAccountCode = " & (txtCode.Text), mCon
           On Error Resume Next
           If Not (Rec.EOF Or Rec.BOF) Then
               txtCode.Text = Rec!vchSubAccountCode
               txtCode.Tag = Rec!intSubAccountID
               'txtHead.Text = rec!vchSubAccountHead
               txtAddress1.Text = Rec!vchAddress1
               txtAddress2.Text = Rec!vchAddress2
               txtAddress3.Text = Rec!vchAddress3
               If Rec!tinType Then
                   optDebtors.Value = True
               Else
                   optCreditors.Value = True
               End If
           Else
               mEditFlag = False
               txtCode.Text = ""
               txtCode.Tag = ""
               txtHead.Text = ""
               txtAddress1.Text = ""
               txtAddress2.Text = ""
               txtAddress3.Text = ""
               optCreditors.Value = False
               optDebtors.Value = False
           End If
           On Error GoTo 0
           Rec.Close
        End If
        
    End Sub
    Private Sub lstHeads_DblClick()
        Dim mSearchStr      As String
        Dim mSearchID       As Variant
        Dim mCharCnt As Integer
        'Dim mSubID          As Variant
        Dim mStrCnt         As Integer
        '-------------------------------------------------------------------------'
        Dim mSQL As String
        Dim mCnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim objDB As New clsDB
        '--------------------------------------------------------------------------'
            If lstHeads.ListIndex > -1 Then
                    mSearchStr = lstHeads.Text
                    mSearchID = lstHeads.ItemData(lstHeads.ListIndex)
                    If mSearchStr <> "" Then
                        mCharCnt = InStr(mSearchStr, " ")
                        mSubID = Left(mSearchStr, mCharCnt)
                        If Val(txtFlag.Text) = 1 Then
                            txtCode.Text = Right(mSubID, 3)
                            txtCode.Tag = mSearchID
                            mStrCnt = InStr(mSearchStr, " ")
                            txtSubHead.Text = mID(mSearchStr, mStrCnt)
                            mSearchStr = ""
                        Else
                            txtHeadID.Text = mSubID 'Right(mSubID, 3)
                            txtCode.Tag = mSearchID
                            mStrCnt = InStr(mSearchStr, " ")
                            txtName.Text = mID(mSearchStr, mStrCnt)
                            mSearchStr = ""
                            txtHead.Visible = False
           '--------------------------------------------------------------------------'
                                       ' Displaying the Address '
                                       frmeAddress.Visible = True
                                       objDB.SetConnection mCnn
                                       mSQL = "Select intSubAccountID,vchAddress1,vchAddress2,vchAddress3,intCategoryID,fltOpeningBalance from faSubsidiaryAccounts Where vchSubAccountCode =  " & txtHeadID.Text
                                       rs.Open mSQL, mCnn, adOpenStatic, adLockPessimistic
                                       txtEditFlag.Text = rs!intSubAccountID
                                       txtAddress1.Text = IIf(IsNull(rs!vchAddress1), "", rs!vchAddress1)
                                       txtAddress2.Text = IIf(IsNull(rs!vchAddress2), "", rs!vchAddress2)
                                       txtAddress3.Text = IIf(IsNull(rs!vchAddress3), "", rs!vchAddress3)
                                       txtOpeningBalance.Text = IIf(IsNull(rs!fltOpeningBalance), "", rs!fltOpeningBalance)
                                       If rs!intCategoryID = 1 Then optDebtors.Value = True
                                       If rs!intCategoryID = 0 Then optCreditors.Value = True
                                       rs.Close
           '--------------------------------------------------------------------------'
       
                        End If
                        mSearchID = -1
                        'Call txtCode_LostFocus
                    End If
            End If
            lstHeads.Visible = False
    End Sub
   
    Private Sub lstHeads_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call lstHeads_DblClick
        End If
    End Sub
    
    Private Sub txtName_KeyPress(KeyAscii As Integer)
        If KeyAscii = 39 Then KeyAscii = 0
    End Sub
    
    Private Sub txtOpeningBalance_KeyPress(KeyAscii As Integer)
        KeyAscii = checkNumeric(KeyAscii)
    End Sub
    Private Sub txtOpeningBalance_LostFocus()
        txtOpeningBalance.Text = Format(txtOpeningBalance.Text, "0.00")
    End Sub
