VERSION 5.00
Begin VB.Form frmAccountHeads 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " A c c o u n t   H e a d s"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstGroups 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   4710
      TabIndex        =   41
      Top             =   2220
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   5850
      TabIndex        =   38
      Top             =   6090
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4620
      TabIndex        =   37
      Top             =   6090
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
      ScaleWidth      =   10890
      TabIndex        =   39
      Top             =   0
      Width           =   10890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5340
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   10890
      Begin VB.ListBox lstHeads 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   4710
         TabIndex        =   42
         Top             =   1560
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ListBox lstMasters 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7500
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkSecondary 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Secondary Account Heads"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2760
         TabIndex        =   24
         Top             =   2760
         Width           =   2595
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2610
         TabIndex        =   33
         Top             =   4335
         Width           =   7695
         Begin VB.OptionButton optDebit 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5130
            TabIndex        =   35
            Top             =   315
            Width           =   780
         End
         Begin VB.OptionButton optCredit 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5985
            TabIndex        =   36
            Top             =   315
            Width           =   855
         End
         Begin VB.TextBox txtOpening 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   34
            Top             =   270
            Width           =   3015
         End
      End
      Begin VB.Frame fraPrimary 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Primary Account Heads"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   2610
         TabIndex        =   6
         Top             =   540
         Width           =   7680
         Begin VB.TextBox txtMinorByDetailHide 
            Height          =   315
            Left            =   5490
            TabIndex        =   44
            Top             =   1320
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox txtGroup 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   22
            Top             =   1650
            Width           =   2955
         End
         Begin VB.TextBox txtPrimaryAlias 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   20
            Top             =   1320
            Width           =   2955
         End
         Begin VB.CommandButton cmdSearchHead 
            Caption         =   "..."
            Height          =   300
            Left            =   7080
            TabIndex        =   18
            Top             =   990
            Width           =   375
         End
         Begin VB.CommandButton cmdMinorSearch 
            Caption         =   "..."
            Height          =   300
            Left            =   7080
            TabIndex        =   14
            Top             =   645
            Width           =   375
         End
         Begin VB.TextBox txtMajorCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Top             =   330
            Width           =   1560
         End
         Begin VB.CommandButton cmdMajorSearch 
            Caption         =   "..."
            Height          =   300
            Left            =   7080
            TabIndex        =   10
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox txtMajorHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3390
            TabIndex        =   9
            Top             =   330
            Width           =   3615
         End
         Begin VB.TextBox txtMinorHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3390
            TabIndex        =   13
            Top             =   660
            Width           =   3615
         End
         Begin VB.TextBox txtMinorCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   660
            Width           =   1560
         End
         Begin VB.TextBox txtDetailedHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3390
            TabIndex        =   17
            Top             =   990
            Width           =   3615
         End
         Begin VB.TextBox txtDetailedCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   16
            Top             =   990
            Width           =   1560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primary Group "
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   435
            TabIndex        =   21
            Top             =   1680
            Width           =   1380
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alias"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1320
            TabIndex        =   19
            Top             =   1350
            Width           =   435
         End
         Begin VB.Label lblprimaryachead 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minor Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            TabIndex        =   11
            Top             =   675
            Width           =   1020
         End
         Begin VB.Label lblsecondaryhead 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   765
            TabIndex        =   7
            Top             =   345
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Detailed Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   555
            TabIndex        =   15
            Top             =   1020
            Width           =   1215
         End
      End
      Begin VB.Frame fraSecondary 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Secondary Account Heads"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   2610
         TabIndex        =   23
         Top             =   2775
         Width           =   7695
         Begin VB.TextBox txtSecondaryCodeHide 
            Height          =   285
            Left            =   4470
            LinkItem        =   "v"
            TabIndex        =   43
            Top             =   390
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.TextBox txtPrimaryCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1785
            TabIndex        =   26
            Top             =   375
            Width           =   1560
         End
         Begin VB.TextBox txtSecondaryAlias 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1785
            TabIndex        =   32
            Top             =   1065
            Width           =   3000
         End
         Begin VB.CommandButton cmdSearchSecondary 
            Caption         =   "..."
            Height          =   300
            Left            =   7080
            TabIndex        =   30
            Top             =   705
            Width           =   375
         End
         Begin VB.TextBox txtSecondaryHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1770
            TabIndex        =   29
            Top             =   720
            Width           =   5205
         End
         Begin VB.TextBox txtSecondaryCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            LinkTimeout     =   0
            MaxLength       =   4
            TabIndex        =   27
            Top             =   375
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secondary Code"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   390
            TabIndex        =   25
            Top             =   405
            Width           =   1365
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alias"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1305
            TabIndex        =   31
            Top             =   1095
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Secondary Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   345
            TabIndex        =   28
            Top             =   780
            Width           =   1410
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Type of Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   570
         TabIndex        =   1
         Top             =   540
         Width           =   1965
         Begin VB.OptionButton optAsset 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Asset"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   450
            TabIndex        =   5
            Top             =   1125
            Width           =   1470
         End
         Begin VB.OptionButton optLiability 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Liability"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   450
            TabIndex        =   4
            Top             =   870
            Width           =   1470
         End
         Begin VB.OptionButton optExpenditure 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Expenditure"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   450
            TabIndex        =   3
            Top             =   615
            Width           =   1470
         End
         Begin VB.OptionButton optIncome 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Income"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   450
            TabIndex        =   2
            Top             =   360
            Width           =   1470
         End
      End
   End
End
Attribute VB_Name = "frmAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*****************************************************************************************
'* Application ID           :                                                            *
'* Application Name         : Saankhya Double Entry                                      *
'* Screen id                : Payments                                                   *
'* Version No               : Ver 2.0.0                                                  *
'* Form Designed By         : Aiby                                                       *
'* Created on               :                                                            *
'* Coded By                 :                                                            *
'* Coded on                 :                                                            *
'* Reviewed By              :                                                            *
'* Reviewed on              : 16-Oct-2007                                                *
'* Purpose                  : To define and modify Account Heads                         *
'*                                                                                       *
'*                                                                                       *
'* Name of Database         : DB_Finance                                                 *
'* DSN                      : dsnFA ( UserName=FAUser; PWD=FAUser )                      *
'* Name of Table(s)         :                                                            *
'* Look up Table(s)         :                                                            *
'*                          :                                                            *
'*                                                                                       *
'* Stored Procedures        :                                                            *
'*                          :                                                            *
'*                                                                                       *
'*=======================================================================================*
    Option Explicit
    Dim mSearchID As Variant
    Dim mGroupID As Variant
    Dim mEditFlag As Boolean

    Private Sub ListAccountHeads(mListBy As Long)
        Dim mTypeID As Variant
        Dim mMinorHeadID As Variant
        If optIncome.Value Then mTypeID = 1
        If optExpenditure.Value Then mTypeID = 2
        If optLiability.Value Then mTypeID = 3
        If optAsset.Value Then mTypeID = 4
        If Not IsNumeric(mTypeID) Then mTypeID = Null
        lstMasters.TabIndex = mListBy
        Select Case mListBy
            Case 1
                lstMasters.Tag = 1
                Call PopulateList(lstMasters, "spGetMajorAccountHeads " & mTypeID, , , , True)
            Case 2
                lstMasters.Tag = 2
                If Val(txtMajorCode.Tag) > 0 Then
                    Call PopulateList(lstMasters, "spGetMinorAccountHeads " & mTypeID & ", " & Val(txtMajorCode.Tag), , , , True)
                Else
                    Call PopulateList(lstMasters, "spGetMinorAccountHeads " & mTypeID, , , , True)
                End If
        End Select
        lstMasters.Left = 7500
        lstMasters.Top = 450
        lstMasters.Width = 3200
        lstMasters.Height = 4600
        lstMasters.Visible = True
        lstMasters.SetFocus
    End Sub

    Private Sub FormInitialize()
        txtMajorCode.Text = ""
        txtMajorCode.Tag = ""
        txtMajorHead.Text = ""
        txtMajorHead.Tag = ""
        txtMinorCode.Text = ""
        txtMinorCode.Tag = ""
        txtMinorHead.Text = ""
        txtMinorHead.Tag = ""
        txtDetailedCode.Text = ""
        txtDetailedCode.Tag = ""
        txtDetailedHead.Text = ""
        txtPrimaryAlias.Text = ""
        txtGroup.Text = ""
        txtPrimaryCode.Text = ""
        txtSecondaryCode.Text = ""
        txtSecondaryHead.Text = ""
        txtSecondaryAlias.Text = ""
        txtOpening.Text = ""
        optDebit.Value = True
        optIncome.Value = True
        mEditFlag = False
        chkSecondary.Value = 0
    End Sub

    Private Sub ShowSearchAccountHead()
        Dim mSQL As String
        If Val(txtMinorCode.Tag) > 0 Then
            mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intMinorAccountHeadID = " & Val(txtMinorCode.Tag)
        ElseIf Val(txtMajorCode.Tag) > 0 Then
            mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intMajorAccountHeadID = " & Val(txtMajorCode.Tag)
        End If
        Dim mTypeID As Variant
        If optIncome.Value Then mTypeID = 1
        If optExpenditure.Value Then mTypeID = 2
        If optLiability.Value Then mTypeID = 3
        If optAsset.Value Then mTypeID = 4
        If Not IsNumeric(mTypeID) Then mTypeID = Null
        
        Select Case mTypeID
            Case 1
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 1 And tinSecondaryAccountFlag=0"
            Case 2
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 2 And tinSecondaryAccountFlag=0"
            Case 3
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 3 And tinSecondaryAccountFlag=0"
            Case 4
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 4 And tinSecondaryAccountFlag=0"
         End Select
         
        frmSearchAccountHeads.SQLString = mSQL
        frmSearchAccountHeads.Show vbModal
        txtDetailedCode.SetFocus
    End Sub

    Private Sub chkSecondary_Click()
        fraSecondary.Enabled = chkSecondary.Value
        If chkSecondary.Value Then
            mEditFlag = False
            txtPrimaryCode.Text = Val(Trim(txtDetailedCode.Text))
            txtPrimaryCode.Enabled = False
            txtSecondaryCode.SetFocus
            txtSecondaryCode.Text = Trim(txtSecondaryCode.Text)
            txtSecondaryHead.Text = Trim(txtSecondaryHead.Text)
            txtSecondaryAlias.Text = Trim(txtSecondaryAlias.Text)
        Else
            mEditFlag = False
            Call FormInitialize
        End If
    End Sub

    Private Sub cmdCancel_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdMajorSearch_Click()
        Call ListAccountHeads(1)
    End Sub

    Private Sub cmdMinorSearch_Click()
        Call ListAccountHeads(2)
    End Sub

    Private Sub cmdNew_Click()
'        mEditFlag = False
'        txtSecondaryCode.Text = ""
'        txtSecondaryHead.Text = ""
'        txtSecondaryAlias.Text = ""
'        txtGroup.Text = ""
'        txtOpening.Text = ""
    End Sub

    Private Sub cmdSave_Click()
            Dim objDB                   As New clsDB
            Dim objAcc                  As New clsAccounts
            Dim objAc                   As New clsAccounts
            Dim mCnn                    As New ADODB.Connection
            Dim mDetailedAccountHeadID  As Long
            Dim mMajorAccountHeadID     As Long
            Dim mMinorAccountHeadID     As Long
            Dim mDetailedCode           As String
            Dim mSecondaryCode          As String
            Dim mTypeID                 As Long
            Dim mAmt                    As Double
            Dim arrInput                As Variant
            Dim Rec                     As New ADODB.Recordset
            Dim ArrIn                   As Variant
            '---------------------------------------------------'
            '  Validations                                      '
            '---------------------------------------------------'
            objAcc.SetAccountID (Val(txtDetailedCode.Tag))
            mDetailedAccountHeadID = objAcc.AccountHeadID
            If mDetailedAccountHeadID = -1 Then
                MsgBox "Select a Primary account head", vbInformation
                Call FormInitialize
                chkSecondary.Value = 0
                fraPrimary.Enabled = True
                Call cmdSearchHead_Click
                Exit Sub
            Else
                mDetailedAccountHeadID = objAcc.AccountHeadID
                mMajorAccountHeadID = objAcc.MajorAccountHeadID
                mMinorAccountHeadID = objAcc.MinorAccountHeadID
                mTypeID = objAcc.mType
                mDetailedCode = objAcc.AccountCode
            End If
            If optCredit.Value Then
                mAmt = Abs(Val(txtOpening.Text)) * -1
            Else
                mAmt = Abs(Val(txtOpening.Text))
            End If
            If chkSecondary Then
                If Trim(txtSecondaryCode) = "" Then
                    txtSecondaryCode.SetFocus
                    Exit Sub
                End If
                If Trim(txtSecondaryHead) = "" Then
                    txtSecondaryHead.SetFocus
                    Exit Sub
                End If
            End If
          
            '---------------------------------------------------'
            '  Updating                                         '
            '---------------------------------------------------'
            If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                
                
                If chkSecondary Then
                    arrInput = Array(IIf(mEditFlag, Val(txtSecondaryCode.Tag), Null), _
                    Trim(txtDetailedCode.Text) & Trim(txtSecondaryCode.Text), _
                    Trim(txtSecondaryHead.Text), _
                    1, _
                    Format(mAmt, "0.00"), _
                    mMinorAccountHeadID, _
                    mMajorAccountHeadID, _
                    mGroupID, _
                    mTypeID, _
                    gbLocalBodyID, _
                    gbFinancialYearID, _
                    IIf(optDebit, 1, 0), _
                    Trim(txtSecondaryAlias.Text), _
                    mDetailedAccountHeadID _
                    )
                Else
                    arrInput = Array(IIf(mEditFlag, Val(txtDetailedCode.Tag), Null), _
                    mDetailedCode, _
                    Trim(txtDetailedHead.Text), _
                    0, _
                    Format(mAmt, "0.00"), _
                    mMinorAccountHeadID, _
                    mMajorAccountHeadID, _
                    mGroupID, _
                    mTypeID, _
                    gbLocalBodyID, _
                    gbFinancialYearID, _
                    IIf(optDebit, 1, 0), _
                    Trim(txtPrimaryAlias.Text), _
                    mDetailedAccountHeadID _
                    )
                End If
                objDB.ExecuteSP "spSaveSecondaryHead", arrInput, , , mCnn
                
                Dim mSQL As String
                 mSQL = "Select Count(*) From faTransactions Where intTransactionID = 0"
                Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
                If Rec.Fields(0).Value = 0 Then
                    Dim intTransactionID_1   As Double
                    Dim mintLocalBodyID_2  As Long
                    Dim mintFinancialYearID_3  As Long
                    Dim mdtTransactionDate_4   As Date
                    Dim mintExternalApplicationID_5    As Long
                    Dim mintExternalApplicationModuleID_6  As Long
                    Dim mintFunctionID_7   As Variant
                    Dim mintFunctionaryID_8   As Variant
                    Dim mintFieldID_9 As Variant
                    Dim mintFundID_10 As Variant
                    Dim mintBudgetCentreID_11  As Variant
                    Dim mvchNarration_12   As String
                    Dim mintTransactionTypeID_13   As Variant
                    Dim mintVoucherNo_14   As Variant
                    Dim mintProcessID_15    As Variant
                    Dim mintGroupID_17    As Variant
                    Dim mvchGroup_16   As String
                    Dim mintKeyID_18   As Variant
                    Dim mnumSubLedgerID_19    As Variant
                    Dim mintUserID_20  As Variant
                    
                    intTransactionID_1 = 0
                    mintLocalBodyID_2 = gbLocalBodyID
                    mintFinancialYearID_3 = gbFinancialYearID
                    mdtTransactionDate_4 = gbStartingDate
                    mintExternalApplicationID_5 = AppID.Saankhya
                    mintExternalApplicationModuleID_6 = 0
                    mintFunctionID_7 = Null
                    mintFunctionaryID_8 = Null
                    mintFieldID_9 = Null
                    mintFundID_10 = Null
                    mintBudgetCentreID_11 = Null
                    mvchNarration_12 = "Opening Balance"
                    mintTransactionTypeID_13 = Null
                    mintVoucherNo_14 = Null
                    mintProcessID_15 = Null
                    mvchGroup_16 = "JV"
                    mintGroupID_17 = 40
                    mintKeyID_18 = Null
                    mnumSubLedgerID_19 = Null
                    mintUserID_20 = 0
                    
                    arrInput = Array( _
                    intTransactionID_1, _
                    mintLocalBodyID_2, _
                    mintFinancialYearID_3, _
                    mdtTransactionDate_4, _
                    mintExternalApplicationID_5, _
                    mintExternalApplicationModuleID_6, _
                    mintFunctionID_7, _
                    mintFunctionaryID_8, _
                    mintFieldID_9, _
                    mintFundID_10, _
                    mintBudgetCentreID_11, _
                    mvchNarration_12, _
                    mintTransactionTypeID_13, _
                    mintProcessID_15, _
                    mvchGroup_16, _
                    mintGroupID_17, _
                    mintKeyID_18, _
                    mnumSubLedgerID_19, _
                    gbUserID, _
                    mintVoucherNo_14)
                    
                    objDB.ExecuteSP "spSaveTransactions", arrInput, , , mCnn
                End If
                
                arrInput = Array(0, _
                    2, _
                    Val(txtDetailedCode.Tag), _
                    Format(mAmt, "0.00"), _
                    IIf(optDebit, 1, 0), _
                    Null, _
                    "Opening Balance", _
                    Null _
                    )
                objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            End If
            Call FormInitialize
    End Sub

    Private Sub cmdSearchHead_Click()
        Call ShowSearchAccountHead
        txtDetailedCode.SetFocus
    End Sub

    Private Sub cmdSearchSecondary_Click()
        Dim mSQL As String
        mSQL = "Select (vchAccountHeadCode + '       ' + vchAccountHead) , intAccountHeadID from faAccountHeads Where faAccountHeads.tinSecondaryAccountFlag=1 Order By vchAccountHeadCode"
        Call PopulateList(lstHeads, mSQL, , , , True)
        lstHeads.Visible = True
        lstHeads.SetFocus
    End Sub

    Private Sub Form_Activate()
        frmAccountHeads.Top = 0
        frmAccountHeads.Left = (frmMenu.Width - Me.Width) / 2
    End Sub

    Private Sub Form_Load()
        Call FormInitialize
    End Sub

    Private Sub lstGroups_DblClick()
        If lstGroups.ItemData(lstGroups.ListIndex) > 0 Then
            mGroupID = lstGroups.ItemData(lstGroups.ListIndex)
        End If
        Call lstGroups_LostFocus
        lstGroups.Visible = False
    End Sub

    Private Sub lstGroups_LostFocus()
        If IsNumeric(mGroupID) Then
                Dim objDB As New clsDB
                Dim mCon As New ADODB.Connection
                Dim Rec As New ADODB.Recordset
                objDB.SetConnection mCon
                If mGroupID > 0 Then ' Changed by Aiby - Modify this later
                    Rec.Open "Select * from faAccountGroups Where faAccountGroups.intGroupId = " & mGroupID, mCon
                    'Else
                    'Rec.Open "Select * from faAccountGroups", mCon
                    'End If
                    If Not (Rec.BOF And Rec.EOF) Then
                        txtGroup.Text = Rec!vchGroup
                        txtGroup.Tag = Rec!intGroupID
                    Else
                        txtGroup.Text = ""
                        txtGroup.Tag = ""
                    End If
                End If
        End If
        lstGroups.Visible = False
    '    Rec.Close
    End Sub

    Private Sub lstHeads_DblClick()
        Dim mSearchStr      As String
        Dim mSearchID   As Variant
        Dim mCharCnt As Integer
        Dim objAc As New clsAccounts
        Dim objDB As New clsDB
        Dim mCon As New ADODB.Connection
        Dim Rec1 As New ADODB.Recordset
        Dim Rec2 As New ADODB.Recordset
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim arInput As Variant
    
        If lstHeads.ListIndex > -1 Then
            mSearchStr = lstHeads.Text
            mSearchID = lstHeads.ItemData(lstHeads.ListIndex)
            If mSearchStr <> "" Then
                txtSecondaryCode.Tag = mSearchID
                objDB.SetConnection mCon
                Rec.CursorLocation = adUseClient
                Rec.Open "spGetSecondaryDetailsForDisplay " & mSearchID, mCon, adOpenStatic, adLockOptimistic
                If Not (Rec.BOF And Rec.EOF) Then
                    On Error Resume Next ' For Testing Purpose
                    txtMajorCode.Text = Rec!vchMajorAccountHeadCode
                    txtMajorHead.Text = Rec!vchMajorAccountHead
                    txtMajorCode.Tag = Rec!intMajorAccountHeadID
                    txtMinorCode.Text = Rec!vchMinorAccountHeadCode
                    txtMinorHead.Text = Rec!vchMinorAccountHead
                    txtMinorCode.Tag = Rec!intMinorAccountHeadID
                    txtOpening.Text = Abs(Rec!fltOpeningBalance)
                    If (Rec!tinDebitOrCredit) = 1 Then
                        optDebit = True
                    Else
                        optCredit = True
                    End If
                    txtGroup.Text = Rec!vchGroup
                    On Error GoTo 0
                End If
                Rec.Close
                Rec1.Open "Select * from faAccountHeads Where faAccountHeads.intAccountHeadID=" & mSearchID, mCon
                If Not (Rec1.BOF And Rec1.EOF) Then
                    txtPrimaryCode.Tag = Rec1!intPrimaryHeadID
                    txtSecondaryHead.Text = Rec1!vchAccountHead
                    txtSecondaryAlias.Text = Rec1!vchAlias
                    txtSecondaryCodeHide.Text = Rec1!vchAccountHeadCode
                    txtSecondaryCode.Tag = Rec1!intAccountHeadID
                End If
                Rec1.Close
                Rec2.Open "Select * from faAccountHeads Where faAccountHeads.intAccountHeadID=" & Val(txtPrimaryCode.Tag), mCon
                If Not (Rec2.EOF And Rec2.BOF) Then
                    txtPrimaryCode.Text = Rec2!vchAccountHeadCode
                    txtDetailedCode.Text = Rec2!vchAccountHeadCode
                    txtDetailedHead.Text = Rec2!vchAccountHead
                    txtDetailedCode.Tag = Rec2!intAccountHeadID
                    If IsNull(Rec2!vchAlias) Then
                        txtPrimaryAlias.Text = ""
                    Else
                        txtPrimaryAlias.Text = Rec2!vchAlias
                    End If
                End If
                Rec2.Close
                txtSecondaryCode.Text = Right(txtSecondaryCodeHide.Text, Len(Trim(txtSecondaryCodeHide.Text)) - Len(Trim(txtPrimaryCode.Text)))
                mSearchStr = ""
                mSearchID = -1
                mEditFlag = True
            End If
        End If
        lstHeads.Visible = False
        txtPrimaryCode.Enabled = False
        txtSecondaryCode.SetFocus
    End Sub

    Private Sub lstHeads_LostFocus()
        lstHeads.Visible = False
    End Sub

    Private Sub lstMasters_DblClick()
        If lstMasters.ItemData(lstMasters.ListIndex) > 0 Then
            mSearchID = lstMasters.ItemData(lstMasters.ListIndex)
        End If
        Call lstMasters_LostFocus
    End Sub
    
    Private Sub lstMasters_LostFocus()
        lstMasters.Visible = False
        Select Case Val(lstMasters.Tag)
            Case 1: txtMajorCode.SetFocus
                    Call txtMajorCode_GotFocus
            Case 2: txtMinorCode.SetFocus
                    Call txtMinorCode_GotFocus
        End Select
        lstMasters.Tag = ""
    End Sub

    Private Sub txtGroup_GotFocus()
        Dim mSQL As String
        mSQL = "Select vchGroup,intGroupId From faAccountGroups Order By vchGroup"
        Call PopulateList(lstGroups, mSQL, , , , True)
        lstGroups.Visible = True
        lstGroups.SetFocus
    End Sub

    Private Sub txtMajorCode_GotFocus()
        If IsNumeric(mSearchID) Then
            Dim objMajorAc As New clsAccounts
            objMajorAc.SetMajorAccountHead (mSearchID)
            If objMajorAc.MajorAccountHeadID > 0 Then
                txtMajorCode.Tag = objMajorAc.MajorAccountHeadID
                txtMajorCode.Text = objMajorAc.MajorAccountHeadCode
                txtMajorHead.Text = objMajorAc.MajorAccountHead
                
                If Val(txtMajorCode.Tag) <> Val(txtMinorHead.Tag) Then
                    txtMinorCode.Text = ""
                    txtMinorCode.Tag = ""
                    txtMinorHead.Text = ""
                    txtMinorHead.Tag = ""
                    txtDetailedCode.Text = ""
                    txtDetailedCode.Tag = ""
                    txtDetailedHead.Text = ""
                    txtPrimaryAlias.Text = ""
                    txtGroup.Text = ""
                    txtPrimaryCode.Text = ""
                    txtSecondaryCode.Text = ""
                    txtSecondaryHead.Text = ""
                    txtSecondaryAlias.Text = ""
                    txtOpening.Text = ""
                    optDebit.Value = True
                    optIncome.Value = True
                    mEditFlag = False
                End If
            End If
            mSearchID = Null
        End If
    End Sub

    Private Sub txtMajorCode_LostFocus()
        Dim mDB As New clsDB
        Dim mCn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        If Len(txtMajorCode) Then
            mDB.SetConnection mCn
            Rec.CursorLocation = adUseClient
            Set Rec = GetRecordSet("spGetMajorAccountHeadDetails  Null," & Trim(txtMajorCode.Text))
            If Not (Rec.BOF And Rec.EOF) Then
                txtMajorCode.Tag = Rec!intMajorAccountHeadID
                txtMajorCode.Text = Rec!vchMajorAccountHeadCode
                txtMajorHead.Text = Rec!vchMajorAccountHead
                Select Case Rec!tinType
                    Case 1: optIncome.Value = True
                    Case 2: optExpenditure.Value = True
                    Case 3: optLiability.Value = True
                    Case 4: optAsset.Value = True
                End Select
            End If
            'Rec.Close
        End If
        
    End Sub

    Private Sub txtMinorCode_GotFocus()
        If IsNumeric(mSearchID) Then
            Dim objMinorAc As New clsAccounts
            Dim objMajorAc As New clsAccounts
            
            objMinorAc.SetMinorAccountHead (mSearchID)
             If objMinorAc.MinorAccountHeadID > 0 Then
                txtMinorCode.Tag = objMinorAc.MinorAccountHeadID
                txtMinorCode.Text = objMinorAc.MinorAccountHeadCode
                txtMinorHead.Text = objMinorAc.MinorAccountHead
                txtMajorCode.Tag = objMinorAc.MajorAccountHeadID
                 If Val(txtMajorCode.Tag) > 0 Then
                    objMajorAc.SetMajorAccountHead (Val(txtMajorCode.Tag))
                        If objMajorAc.MajorAccountHeadID > 0 Then
                        txtMajorCode.Text = objMajorAc.MajorAccountHeadCode
                        txtMajorHead.Text = objMajorAc.MajorAccountHead
                    Else
                        txtMajorCode.Text = ""
                        txtMajorHead.Text = ""
                    End If
                End If
            End If
            mSearchID = Null
        End If
    End Sub

    Private Sub txtMinorCode_LostFocus()
        Dim mDB As New clsDB
        Dim mCn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecMajorDetails As New ADODB.Recordset
        Dim mSQL As String
        
        If Len(txtMinorCode) Then
            mDB.SetConnection mCn
            Rec.CursorLocation = adUseClient
            Set Rec = GetRecordSet("spGetMinorAccountHeadDetails  Null," & Trim(txtMinorCode.Text))
            If Not (Rec.BOF And Rec.EOF) Then
                txtMinorCode.Tag = Rec!intMinorAccountHeadID
                txtMinorCode.Text = Rec!vchMinorAccountHeadCode
                txtMinorHead.Text = Rec!vchMinorAccountHead
                txtMajorCode.Tag = Rec!intMajorAccountHeadID
                Select Case Rec!tinType
                    Case 1: optIncome.Value = True
                    Case 2: optExpenditure.Value = True
                    Case 3: optLiability.Value = True
                    Case 4: optAsset.Value = True
                End Select
            End If
            If Val(txtMinorCode.Tag) <> Val(txtMinorByDetailHide.Tag) Then
                txtDetailedCode.Text = ""
                txtDetailedCode.Tag = ""
                txtDetailedHead.Text = ""
                txtPrimaryAlias.Text = ""
                txtGroup.Text = ""
                txtPrimaryCode.Text = ""
                txtSecondaryCode.Text = ""
                txtSecondaryHead.Text = ""
                txtSecondaryAlias.Text = ""
                txtOpening.Text = ""
                optDebit.Value = True
                optIncome.Value = True
                mEditFlag = False
            End If
            Rec.Close
        End If
    End Sub
 
    Private Sub txtDetailedCode_GotFocus()
           Dim objAc As New clsAccounts
           Dim mSQL As String
           If gbSearchStr <> "" Then
               Dim mStr As String
               txtDetailedCode.Text = Trim(Token(gbSearchStr, " "))
               txtDetailedHead.Text = Trim(gbSearchStr)
               txtDetailedCode.Tag = gbSearchID
               objAc.SetAccounts (gbSearchID)
               txtDetailedHead.Tag = objAc.AccountType
               txtPrimaryCode.Text = objAc.AccountCode
               txtMinorByDetailHide.Tag = objAc.MinorAccountHeadID
               mGroupID = objAc.GroupID
               txtGroup.Tag = mGroupID
               gbSearchStr = ""
               gbSearchID = -1
               Call DisplayHeadsByDetailedHead
           End If
           txtDetailedHead.SelStart = 0
           txtDetailedHead.SelLength = Len(txtDetailedHead)
    End Sub
    Private Sub DisplayHeadsByDetailedHead()
            Dim objAcc As New clsAccounts
            Dim mSQL As String
            If Val(txtDetailedCode.Tag) > 0 Then
                mEditFlag = True
                objAcc.SetAccounts (Val(txtDetailedCode.Tag))
                    txtMajorCode.Tag = objAcc.MajorAccountHeadID
                    txtMinorCode.Tag = objAcc.MinorAccountHeadID
                    txtPrimaryAlias.Text = objAcc.Alias
                    txtGroup.Text = objAcc.Group
                    txtMajorHead.Tag = objAcc.MinorAccountHeadID
                    txtGroup.Tag = objAcc.GroupID
                   txtOpening.Text = Abs(objAcc.OpeningBalance)
                    If objAcc.DebitOrCredit Then
                        optDebit = True
                    Else
                        optCredit = True
                    End If
                objAcc.SetMajorAccountHead (Val(txtMajorCode.Tag))
                    txtMajorHead.Text = objAcc.MajorAccountHead
                    txtMajorCode.Text = objAcc.MajorAccountHeadCode
                objAcc.SetMinorAccountHead (Val(txtMinorCode.Tag))
                    txtMinorHead.Text = objAcc.MinorAccountHead
                    txtMinorCode.Text = objAcc.MinorAccountHeadCode
                    txtMinorHead.Tag = objAcc.MajorAccountHeadID
            End If
    End Sub
    
    Private Sub txtSecondaryCode_LostFocus()
        Dim objAcc As New clsAccounts
        Dim mSecondaryCode As String
        Dim mPrimary As String
        Dim mSecondary As String
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCon As New ADODB.Connection
        
        mSecondaryCode = Trim(txtPrimaryCode.Text) + Trim(txtSecondaryCode.Text)
        objAcc.SetAccountCode (mSecondaryCode)
        If objAcc.SecondaryAccountHead = True And objAcc.AccountHeadID > 0 Then
            mEditFlag = True
            txtPrimaryCode.Tag = objAcc.primaryID
            objAcc.SetAccountID (Val(txtPrimaryCode.Tag))
                mPrimary = objAcc.AccountCode
            objAcc.SetAccountCode (mSecondaryCode)
                mSecondary = objAcc.AccountCode
                txtSecondaryCode.Text = Right(mSecondary, Len(mSecondary) - Len(mPrimary))
                txtSecondaryHead.Text = objAcc.AccountHead
                txtSecondaryAlias.Text = objAcc.Alias
                txtGroup.Text = objAcc.Group
                txtOpening.Text = objAcc.OpeningBalance
        Else
'                mEditFlag = False
'                txtSecondaryCode.Text = Trim(txtSecondaryCode.Text)
'                txtSecondaryHead.Text = ""
'                txtSecondaryAlias.Text = ""
'                txtGroup.Text = ""
'                txtOpening.Text = ""
       End If
    End Sub
