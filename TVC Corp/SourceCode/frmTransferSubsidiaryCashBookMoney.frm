VERSION 5.00
Begin VB.Form frmTransferSubsidiaryCashBookMoney 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer Money to Subsidiary Cash Book"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "frmTransferSubsidiaryCashBookMoney.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearchVouchers 
      Caption         =   "..."
      Height          =   285
      Left            =   7290
      TabIndex        =   40
      Top             =   4155
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame fmeSeatUser 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1065
      Left            =   4770
      TabIndex        =   5
      Top             =   900
      Width           =   3825
      Begin VB.CommandButton cmdUser 
         Caption         =   "..."
         Height          =   330
         Left            =   3360
         TabIndex        =   37
         Top             =   735
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   690
         TabIndex        =   36
         Top             =   750
         Width           =   2670
      End
      Begin VB.CommandButton cmdSeat 
         Caption         =   "..."
         Height          =   330
         Left            =   3360
         TabIndex        =   35
         Top             =   405
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtSeat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   690
         TabIndex        =   34
         Top             =   420
         Width           =   2670
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User:"
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
         Left            =   270
         TabIndex        =   39
         Top             =   795
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seat:"
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
         Left            =   285
         TabIndex        =   38
         Top             =   450
         Width           =   390
      End
   End
   Begin VB.Frame fmeOthers 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3945
      Left            =   90
      TabIndex        =   6
      Top             =   780
      Width           =   8550
      Begin VB.TextBox txtCashBookID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   31
         Top             =   525
         Width           =   1380
      End
      Begin VB.CheckBox chkExpenditure 
         Caption         =   "Is Expenditure Recorded"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5055
         TabIndex        =   21
         Top             =   2925
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   20
         Text            =   "30-Oct-2009"
         Top             =   855
         Width           =   1380
      End
      Begin VB.TextBox txtSubCashBook 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   19
         Top             =   1290
         Width           =   6135
      End
      Begin VB.CommandButton cmdSubCash 
         Caption         =   "..."
         Height          =   315
         Left            =   8025
         TabIndex        =   18
         Top             =   1290
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2940
         Width           =   2670
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   16
         Top             =   2610
         Width           =   6135
      End
      Begin VB.CommandButton cmdFunction 
         Caption         =   "..."
         Height          =   300
         Left            =   4575
         TabIndex        =   15
         Top             =   2295
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtFunction 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   14
         Top             =   2280
         Width           =   2670
      End
      Begin VB.CommandButton cmdFunctionary 
         Caption         =   "..."
         Height          =   300
         Left            =   4575
         TabIndex        =   13
         Top             =   1965
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtFunctionary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   12
         Top             =   1950
         Width           =   2670
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3270
         Width           =   2670
      End
      Begin VB.TextBox txtAccCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1890
         TabIndex        =   10
         Top             =   1620
         Width           =   1830
      End
      Begin VB.TextBox txtAccountHead 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3735
         TabIndex        =   9
         Top             =   1620
         Width           =   4290
      End
      Begin VB.CommandButton cmdAccHead 
         Caption         =   "..."
         Height          =   315
         Left            =   8040
         TabIndex        =   8
         Top             =   1620
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtVoucherNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5955
         MaxLength       =   15
         TabIndex        =   7
         Top             =   3345
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Subsidiary Cash Book ID"
         Enabled         =   0   'False
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
         Left            =   75
         TabIndex        =   32
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
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
         Left            =   1455
         TabIndex        =   30
         Top             =   900
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Subsidiary Cash Book:"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount:"
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
         Left            =   1260
         TabIndex        =   28
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         Left            =   1185
         TabIndex        =   27
         Top             =   2685
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Function:"
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
         Left            =   1155
         TabIndex        =   26
         Top             =   2355
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Functionary:"
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
         Left            =   900
         TabIndex        =   25
         Top             =   2025
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Reference:"
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
         Left            =   1050
         TabIndex        =   24
         Top             =   3360
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Account Head:"
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
         Left            =   780
         TabIndex        =   23
         Top             =   1665
         Width           =   1095
      End
      Begin VB.Label lblVoucherNo 
         AutoSize        =   -1  'True
         Caption         =   "Voucher No"
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
         Left            =   5055
         TabIndex        =   22
         Top             =   3375
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9180
      TabIndex        =   0
      Top             =   5025
      Width           =   9240
      Begin VB.CommandButton cmdApprove 
         Caption         =   "Approve"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1275
         TabIndex        =   33
         Top             =   495
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2689
         TabIndex        =   3
         Top             =   495
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3855
         TabIndex        =   2
         Top             =   60
         Width           =   1110
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Transfer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4035
         TabIndex        =   1
         Top             =   510
         Visible         =   0   'False
         Width           =   1110
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subsidiary Cash Book Transfer"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   270
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   -60
      Picture         =   "frmTransferSubsidiaryCashBookMoney.frx":1CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11610
   End
End
Attribute VB_Name = "frmTransferSubsidiaryCashBookMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private intSubsidiaryCashBookID As Variant
    '*********************************************************************************************'
    '              Form for an Acc.Clerk to self approve the Subsidiary Cash Book                 '
    '*********************************************************************************************'
    
     Private Function GetUserName(mUserID As Variant) As String
        Dim mCnnUserName    As New ADODB.Connection
        Dim RecUserName     As New ADODB.Recordset
        Dim objUserName     As New clsDB
        Dim mSQLUserName    As String
        
        '*********************************************************************************************'
        '                     Function to get the User Name from DB_Masters                           '
        '*********************************************************************************************'
        On Error GoTo err
        objUserName.CreateNewConnection mCnnUserName, enuSourceString.DBMaster
        
        mSQLUserName = "Select * From GM_User"
        mSQLUserName = mSQLUserName + " Where numUserID = " & mUserID
        RecUserName.Open mSQLUserName, mCnnUserName
        If Not (RecUserName.EOF And RecUserName.BOF) Then
            GetUserName = IIf(IsNull(RecUserName!vchEmpName), "", RecUserName!vchEmpName)
        End If
        RecUserName.Close
        Exit Function
err:
        MsgBox err.Description
    End Function
    
    Private Function GetSeatName(mSeatID As Variant)
        Dim mCnnSeatName    As New ADODB.Connection
        Dim RecSeatName     As New ADODB.Recordset
        Dim objSeatName     As New clsDB
        Dim mSQLSeatName    As String
        
        '*********************************************************************************************'
        '                         Function to get the Seat Name from DB_Masters                       '
        '*********************************************************************************************'
        On Error GoTo err
        objSeatName.CreateNewConnection mCnnSeatName, enuSourceString.DBMaster
        
        mSQLSeatName = "Select * From GL_Seats"
        mSQLSeatName = mSQLSeatName + " Where numSeatID = " & mSeatID
        RecSeatName.Open mSQLSeatName, mCnnSeatName
        If Not (RecSeatName.EOF And RecSeatName.BOF) Then
            GetSeatName = IIf(IsNull(RecSeatName!chvSeatTitle), "", RecSeatName!chvSeatTitle)
        End If
        RecSeatName.Close
        Exit Function
err:
        MsgBox err.Description
    End Function
    
    Private Sub chkExpenditure_Click()
        If chkExpenditure.value = 1 Then
            txtVoucherNo.Enabled = True
        Else
            txtVoucherNo.Enabled = False
        End If
    End Sub

    Private Sub cmdAccHead_Click()
        Dim mSql        As String
        Dim mCount      As Integer
        Dim mAccCode    As Variant
        
        On Error GoTo err
        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 "
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
        mAccCode = Split(gbSearchStr, "  ")
        If gbSearchID <> -1 Then
            txtAccountHead.Text = mAccCode(1)
            txtAccCode.Text = mAccCode(0)
            txtAccountHead.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdApprove_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim mStatus As Integer
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        Dim mID     As Variant
        
        '*********************************************************************************************'
        '                     Procedure to self approve the Subsidiary Cash Book                      '
        '*********************************************************************************************'
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mID = ""
        'If gbUserTypeID = 3 Then
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            mSql = "Select dtDate,intTypeID,numApprovedUserID From faSubsidiaryCashBook"
            mSql = mSql + " Where intID = " & val(txtCashBookID.Text)
            mSql = mSql + " And intTransferID = " & val(txtCashBookID.Tag)
            Rec.Open mSql, mCnn
            If Not IsNull(Rec!numApprovedUserID) Then
                If Not (IsNull(Rec!dtDate) And IsNull(Rec!intTypeID)) Then
                    If Rec!dtDate <> Date Or Rec!intTypeID <> 50 Then
                        MsgBox "Can't update the Status", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            Rec.Close
            mStatus = 1
            mCnn.Execute "Update faSubsidiaryCashBook Set tnyStatus = " & mStatus & " , numApprovedUserID = " & gbUserID & " , dtApprovalDate = '" & CheckDateInMMM(Date) & "' Where intID = " & val(txtCashBookID.Text)
        ElseIf gbSeatGroupID = gbSeatGroupChiefCashier Then
            mSql = "Select dtDate,intTypeID,numApprovedUserID From faSubsidiaryCashBook"
            mSql = mSql + " Where intID = " & val(txtCashBookID.Text)
            mSql = mSql + " And intTransferID = " & val(txtCashBookID.Tag)
            Rec.Open mSql, mCnn
            If Not IsNull(Rec!numApprovedUserID) Then
                If Not (IsNull(Rec!dtDate) And IsNull(Rec!intTypeID)) Then
                    If Rec!dtDate <> Date Or Rec!intTypeID <> 50 Then
                        MsgBox "Can't update the Status", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            Rec.Close
            mStatus = 1
            mCnn.Execute "Update faSubsidiaryCashBook Set tnyStatus = " & mStatus & " , numApprovedUserID = " & gbUserID & " , dtApprovalDate = '" & CheckDateInMMM(Date) & "' Where intID = " & val(txtCashBookID.Text)
            
        End If
        MsgBox "Successfully Saved", vbInformation
        cmdApprove.Enabled = False
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdClose_Click()
         Unload Me
    End Sub

    Private Sub cmdFunction_Click()
        On Error GoTo err:
        frmSearchFunction.Show vbModal
        If Not gbSearchStr = "" Then
            txtFunction.Text = gbSearchStr
            txtFunction.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdFunctionary_Click()
        On Error GoTo err:
            frmSearchFunctionary.Show vbModal
            If Not gbSearchStr = "" Then
                txtFunctionary.Text = gbSearchStr
                txtFunctionary.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    

    Private Sub cmdSearchVouchers_Click()
        Dim mCnn  As New ADODB.Connection
        Dim Rec   As New ADODB.Recordset
        Dim objDb As New clsDB
        Dim mSql  As String
        
        On Error GoTo err
        frmSearchVouchers.PreviousYearMode = 0
        frmSearchVouchers.CheckMode = 20
        frmSearchVouchers.chkReceipt.Enabled = False
        'frmSearchVouchers.chkPayment.Enabled = True
        frmSearchVouchers.chkContra.Enabled = False
        frmSearchVouchers.chkJournal.Enabled = False
        frmSearchVouchers.chkInterrupted.Visible = False
        'frmSearchVouchers.chkInterrupted.value = 1
        frmSearchVouchers.Show vbModal
        txtVoucherNo.Text = gbSearchCode
        txtVoucherNo.Tag = gbSearchID
        gbSearchCode = ""
        gbSearchID = -1
        
        Exit Sub
err:
      MsgBox err.Description
    End Sub

    Private Sub cmdSubCash_Click()
        On Error GoTo err
        frmSearchSubsidiaryAccountHeads.SubLedgerType = 12
        'frmSearchSubsidiaryAccountHeads.cmbSubLegerType.Text = "Subsidiary Cash Book"
        frmSearchSubsidiaryAccountHeads.Show vbModal
        If gbSearchID = -1 Then
            MsgBox "Please Select Subsidiary Account"
        Else
            txtSubCashBook.Text = gbSearchStr
            txtSubCashBook.Tag = gbSearchID
        End If
        frmSearchSubsidiaryAccountHeads.SubLedgerType = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
        txtSeat.Text = gbSearchStr
        txtSeat.Tag = gbSearchID
    End Sub

    Private Sub cmdTransfer_Click()
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mCnn            As New ADODB.Connection
        Dim mSql            As String
        Dim mID             As Integer
        Dim mTransferID     As Integer 'TransferID
        Dim mSubsidiaryAccountHeadID As Integer
        Dim arrIn           As Variant
        Dim arrOut          As Variant
        Dim mExpenditure    As Integer
        
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If txtCashBookID.Text = "" Then
            mID = -1
        Else
            mID = val(txtCashBookID.Text)
        End If
        If txtCashBookID.Tag = "" Then
            mTransferID = -1
        Else
            mTransferID = txtCashBookID.Tag
        End If
        '--------Validations-------------------------------------------------
        If txtSubCashBook.Text = "" Then
            MsgBox "Please Select SubSidiary Account", vbApplicationModal
            Exit Sub
        Else
            mSubsidiaryAccountHeadID = Trim(txtSubCashBook.Tag)
        End If
        
        If txtDate.Text = "" Then
            MsgBox "Please Enter Date", vbApplicationModal
            Exit Sub
        End If
        If txtSeat.Text = "" Then
            MsgBox "Please Select the Seat of User", vbApplicationModal
            Exit Sub
        End If
        If txtUser.Text = "" Then
            MsgBox "Please Select user", vbApplicationModal
            Exit Sub
        End If
        If txtAmount.Text = "" Then
            MsgBox "Please Enter Amount", vbApplicationModal
            Exit Sub
        End If
        If txtAccountHead.Text = "" Then
            MsgBox "Please Select AccountHead", vbApplicationModal
            Exit Sub
        End If
        If txtFunctionary.Text = "" Then
            MsgBox "Please Select Functionary"
            Exit Sub
        End If
        If txtFunction.Text = "" Then
            MsgBox "Please Select Function"
            Exit Sub
        End If
        '---------------------------------------------------------------
        'Checking weather the User already exists or not
'
'        mSql = "Select Count(*) Count From faSubsidiaryCashBook Where  intTypeID=50 And tnyStatus<>2 And numSeatID=" & Val(txtSeat.Tag) & " And numUserID=" & Val(txtUser.Tag)
'        Rec.Open mSql, mCnn
'            If Rec!Count <> 0 Then
'                MsgBox (txtUser.Text & "  Already Have a Subsidiary Account Head")
'                Exit Sub
'            End If
        '---------------------------------------------------------------
        '---------------------------------------------------------------
        If chkExpenditure.value = 1 Then
            mExpenditure = 1
        Else
            mExpenditure = 0
        End If
'        If mExpenditure = 1 Then
'            If Trim(txtVoucherNo.Text) = "" Then
'                MsgBox "Please enter the Voucher No", vbInformation
'                txtVoucherNo.SetFocus
'                Exit Sub
'            End If
'        End If
        If Trim(txtVoucherNo.Text) = "" Then
            MsgBox "Please enter the Voucher No", vbInformation
            txtVoucherNo.SetFocus
            Exit Sub
        End If
        
        arrIn = Array(mID, _
                        mTransferID, _
                        mSubsidiaryAccountHeadID, _
                         50, _
                        Format(CDate(txtDate.Text), "dd/mmm/yyyy"), _
                        val(txtUser.Tag), _
                        val(txtSeat.Tag), _
                        val(txtAccountHead.Tag), _
                        val(txtFunctionary.Tag), _
                        val(txtFunction.Tag), _
                        val(txtAmount.Text), _
                        Null, _
                        Null, _
                        txtReference.Text, _
                        txtRemarks.Text, 0, mExpenditure, txtVoucherNo.Tag)
        objDb.ExecuteSP "spSaveSubsidiaryCashBook", arrIn, arrOut, , mCnn, adCmdStoredProc
        MsgBox "Successfully Transferred", vbApplicationModal
        cmdTransfer.Enabled = False
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub cmdNew_Click()
        Call FormInitialize
        cmdTransfer.Enabled = True
    End Sub
    Private Sub cmdUser_Click()
        Dim mSql    As String
        If txtSeat.Text <> "" Then
            mSql = "Select faUserSeatAssign.numuserID,vchuserName From faUserSeatAssign "
            mSql = mSql + " Inner Join faUser On faUser.numuserID=faUserSeatAssign.numUserID Where numSeatID=" & val(txtSeat.Tag)
            frmSearchMasters.SQLQry = mSql
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            txtUser.Text = gbSearchStr
            txtUser.Tag = gbSearchID
            txtUser.SetFocus
        Else
            MsgBox "Please Select Seat Before Selecting user", vbApplicationModal
            Exit Sub
        End If
    End Sub

    Private Sub Form_Load()
        On Error GoTo err
        Call FormInitialize
        'If gbUserTypeID = 3 Then
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            fmeOthers.Enabled = False
            fmeSeatUser.Enabled = False
            cmdApprove.Visible = True
            cmdApprove.Left = 3276
            cmdApprove.Top = 60
            cmdClose.Left = 4435
            cmdNew.Visible = False
        ElseIf gbSeatGroupID = gbSeatGroupChiefCashier Then
            fmeOthers.Enabled = False
            fmeSeatUser.Enabled = False
            cmdApprove.Visible = True
            cmdApprove.Left = 3276
            cmdApprove.Top = 60
            cmdClose.Left = 4435
            cmdNew.Visible = False
        Else
            cmdApprove.Visible = False
        End If
        If SubSidiaryCashBookID <> "" Then
            Call FillDetails
            'fmeOthers.Enabled = False
            'chkExpenditure.Value = 1
        End If
        
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDate_LostFocus()
        If Trim(txtDate) <> "" Then
            txtDate.Text = CheckDateInMMM(txtDate.Text)
        Else
            txtDate.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        txtDate.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
    End Sub

    Private Sub txtFunction_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub

    Private Sub txtSeat_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub

    Private Sub txtSubCashBook_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub

    Private Sub FillDetails()
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mCnn            As New ADODB.Connection
        Dim mSql            As String
        
        On Error GoTo err
        If objDb.SetConnection(mCnn) Then
            mSql = "Select S.*,S.intSubsidiaryAccountHeadID[SubsidiaryAccountHeadID],A.vchAccountHead AccHead,A.vchAccountHeadCode Code,Fs.vchFunctionary Functionary,F.vchFunction FunName"
            mSql = mSql + " ,SA.vchTitle as SubCashTitle, V.intVoucherNo as VrNo From faSubsidiaryCashBook S"
            mSql = mSql + " Inner Join faAccountHeads A On A.intAccountHeadID=S.intAccountHeadID"
            mSql = mSql + " Left Join faFunctionaries Fs On Fs.intFunctionaryID=S.intFunctionaryID"
            mSql = mSql + " Left Join faFunctions F On F.intFunctionID=S.intFunctionID"
            mSql = mSql + " Left Join faSubSidiaryAccountHeads SA On SA.intSubsidiaryAccountHeadID = S.intSubsidiaryAccountHeadID "
            mSql = mSql + " Left Join faVouchers V On V.intVoucherID = S.intVoucherID"
            '''    mSql = mSql + " Left Join faSeats On faSeats.numSeatID=S.numSeatID"
            '''    mSql = mSql + " Left Join faUser On faUser.numuserID=s.numUserID"
            mSql = mSql + " Where intID = " & SubSidiaryCashBookID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                txtFunctionary.Text = IIf(IsNull(Rec!Functionary), "", Rec!Functionary)
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                txtFunction.Text = IIf(IsNull(Rec!FunName), "", Rec!FunName)
                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                txtAmount.Text = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                txtReference.Text = IIf(IsNull(Rec!vchReference), "", Rec!vchReference)
                txtAccCode.Text = IIf(IsNull(Rec!Code), "", Rec!Code)
                txtAccountHead.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                txtAccountHead.Text = IIf(IsNull(Rec!AccHead), "", Rec!AccHead)
                txtSubCashBook.Text = IIf(IsNull(Rec!SubCashTitle), "", Rec!SubCashTitle)
                txtSubCashBook.Tag = IIf(IsNull(Rec!SubsidiaryAccountHeadID), "", Rec!SubsidiaryAccountHeadID)
                txtVoucherNo.Text = IIf(IsNull(Rec!VrNo), "", Rec!VrNo)
                txtVoucherNo.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                txtSeat.Tag = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
                If txtSeat.Tag <> "" Then
                    txtSeat.Text = GetSeatName(txtSeat.Tag)
                End If
                txtUser.Tag = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
                If txtUser.Tag <> "" Then
                    txtUser.Text = GetUserName(txtUser.Tag)
                End If
                txtCashBookID.Text = IIf(IsNull(Rec!intID), "", Rec!intID)
                txtCashBookID.Tag = IIf(IsNull(Rec!intTransferID), "", Rec!intTransferID)
                If Rec!tnyStatus > 0 Then
                    cmdTransfer.Enabled = False
                End If
            End If
        Else
            MsgBox "Connection to Finance does not exist, Please Contact your System Administrator", vbInformation
        End If
        SubSidiaryCashBookID = ""
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub txtUser_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
       Call KeyPress(KeyAscii)
    End Sub
    Private Sub txtAccCode_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    Private Sub txtFunctionary_KeyPress(KeyAscii As Integer)
       Call KeyPress(KeyAscii)
    End Sub
    Private Sub KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        Else
            KeyAscii = 0
        End If
    End Sub
  
'    Private Sub txtUser_LostFocus()
'        Dim objDb           As New clsDB
'        Dim Rec             As New ADODB.Recordset
'        Dim mCnn            As New ADODB.Connection
'        Dim mSql            As String
'        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        mSql = "Select Count(*) Count From faSubsidiaryCashBook Where  tnyStatus in (0,1) And numSeatID=" & Val(txtSeat.Tag) & " And numUserID=" & Val(txtUser.Tag)
'        Rec.Open mSql, mCnn
'            If Not (Rec.EOF) And Rec!Count >= 1 Then
'                MsgBox (txtUser.Text & "  Already Have a Subsidiary Account Head")
'                Exit Sub
'            End If
'    End Sub
    
    Public Property Let SubSidiaryCashBookID(mData As Variant)
        intSubsidiaryCashBookID = mData
    End Property
    
    Public Property Get SubSidiaryCashBookID() As Variant
        SubSidiaryCashBookID = intSubsidiaryCashBookID
    End Property
