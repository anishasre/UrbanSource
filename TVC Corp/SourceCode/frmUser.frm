VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   3570
      TabIndex        =   19
      Top             =   4530
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2310
      TabIndex        =   18
      Top             =   4530
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
      ScaleWidth      =   6855
      TabIndex        =   20
      Top             =   0
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Height          =   3705
      Left            =   0
      TabIndex        =   21
      Top             =   690
      Width           =   6855
      Begin VB.ListBox lstUsers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   90
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   315
         Left            =   4710
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtConfirmPassWord 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2250
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Frame Frame3 
         Caption         =   "Privilege"
         Height          =   1695
         Left            =   3480
         TabIndex        =   12
         Top             =   1860
         Width           =   2415
         Begin VB.CheckBox chkPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Print"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   1380
            Width           =   1125
         End
         Begin VB.CheckBox chkView 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "View"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   1140
            Width           =   1125
         End
         Begin VB.CheckBox chkDelete 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Delete"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   720
            TabIndex        =   15
            Top             =   870
            Width           =   1245
         End
         Begin VB.CheckBox chkEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Edit"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   630
            Width           =   1125
         End
         Begin VB.CheckBox chkAddNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Add New"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   720
            TabIndex        =   13
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "User Type"
         Height          =   1695
         Left            =   930
         TabIndex        =   8
         Top             =   1860
         Width           =   2415
         Begin VB.OptionButton optOperator 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Operator"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   690
            TabIndex        =   11
            Top             =   930
            Width           =   1605
         End
         Begin VB.OptionButton optAccountsOfficer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Accounts Office"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   690
            TabIndex        =   10
            Top             =   690
            Width           =   1605
         End
         Begin VB.OptionButton optApprover 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Approver"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   690
            TabIndex        =   9
            Top             =   450
            Width           =   1605
         End
      End
      Begin VB.TextBox txtPassWord 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2250
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2250
         MaxLength       =   50
         TabIndex        =   3
         Top             =   660
         Width           =   2415
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2250
         MaxLength       =   100
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   735
         TabIndex        =   6
         Top             =   1290
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1335
         TabIndex        =   4
         Top             =   1000
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Login ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1515
         TabIndex        =   2
         Top             =   710
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1290
         TabIndex        =   0
         Top             =   420
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mEditFlag As Boolean
    Dim objUser As New clsUser
    
    Private Sub Display(mID As Long)
    
        objUser.SetUser (mID)
        If objUser.UserID > 0 Then
            txtUser.Text = objUser.UserName
            txtLogin.Text = objUser.LoginName
            txtLogin.Tag = objUser.UserID
            txtPassWord.Text = objUser.PassWord
            txtConfirmPassWord.Text = objUser.PassWord
            Select Case objUser.UserTypeID
                Case Is = 2: optApprover.Value = True
                Case Is = 3: optAccountsOfficer.Value = True
                Case Is = 4: optOperator.Value = True
            End Select
            chkAddNew.Value = IIf(objUser.AddFlag, 1, 0)
            chkEdit.Value = IIf(objUser.EditFlag, 1, 0)
            chkDelete.Value = IIf(objUser.DeleteFlag, 1, 0)
            chkView.Value = IIf(objUser.ViewFlag, 1, 0)
            chkPrint.Visible = IIf(objUser.PrintFlag, 1, 0)
        End If
    End Sub

    Private Sub SetPrivilege()
        If Not mEditFlag Then
            If optApprover Then
                chkAddNew.Value = 1
                chkEdit.Value = 1
                chkDelete.Value = 1
                chkView.Value = 1
                chkPrint.Value = 1
            End If
            If optAccountsOfficer Then
                chkAddNew.Value = 0
                chkEdit.Value = 0
                chkDelete.Value = 0
                chkView.Value = 1
                chkPrint.Value = 1
            End If
            If optOperator Then
                chkAddNew.Value = 1
                chkEdit.Value = 0
                chkDelete.Value = 0
                chkView.Value = 1
                chkPrint.Value = 0
            End If
        End If
    End Sub

    Private Sub FormInitialize()
        mEditFlag = False
        txtConfirmPassWord = ""
        txtLogin = ""
        txtPassWord = ""
        txtUser = ""
        
        chkAddNew.Value = 0
        chkEdit.Value = 0
        chkDelete.Value = 0
        chkView.Value = 0
        chkPrint.Value = 0
        
        optApprover = False
        optOperator = True
        optAccountsOfficer = False
        Call PopulateList(lstUsers, "Select vchUserName, numUserID From faUser Order by vchUserName", , , , True)
    End Sub
    
    Private Sub chkAddNew_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub chkDelete_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub chkEdit_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub chkPrint_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub chkView_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub cmdCancel_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdSave_Click()
        Dim mLbID As Integer
        Dim mIntID As Integer
        Dim objUser As New clsUser
        Dim arrInput As Variant
        Dim objDB As New clsDB
        Dim mCon As New ADODB.Connection
        '---------------------------------------------'
        ' Validations                                 '
        '---------------------------------------------'
        If Trim(txtUser.Text) = "" Then
            txtUser.SetFocus
            Exit Sub
        End If
        If Trim(txtLogin.Text) = "" Then
            txtUser.SetFocus
            Exit Sub
        End If
        If txtPassWord.Text <> txtConfirmPassWord.Text Then
            MsgBox "Check your password!!", vbInformation
            txtPassWord.SetFocus
            Exit Sub
        End If
        '---------------------------------------------'
        ' Fetching Data for Updation                  '
        '---------------------------------------------'
       
        objUser.UserID = Val(txtLogin.Tag)
        objUser.UserName = Trim(txtUser.Text)
        objUser.LoginName = Trim(txtLogin.Text)
        objUser.PassWord = Trim(txtPassWord.Text)
        'If optAdmin.Value = True Then
            'objUser.UserTypeID = 1
        If optApprover.Value = True Then
            objUser.UserTypeID = 2
        ElseIf optAccountsOfficer.Value = True Then
            objUser.UserTypeID = 3
        ElseIf optOperator.Value = True Then
            objUser.UserTypeID = 4
        End If
        objUser.AddFlag = chkAddNew.Value
        objUser.EditFlag = chkEdit.Value
        objUser.DeleteFlag = chkDelete.Value
        objUser.ViewFlag = chkView.Value
        objUser.PrintFlag = chkPrint.Value
        objUser.CreateNewUser
        Call FormInitialize
    End Sub
    
    Private Sub cmdSearch_Click()
        lstUsers.Height = 3540
        lstUsers.Visible = True
        lstUsers.SetFocus
    End Sub
    
    Private Sub Form_Activate()
        Me.Top = (frmMenu.Height - Me.Height) / 2 - 650
        Me.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
    
    Private Sub Form_Load()
        Call FormInitialize
    End Sub
    
    Private Sub lstUsers_DblClick()
    Dim mID As Long
    If lstUsers.ListIndex > -1 Then
        mID = lstUsers.ItemData(lstUsers.ListIndex)
        If mID > 0 Then
            Call Display(mID)
            lstUsers.Visible = False
            txtUser.SetFocus
        End If
    End If
    End Sub
    
    Private Sub lstUsers_LostFocus()
        lstUsers.Visible = False
    End Sub
    
    Private Sub optAccountsOfficer_Click()
        Call SetPrivilege
    End Sub
    
    Private Sub optAccountsOfficer_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub optApprover_Click()
        Call SetPrivilege
    End Sub
    
    Private Sub optApprover_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub optOperator_Click()
        Call SetPrivilege
    End Sub
    
    Private Sub optOperator_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtConfirmPassWord_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtConfirmPassWord_LostFocus()
        If txtPassWord.Text <> txtConfirmPassWord.Text Then
            MsgBox "Check your password!!", vbInformation
            txtPassWord.SetFocus
            Exit Sub
        End If
    End Sub
    
    Private Sub txtLogin_GotFocus()
        txtLogin.SelStart = 0
        txtLogin.SelLength = Len(txtLogin.Text)
    End Sub
    
    Private Sub txtLogin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtPassWord_GotFocus()
        txtPassWord.SelStart = 0
        txtPassWord.SelLength = Len(txtPassWord)
    End Sub
    
    Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtPassWord_LostFocus()
        txtConfirmPassWord.SetFocus
    End Sub
    
    Private Sub txtUser_GotFocus()
        txtUser.SelStart = 0
        txtUser.SelLength = Len(txtUser)
    End Sub
    
    Private Sub txtUser_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub DispalyUser(mLoginName As String)
        txtLogin.Text = Trim(txtLogin.Text)
        objUser.SetUserByLogin (txtLogin.Text)
        If objUser.LoginName <> "" Then
            txtUser.Text = objUser.UserName
            txtLogin.Tag = objUser.UserID
            txtLogin.Text = objUser.LoginName
            txtPassWord.Text = objUser.PassWord
            txtConfirmPassWord.Text = objUser.PassWord
            Select Case objUser.UserTypeID
                Case Is = 2: optApprover.Value = True
                Case Is = 3: optAccountsOfficer.Value = True
                Case Is = 4: optOperator.Value = True
            End Select
            chkAddNew.Value = IIf(objUser.AddFlag, 1, 0)
            chkEdit.Value = IIf(objUser.EditFlag, 1, 0)
            chkDelete.Value = IIf(objUser.DeleteFlag, 1, 0)
            chkView.Value = IIf(objUser.ViewFlag, 1, 0)
            chkPrint.Visible = IIf(objUser.PrintFlag, 1, 0)
        End If
        mEditFlag = True
    End Sub
