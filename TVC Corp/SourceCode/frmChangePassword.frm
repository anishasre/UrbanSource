VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2730
   ClientLeft      =   4950
   ClientTop       =   4125
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   945
      Left            =   120
      TabIndex        =   7
      Top             =   990
      Width           =   4425
      Begin VB.TextBox txtNewPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   180
         Width           =   2025
      End
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   540
         Width           =   2025
      End
      Begin VB.Label lblNewPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter  New Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   1965
      End
      Begin VB.Label lblConfirmPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm New Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   540
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   705
      Left            =   120
      TabIndex        =   5
      Top             =   210
      Width           =   4425
      Begin VB.TextBox txtOldPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label lblOldPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter  Old Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   330
         TabIndex        =   6
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC9C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2100
      Width           =   1725
   End
   Begin VB.CommandButton cmdChangePassword 
      BackColor       =   &H00FFC9C0&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2100
      Width           =   1725
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormInitialize()
    txtConfirmPassword.Text = ""
    txtNewPassword.Text = ""
    txtOldPassword.Text = ""
End Sub
Private Sub ChangePWD()
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim aryIn As Variant
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
        objDb.CreateNewConnection mCnn, enuSourceString.DBMaster
    mSQL = "Select dbo.fnDecrypt(chbPassword) from GM_User Where numUserID = " & gbUserID
    Rec.Open mSQL, mCnn
    If Rec(0) <> txtOldPassword.Text Then
        MsgBox "Give Your Existing Password Correctly", vbInformation
        txtOldPassword.SetFocus
        Exit Sub
    End If
    '-------------------------------------------'
    '                   Validations             '
    '-------------------------------------------'
    If txtOldPassword.Text = "" Then
        MsgBox "Pleae give your Existing Password", vbInformation
        txtOldPassword.SetFocus
        Exit Sub
    End If
    
    If txtNewPassword.Text = "" Then
        MsgBox "Please give the New Password", vbInformation
        txtNewPassword.SetFocus
        Exit Sub
    End If
    
    If txtConfirmPassword.Text = "" Then
        MsgBox "Please Confirm your Password", vbInformation
        txtConfirmPassword.SetFocus
        Exit Sub
    End If
    '-------------------------------------------'
    If txtNewPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "Difference in new Password and Confirm Password", vbInformation
        txtNewPassword.SetFocus
        Exit Sub
    End If
    aryIn = Array(gbUserID, txtConfirmPassword.Text)
    objDb.ExecuteSP "SpGM_User_U", aryIn, , , mCnn, adCmdStoredProc
    'MsgBox "Your Password is Successfylly Changed!!!", vbInformation
    MsgBox "Your Password Changed Successfully !!!", vbInformation
    Call FormInitialize
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePassword_Click()
    Call ChangePWD
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
End Sub

