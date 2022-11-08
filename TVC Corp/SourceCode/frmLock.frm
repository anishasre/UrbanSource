VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S a a n k h y a ( Locked )"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLock.frx":1CCA
   ScaleHeight     =   3345
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picKey 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   345
      Picture         =   "frmLock.frx":A242
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   5715
      Top             =   2100
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2925
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0080C0FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2235
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2925
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1875
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1875
      TabIndex        =   1
      Top             =   2175
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   6
      Top             =   2535
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   870
      TabIndex        =   5
      Top             =   2175
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   660
      TabIndex        =   0
      Top             =   885
      Width           =   4335
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    '*********************************************************************************************'
    '               Form to lock the application from unathorized access                          '
    '*********************************************************************************************'
    Private Sub cmdCancel_Click()
        lblMessage.Caption = "This application has been Locked " & vbCrLf & vbCrLf & " Only " & gbUserName & " or an Administrator can Unlock this application" & vbCrLf & vbCrLf & "Press Ctrl-Shift-Delete to Unlock this application"
        lblUserName.Visible = False
        lblPassword.Visible = False
        txtUserName.Visible = False
        txtPassword.Visible = False
        cmdLogin.Visible = False
        cmdCancel.Visible = False
        picKey.Visible = False
        txtPassword.Text = ""
        Me.Width = 5910
        Me.Height = 2955
    End Sub

    Private Sub cmdLogin_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim mUserName   As String
        Dim mLoginName  As String
        Dim mPassword   As String
        Dim mUserID     As Variant
        Dim mUserTypeID As Variant
        Dim mGbPassword As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.DBMaster
        
        If txtUserName.Text = "" Then
            MsgBox "Please enter the Login Name", vbInformation
            txtUserName.SetFocus
            Exit Sub
        End If
        If txtPassword.Text = "" Then
            MsgBox "Please enter the Password", vbInformation
            txtPassword.SetFocus
            Exit Sub
        End If
        
        mSql = "Select dbo.fnDecrypt(chbPassword) As Password From GM_User"
        mSql = mSql + " Where numUserID=" & gbUserID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mGbPassword = Rec!PassWord
        End If
        Rec.Close
        
        mSql = "Select dbo.fnDecrypt(chbPassword) As Password,numUserID,chvUserID From GM_User"
        mSql = mSql + " Where chvUserID='" & CStr(txtUserName.Text) & "'"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mUserID = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
            mLoginName = IIf(IsNull(Rec!chvUserID), "", Rec!chvUserID)
            mPassword = Rec!PassWord
        End If
        Rec.Close
        mCnn.Close
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If mUserID <> "" Then
            mSql = "Select intUserTypeID,vchUserName From faUser"
            mSql = mSql + " Where numUserID=" & mUserID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mUserTypeID = IIf(IsNull(Rec!intUserTypeID), "", Rec!intUserTypeID)
                mUserName = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
            End If
            Rec.Close
            If (mUserID = gbUserID And txtPassword.Text = mGbPassword) Then
                Unload Me
                frmMenu.Visible = True
                frmMenu.WindowState = 2
            ElseIf mUserTypeID = 1 Then
                If (mPassword = Trim(txtPassword.Text)) Then
                    Unload Me
                    frmMenu.Visible = True
                    frmMenu.WindowState = 2
                Else
                    MsgBox "Login failed", vbInformation
                    Exit Sub
                End If
            Else
                MsgBox "Login failed", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "Login failed", vbInformation
            Exit Sub
        End If
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'        Dim mCnn        As New ADODB.Connection
'        Dim Rec         As New ADODB.Recordset
'        Dim mSql        As String
'        Dim objDb       As New clsDB
'        Dim mUserName   As String
        
        If Shift = 3 And KeyCode = 46 Then
            lblMessage.Caption = "This application has been Locked " & vbCrLf & vbCrLf & " Only " & gbUserName & " or an Administrator can Unlock this application"
            Me.Width = 5910
            Me.Height = 3825
            lblUserName.Visible = True
            lblPassword.Visible = True
            txtUserName.Visible = True
            txtPassword.Visible = True
            cmdLogin.Visible = True
            cmdCancel.Visible = True
            picKey.Visible = True
            txtUserName.Text = GetSetting("Saankhya", "Lock", "UserName", CStr(txtUserName.Text))
'            objDb.CreateNewConnection mCnn, enuSourceString.DBMaster
'
'            mSql = "Select chvUserID From GM_User Where numUserID=" & gbUserID
'            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mUserName = IIf(IsNull(Rec!chvUserID), "", Rec!chvUserID)
'            End If
'            Rec.Close
'            mCnn.Close
'            txtUserName.Text = mUserName
            txtUserName.SetFocus
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub Form_Load()
        Me.Width = 5910
        Me.Height = 2955
        frmMenu.WindowState = 1
        frmMenu.Visible = False
        lblMessage.Caption = "This application has been Locked " & vbCrLf & vbCrLf & " Only " & gbUserName & " or an Administrator can Unlock this application" & vbCrLf & vbCrLf & "Press Ctrl-Shift-Delete to Unlock this application"
        txtUserName.Text = GetSetting("Saankhya", "Lock", "UserName", "")
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        SaveSetting "Saankhya", "Lock", "UserName", CStr(txtUserName.Text)
    End Sub

    Private Sub Timer1_Timer()
        Call cmdCancel_Click
        Timer1.Enabled = False
    End Sub
    
    Private Sub txtPassword_KeyPress(KeyAscii As Integer)
        Timer1.Enabled = False
        Timer1.Enabled = True
        If KeyAscii = 13 Then
            cmdLogin_Click
'            Call PressTabKey
        End If
        If KeyAscii = 27 Then
            Call cmdCancel_Click
        End If
    End Sub

    Private Sub txtUserName_KeyPress(KeyAscii As Integer)
        Timer1.Enabled = False
        Timer1.Enabled = True
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
        If KeyAscii = 27 Then
            Call cmdCancel_Click
        End If
        
    End Sub
