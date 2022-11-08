VERSION 5.00
Begin VB.Form frmSearchAccountHeads 
   BackColor       =   &H00FFFFF7&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Head Search"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "frmSearchAccountHeads.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkListAll 
      BackColor       =   &H00FFFFF7&
      Caption         =   "List All"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8295
      TabIndex        =   8
      Top             =   825
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFF7&
      Caption         =   "...."
      Height          =   315
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5700
      Width           =   375
   End
   Begin VB.ListBox lstHead 
      Height          =   255
      Left            =   8955
      TabIndex        =   7
      Top             =   5745
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSearchKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF7&
      Height          =   285
      Left            =   930
      TabIndex        =   1
      Top             =   5730
      Width           =   7125
   End
   Begin VB.ComboBox cmbMajorAccountHead 
      BackColor       =   &H00FFFFF7&
      Height          =   315
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   735
      Width           =   4080
   End
   Begin VB.ListBox lstAccountHeads 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF7&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   9165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   0
      Top             =   5790
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Press Esc to Cancel"
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
      Left            =   585
      TabIndex        =   6
      Top             =   6150
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   285
      Shape           =   2  'Oval
      Top             =   6210
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Major Account Heads"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   795
      Width           =   1815
   End
End
Attribute VB_Name = "frmSearchAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private mstrSQL As String
   '-----------------------------------'
    Private mVoucherMode As Integer
    '''             100. Receipt Cr
    '''             101. Receipt Dr
    '''             200. Payment Cr
    '''             201. Payment Dr
    '''             300. Contra  Cr
    '''             301. Contra  Dr
    '''             400. Journal Cr
    '''             401. Journal Dr
   '-----------------------------------'
    
    Private Sub FillList()
'''''        Dim mSql As String
'''''        Dim mStr As String
'''''        If Trim(txtSearchKey.Text) <> "" Then
'''''            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where vchAccountHead Like '%" & Trim(txtSearchKey) & "%' And tinHiddenFlag = 0"
'''''            PopulateList lstAccountHeads, mSql, , , , True
'''''            If InStr(1, mSql, "vchMinorAccountHeadCode") Then
'''''                mStr = TokenCrop$(mSql, "vchMinorAccountHeadCode + '  ' + ")
'''''            Else
'''''                mStr = TokenCrop$(mSql, "vchAccountHeadCode + '  ' + ")
'''''            End If
'''''            mStr = "Select (" & mSql
'''''            PopulateList lstHead, mStr, , , , True
'''''        Else
'''''            PopulateList lstAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0", , , , True
'''''            PopulateList lstHead, "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads", , , , True
'''''        End If
'''''        lstAccountHeads.SetFocus
'''''        gbSearchID = -1
'''''        gbSearchStr = ""

        Dim mSQLAcc As String
        Dim mSQLHead As String
        
        Select Case mVoucherMode
                Case 100:   ''Receipt Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Null"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Null"
                Case 101:   ''Receipt Debit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Not Null"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Not Null"
                Case 200:   ''Payment Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Not Null"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Not Null"
                Case 201:   ''Payment Debit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Null"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Null"
                Case 300, 301    ''Contra Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Not Null"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Not Null"
                Case 400, 401:  ''Journal Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Null"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0 And intGroupID is Null"
                Case Else
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where (vchAccountHead Like '%" & Trim(txtSearchKey) & "%' or vchAccountHeadCode Like '%" & Trim(txtSearchKey) & "%' ) And tinHiddenFlag = 0"
            End Select
            PopulateList lstAccountHeads, mSQLAcc, , , , True
            PopulateList lstHead, mSQLHead, , , , True
    End Sub
    Private Sub chkListAll_Click()
        Dim mSQLAcc As String
        Dim mSQLHead As String
        If chkListAll.Value = 1 Then
            Select Case mVoucherMode
                Case 100:   ''Receipt Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By vchAccountHeadCode"
                Case 101:   ''Receipt Debit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Not Null Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Not Null Order By vchAccountHeadCode"
                Case 200:   ''Payment Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Not Null Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Not Null Order By vchAccountHeadCode"
                Case 201:   ''Payment Debit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By vchAccountHeadCode"
                Case 300, 301    ''Contra Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Not Null Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Not Null Order By vchAccountHeadCode"
                Case 400, 401:  ''Journal Credit
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By vchAccountHeadCode"
                Case Else
                        mSQLAcc = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 Order By vchAccountHeadCode"
                        mSQLHead = "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0 Order By vchAccountHeadCode"
            End Select
            
            PopulateList lstAccountHeads, mSQLAcc, , , , True
            PopulateList lstHead, mSQLHead, , , , True
            
'            PopulateList lstAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0", , , , True
'            PopulateList lstHead, "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where tinHiddenFlag = 0", , , , True
        End If
    End Sub
    Private Sub cmdsearch_Click()
        Call FillList
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gbSearchStr = ""
        gbSearchID = -1
        If KeyCode = vbKeyEscape Then
            mstrSQL = ""
            Unload Me
            
        ElseIf KeyCode = 13 Then
            If lstAccountHeads.ListIndex > -1 Then
                gbSearchStr = lstAccountHeads.Text
                gbSearchID = lstAccountHeads.ItemData(lstAccountHeads.ListIndex)
                
                mstrSQL = ""
                Unload Me
                
            End If
        End If
    End Sub
    Private Sub Form_Load()
        Dim mStr As String
        PopulateList cmbMajorAccountHead, "Select vchMajorAccountHead From faMajorAccountHeads", , , True
        
        If mstrSQL <> "" Then
            PopulateList lstAccountHeads, mstrSQL, , , , True
            If InStr(1, mstrSQL, "vchMinorAccountHeadCode") Then
                mStr = TokenCrop$(mstrSQL, "vchMinorAccountHeadCode + '  ' + ")
            ElseIf InStr(1, mstrSQL, "vchMajorAccountHeadCode") Then
                mStr = TokenCrop$(mstrSQL, "vchMajorAccountHeadCode + '  ' + ")
            Else
                mStr = TokenCrop$(mstrSQL, "vchAccountHeadCode + '  ' + ")
            End If
            If mstrSQL = "" Then
                mstrSQL = mStr
            Else
                mstrSQL = "Select (" & mstrSQL
            End If
            'PopulateList lstAccountHeads, mstrSQL, , , , True
            PopulateList lstHead, mstrSQL, , , , True
        Else
            If mVoucherMode = 0 Then
                If gbLBPanchayat = 1 Then
                    PopulateList lstAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where intGroupID is Null and intMinorAccountHeadID<>220", , , , True
                    PopulateList lstHead, "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads", , , , True
                Else
                    PopulateList lstAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where intGroupID is Null and intMinorAccountHeadID<>248 ", , , , True
                    PopulateList lstHead, "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads", , , , True
               
                End If
            Else
                PopulateList lstAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads", , , , True
                PopulateList lstHead, "Select (vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads", , , , True
            End If
            ''Commented by Anisha on 18.5.13
'            PopulateList lstAccountHeads, mstrSQL, , , , True
'            PopulateList lstHead, mstrSQL, , , , True
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        mVoucherMode = 0
    End Sub

    Private Sub lstAccountHeads_DblClick()
        Call Form_KeyDown(13, 0)
    End Sub
    Public Property Let SQLString(mSql As String)
        mstrSQL = mSql
    End Property
    
    Public Property Let VoucherMode(mData As Integer)
        mVoucherMode = mData
    End Property
    
    Private Sub txtSearchKey_Change()
            Dim mIndex As Long
            Dim mStr As String
            mStr = txtSearchKey.Text
            If IsNumeric(mStr) Then
                mIndex = SendMyMessage(lstAccountHeads.hwnd, LB_FINDSTRING, -1, ByVal mStr)
            Else
                mIndex = SendMyMessage(lstHead.hwnd, LB_FINDSTRING, -1, ByVal mStr)
            End If
            If mIndex > -1 Then
                lstAccountHeads.ListIndex = mIndex
            End If
    End Sub
    Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
        '38 Up Arrow
        If (KeyCode = 38 Or KeyCode = 40) Then
            Dim objAcc As New clsAccounts
            If KeyCode = 38 And lstAccountHeads.ListIndex > 0 Then
                lstAccountHeads.ListIndex = lstAccountHeads.ListIndex - 1
            End If
            '40 = Down Arrow
            If KeyCode = 40 And lstAccountHeads.ListIndex < (lstAccountHeads.ListCount - 1) Then
                lstAccountHeads.ListIndex = lstAccountHeads.ListIndex + 1
                'Debug.Print lstAccountHeads.ListCount - 1, lstAccountHeads.ListIndex
            End If
            If lstAccountHeads.ListIndex > -1 Then
                objAcc.SetAccounts (lstAccountHeads.ItemData(lstAccountHeads.ListIndex))
                txtSearchKey.Text = objAcc.AccountCode
            End If
        End If
    End Sub
