VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchFunctionary 
   Caption         =   "Search Functionary"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3450
      Top             =   5430
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.ListBox lstFunctionary 
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
      Height          =   4590
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   6015
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   315
      Left            =   5730
      TabIndex        =   2
      Top             =   5130
      Width           =   375
   End
   Begin VB.TextBox txtSearchKey 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   5130
      Width           =   4665
   End
   Begin VB.ListBox lstHead 
      Height          =   255
      Left            =   6090
      TabIndex        =   0
      Top             =   5340
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Functionary"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   5190
      Width           =   825
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "  Functionary"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "frmSearchFunctionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmdSearch_Click()
        On Error GoTo Err:
            Dim mQuery1 As String
            Dim mQuery2 As String

            If txtSearchKey.Text = "" Then
                mQuery1 = "Select (vchFunctionaryCode + '  ' + vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries"
                mQuery2 = "Select (vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries"

                PopulateList lstFunctionary, mQuery1, , True, True, True
                PopulateList lstHead, mQuery2, , True, True, True
                
                Exit Sub
            End If

            If IsNumeric(txtSearchKey.Text) Then
                mQuery1 = "Select (vchFunctionaryCode + '  ' + vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries Where vchFunctionaryCode Like '" & val(txtSearchKey.Text) & "%'"
                mQuery2 = "Select (vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries Where vchFunctionaryCode Like '" & val(txtSearchKey.Text) & "%'"
                
            Else
                mQuery1 = "Select (vchFunctionaryCode + '  ' + vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries Where vchFunctionary Like '%" & Trim(txtSearchKey.Text) & "%'"
                mQuery2 = "Select (vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries Where vchFunctionary Like '%" & Trim(txtSearchKey.Text) & "%'"
            End If
            
            PopulateList lstFunctionary, mQuery1, , True, True, True
            PopulateList lstHead, mQuery2, , True, True, True
            
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Load()
        On Error GoTo Err:
            Dim mQuery1 As String
            Dim mQuery2 As String

            WindowsXPC1.InitIDESubClassing
            
            mQuery1 = "Select (vchFunctionaryCode + '  ' + vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries"
            mQuery2 = "Select (vchFunctionary) as FunctionaryHead, intFunctionaryID  From faFunctionaries"

            PopulateList lstFunctionary, mQuery1, , True, True, True
            PopulateList lstHead, mQuery2, , True, True, True
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub txtSearchKey_Change()
            Dim mIndex As Long
            Dim mStr As String
            mStr = txtSearchKey.Text
            If IsNumeric(mStr) Then
                mIndex = SendMyMessage(lstFunctionary.hwnd, LB_FINDSTRING, -1, ByVal mStr)
            Else
                mIndex = SendMyMessage(lstHead.hwnd, LB_FINDSTRING, -1, ByVal mStr)
            End If
            If mIndex > -1 Then
                lstFunctionary.ListIndex = mIndex
            End If
    End Sub
    Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
        '38 Up Arrow
        If (KeyCode = 38 Or KeyCode = 40) Then
            If KeyCode = 38 And lstFunctionary.ListIndex > 0 Then
                lstFunctionary.ListIndex = lstFunctionary.ListIndex - 1
            End If
            '40 = Down Arrow
            If KeyCode = 40 And lstFunctionary.ListIndex < (lstFunctionary.ListCount - 1) Then
                lstFunctionary.ListIndex = lstFunctionary.ListIndex + 1
                'Debug.Print lstAccountHeads.ListCount - 1, lstAccountHeads.ListIndex
            End If
        End If
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gbSearchStr = ""
        gbSearchID = -1
        If KeyCode = vbKeyEscape Then
            Unload Me
        ElseIf KeyCode = 13 Then
            If lstFunctionary.ListIndex > -1 Then
                gbSearchStr = lstFunctionary.Text
                gbSearchID = lstFunctionary.ItemData(lstFunctionary.ListIndex)
                Unload Me
            End If
        End If
    End Sub

    Private Sub lstFunctionary_DblClick()
        Call Form_KeyDown(13, 0)
    End Sub


