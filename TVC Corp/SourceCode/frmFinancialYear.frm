VERSION 5.00
Begin VB.Form frmFinancialYear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmFinancialYear"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5850
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Height          =   285
      Left            =   3690
      TabIndex        =   9
      Top             =   540
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txtEndingDate 
      Height          =   285
      Left            =   1695
      TabIndex        =   6
      Top             =   1215
      Width           =   1725
   End
   Begin VB.TextBox txtStartDate 
      Height          =   285
      Left            =   1695
      TabIndex        =   5
      Top             =   885
      Width           =   1710
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   1695
      TabIndex        =   4
      Top             =   540
      Width           =   1725
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2933
      TabIndex        =   3
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1703
      TabIndex        =   0
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Starting Date"
      Height          =   195
      Left            =   675
      TabIndex        =   8
      Top             =   930
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   195
      Left            =   1275
      TabIndex        =   7
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Financial Year Settings"
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
      TabIndex        =   2
      Top             =   30
      Width           =   5805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ending Date"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   1290
      Width           =   885
   End
End
Attribute VB_Name = "frmFinancialYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private Sub cmbYear_Click()
        If cmbYear.ListIndex > -1 Then
           cmbYear.Text = cmbYear.ItemData(cmbYear.ListIndex)
           txtStartDate.Text = Format("1-4-" + CStr(cmbYear.Text), "dd-mmm-yyyy")
           txtEndingDate.Text = Format("31-3-" + CStr(cmbYear.Text + 1), "dd-mmm-yyyy")
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    
    Private Sub cmdSave_Click()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim arrIn   As Variant
        Dim mYear   As Integer
        Dim mActive As Integer
        If cmbYear.ListIndex > -1 Then
            mYear = cmbYear.ItemData(cmbYear.ListIndex)
        Else
            MsgBox "Please Select Financial Year"
            Exit Sub
        End If
        If chkActive.Value Then
            mActive = 1
            mSql = "Update faFinancialYear set tinCurrentFinancialYearFlag=0"
            objDb.ExecuteSP mSql, , , , mCnn, adCmdText
        Else
            mActive = 0
        End If
        If mActive = 1 Then
            mSql = "Update faFinancialYear set tinCurrentFinancialYearFlag=0 "
        End If
            arrIn = Array(mYear, mYear, Format(txtStartDate.Text, "dd/mmm/yyyy"), Format(txtEndingDate.Text, "dd/mmm/yyyy"), Format(txtStartDate.Text, "dd/mmm/yyyy"), mActive, gbLocalBodyID)
            objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
            objDb.ExecuteSP "spSaveFinancialYear", arrIn, , , mCnn, adCmdStoredProc


    End Sub
    Private Sub CurrentFinancialYear()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec    As New ADODB.Recordset
        Dim objDb   As New clsDB
        Dim arrIn   As Variant
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select * From faFinancialYear Where tinCurrentFinancialYearFlag=1"
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        If Rec.RecordCount > 0 Then
            MsgBox "Financial Year Settings Alredy done " & vbCrLf + vbCrLf & Rec!intFinancialYear
        End If
    End Sub

    Private Sub Form_Load()
        Call FillYear
        Call CurrentFinancialYear
    End Sub
    Private Sub FillYear()
        Dim mLoop As Integer
        Dim mItem As String
        Dim mYearID As Integer
        mItem = "#0; "
        mYearID = Year(Date)
        For mLoop = mYearID + 5 To mYearID - 5 Step -1
            mItem = CStr(mLoop)
            cmbYear.AddItem (mItem)
            cmbYear.ItemData(cmbYear.NewIndex) = mItem
        Next
    End Sub

