VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMaster 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1980
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabmaster 
      Height          =   3675
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   6482
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Funds"
      TabPicture(0)   =   "frmMaster.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "txtFunds"
      Tab(0).Control(3)=   "cmbFunds"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Functionaries"
      TabPicture(1)   =   "frmMaster.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblFunctionary"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "txtFunctionary"
      Tab(1).Control(3)=   "cboFunctionary"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Fields"
      TabPicture(2)   =   "frmMaster.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtField"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cboField"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Functions"
      TabPicture(3)   =   "frmMaster.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "cboFunctions"
      Tab(3).Control(3)=   "txtFunctions"
      Tab(3).ControlCount=   4
      Begin VB.TextBox txtFunctions 
         Height          =   315
         Left            =   -72615
         TabIndex        =   16
         Top             =   1245
         Width           =   1515
      End
      Begin VB.ComboBox cboFunctions 
         Height          =   315
         Left            =   -72615
         TabIndex        =   15
         Top             =   1665
         Width           =   2835
      End
      Begin VB.ComboBox cboField 
         Height          =   315
         Left            =   2220
         TabIndex        =   12
         Top             =   2100
         Width           =   2835
      End
      Begin VB.TextBox txtField 
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   1680
         Width           =   1515
      End
      Begin VB.ComboBox cboFunctionary 
         Height          =   315
         ItemData        =   "frmMaster.frx":0070
         Left            =   -73020
         List            =   "frmMaster.frx":0072
         TabIndex        =   2
         Top             =   2280
         Width           =   2835
      End
      Begin VB.TextBox txtFunctionary 
         Height          =   315
         Left            =   -73020
         TabIndex        =   7
         Top             =   1860
         Width           =   1515
      End
      Begin VB.ComboBox cmbFunds 
         Height          =   315
         ItemData        =   "frmMaster.frx":0074
         Left            =   -72420
         List            =   "frmMaster.frx":0076
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   2835
      End
      Begin VB.TextBox txtFunds 
         Height          =   315
         Left            =   -72420
         TabIndex        =   3
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Function"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73815
         TabIndex        =   18
         Top             =   1305
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73830
         TabIndex        =   17
         Top             =   1755
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   11
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Field Code"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74220
         TabIndex        =   8
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label lblFunctionary 
         Caption         =   "Functionary Code"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74220
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73740
         TabIndex        =   4
         Top             =   2340
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Fund Code"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73740
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mCon As ADODB.Connection
    Dim mCom As ADODB.Command
    Dim RecFund As New ADODB.Recordset
    Dim RecFunctionary As New ADODB.Recordset
    Dim RecFunction As New ADODB.Recordset
    Dim RecField As New ADODB.Recordset
    Dim i As String
    Private marrFundCode() As String
    Private marrFunctionaryCode() As String
    Private marrFunctionCode() As String
    Private marrFieldCode() As String
    
Private Sub cboField_Click()
    txtField.Text = marrFieldCode(cboField.ListIndex, 0)
End Sub

Private Sub cboFunctionary_Click()
   txtFunctionary.Text = marrFunctionaryCode(cboFunctionary.ListIndex, 0)
End Sub

Private Sub cboFunctions_Click()
    txtFunctions.Text = marrFunctionCode(cboFunctions.ListIndex, 0)
End Sub

Private Sub cmbFunds_Click()
  txtFunds.Text = marrFundCode(cmbFunds.ListIndex, 0)
End Sub

Private Sub Form_Load()
    Call SSTabmaster_Click(0)
End Sub

Private Sub SSTabmaster_Click(PreviousTab As Integer)
    Dim cnt As Integer
    Dim objDb As New clsDB
    objDb.SetConnection mCon
    If SSTabmaster.Tab = 0 Then
        'This code segment is for Function tab in SStab
         RecFund.Open "Select * from faFunds order by faFunds.vchFund", mCon, adOpenStatic, adLockOptimistic
         cnt = 0
         txtFunds.Text = ""
         cmbFunds.Clear
         While RecFund.EOF <> True
             ReDim Preserve marrFundCode(2, 0)
             cmbFunds.AddItem RecFund!vchFund
             cmbFunds.ItemData(cmbFunds.NewIndex) = RecFund!intFundID
             marrFundCode(cnt, 0) = RecFund!vchFundCode
             cnt = cnt + 1
             RecFund.MoveNext
         Wend
         RecFund.Close
    ElseIf SSTabmaster.Tab = 1 Then
         RecFunctionary.Open "Select * from faFunctionaries order by faFunctionaries.vchFunctionary", mCon, adOpenStatic, adLockOptimistic
         cnt = 0
         txtFunctionary.Text = ""
         cboFunctionary.Clear
         While RecFunctionary.EOF <> True
             ReDim Preserve marrFunctionaryCode(30, 0)
             cboFunctionary.AddItem RecFunctionary!vchFunctionary
             cboFunctionary.ItemData(cboFunctionary.NewIndex) = RecFunctionary!intFunctionaryID
             marrFunctionaryCode(cnt, 0) = RecFunctionary!vchFunctionaryCode
             cnt = cnt + 1
             RecFunctionary.MoveNext
         Wend
         RecFunctionary.Close
    ElseIf SSTabmaster.Tab = 3 Then
         RecFunction.Open "Select * from faFunctions order by faFunctions.vchFunction", mCon, adOpenStatic, adLockOptimistic
    
         cnt = 0
         txtFunctions.Text = ""
         cboFunctions.Clear
         While RecFunction.EOF <> True
         ReDim Preserve marrFunctionCode(152, 0)
     
         cboFunctions.AddItem RecFunction!vchFunction
         cboFunctions.ItemData(cboFunctions.NewIndex) = RecFunction!intFunctionID
         marrFunctionCode(cnt, 0) = RecFunction!vchFunctionCode
    
         cnt = cnt + 1
         RecFunction.MoveNext
         Wend
         RecFunction.Close
         
         
    ElseIf SSTabmaster.Tab = 2 Then
         RecField.Open "Select * from faFields order by faFields.vchField", mCon, adOpenStatic, adLockOptimistic
    
         cnt = 0
         txtField.Text = ""
         cboField.Clear
         While RecField.EOF <> True
         ReDim Preserve marrFieldCode(19, 0)
     
         cboField.AddItem RecField!vchField
         cboField.ItemData(cboField.NewIndex) = RecField!intFieldID
         marrFieldCode(cnt, 0) = RecField!vchFieldCode
    
         cnt = cnt + 1
         RecField.MoveNext
         Wend
         RecField.Close
         
    
    End If
End Sub

