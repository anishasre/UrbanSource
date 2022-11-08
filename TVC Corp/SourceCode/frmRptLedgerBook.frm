VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptLedgerBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ledger Book"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   3270
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3120
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
      ScaleWidth      =   6165
      TabIndex        =   7
      Top             =   0
      Width           =   6165
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Top             =   1710
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   22675459
      CurrentDate     =   39343
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   1110
      Width           =   375
   End
   Begin VB.TextBox txtAccountHead 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   1110
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   315
      Left            =   3300
      TabIndex        =   6
      Top             =   1710
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   22675459
      CurrentDate     =   39343
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To"
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
      Left            =   3060
      TabIndex        =   5
      Top             =   1770
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From"
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
      Left            =   1020
      TabIndex        =   3
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account Head"
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
      Left            =   330
      TabIndex        =   0
      Top             =   1140
      Width           =   1140
   End
End
Attribute VB_Name = "frmRptLedgerBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub FormInitialize()
    txtAccountHead.Text = ""
    txtAccountHead.Tag = ""
    dtpFromDate.Value = gbStartingDate
    dtpToDate.Value = gbEndingDate
End Sub

Private Sub cmdSearch_Click()
    frmSearchAccountHeads.SQLString = ""
    frmSearchAccountHeads.Show vbModal
    txtAccountHead.SetFocus
End Sub

Private Sub cmdShow_Click()
    Dim frmNewRpt As New frmRptViewer
    Dim arrInput As Variant
    
    arrInput = Array(Val(txtAccountHead.Tag), dtpFromDate.Value, dtpToDate.Value)
    'Load frmNewRpt
    frmNewRpt.WindowState = vbMaximized
    frmNewRpt.rptFileName = App.Path & "\Reports\rptLedgerBook.rpt"
    frmNewRpt.InputParameters = arrInput
    
    Call frmNewRpt.ShowReport
    frmNewRpt.Show
End Sub
Private Sub dtpFromDate_Click()
    dtpFromDate.Value = dtpFromDate.Value
End Sub

Private Sub dtpToDate_Click()
    dtpToDate.Value = dtpToDate.Value
End Sub

Private Sub Form_Load()
    Call FormInitialize
End Sub

Private Sub txtAccountHead_GotFocus()
    If gbSearchStr <> "" Then
        txtAccountHead.Text = Token(gbSearchStr, " ")
        txtAccountHead.Text = Trim(gbSearchStr)
        txtAccountHead.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
    txtAccountHead.SelStart = 0
    txtAccountHead.SelLength = Len(txtAccountHead.Text)
End Sub
