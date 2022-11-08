VERSION 5.00
Begin VB.Form frmYearEndProcess 
   Caption         =   "Year End Process"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmYearEndProcedure.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   13035
      TabIndex        =   5
      Top             =   1080
      Width           =   13095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00404000&
      ForeColor       =   &H80000004&
      Height          =   1095
      Left            =   -960
      Picture         =   "frmYearEndProcedure.frx":3879
      ScaleHeight     =   1095
      ScaleWidth      =   15135
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GO==>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblDepreciationOfAssets 
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciation Of Assets"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         MouseIcon       =   "frmYearEndProcedure.frx":70F2
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblCapitalWorkToAssets 
         BackStyle       =   0  'Transparent
         Caption         =   "Captial Works To Assets"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         MouseIcon       =   "frmYearEndProcedure.frx":73FC
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblClosingStockEstimate 
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Stocks Estimate"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         MouseIcon       =   "frmYearEndProcedure.frx":7706
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year End Process"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmYearEndProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


    Private Sub cmdGo_Click()
        frmYearEndCaptialtoAssets.Top = Me.Top + 1450
        frmYearEndCaptialtoAssets.Left = Me.Left + 50
        frmYearEndCaptialtoAssets.Show vbModal, Me
    End Sub

    Private Sub Form_Load()
        lblClosingStockEstimate.FontBold = False
        lblCapitalWorkToAssets.FontBold = False
        lblDepreciationOfAssets.FontBold = False
    End Sub
''''    Private Sub lblCapitalWorkToAssets_Click()
''''        frmYearEndCaptialtoAssets.Top = Me.Top
''''        frmYearEndCaptialtoAssets.Left = Me.Left
''''        frmYearEndCaptialtoAssets.Show vbModeless
''''    End Sub
    Private Sub lblClosingStockEstimate_Click()
        frmYearEndCaptialtoAssets.Hide
    End Sub
    Private Sub lblClosingStockEstimate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblClosingStockEstimate.FontBold = True
        lblClosingStockEstimate.FontUnderline = True
        lblCapitalWorkToAssets.FontBold = False
        lblCapitalWorkToAssets.FontUnderline = False
        lblDepreciationOfAssets.FontBold = False
        lblDepreciationOfAssets.FontUnderline = False
    End Sub
    Private Sub lblCapitalWorkToAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblCapitalWorkToAssets.FontBold = True
        lblCapitalWorkToAssets.FontUnderline = True
        lblClosingStockEstimate.FontBold = False
        lblClosingStockEstimate.FontUnderline = False
        lblDepreciationOfAssets.FontBold = False
        lblDepreciationOfAssets.FontUnderline = False
    End Sub
    Private Sub lblDepreciationOfAssets_Click()
        frmYearEndCaptialtoAssets.Hide
    End Sub
    Private Sub lblDepreciationOfAssets_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblDepreciationOfAssets.FontBold = True
        lblDepreciationOfAssets.FontUnderline = True
        lblCapitalWorkToAssets.FontBold = False
        lblCapitalWorkToAssets.FontUnderline = False
        lblClosingStockEstimate.FontBold = False
        lblClosingStockEstimate.FontUnderline = False
    End Sub
    Private Sub Picture1_Click()
        lblClosingStockEstimate.FontBold = False
        lblCapitalWorkToAssets.FontBold = False
        lblDepreciationOfAssets.FontBold = False
    End Sub
