VERSION 5.00
Begin VB.Form frmCounterReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saankhya - Counter Reports"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.CommandButton cmdNo 
         Caption         =   "No"
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "Yes"
         Height          =   315
         Left            =   2490
         TabIndex        =   2
         Top             =   870
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Are you Sure to Close the Counter !"
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Top             =   450
         Width           =   2775
      End
      Begin VB.Label lblMsg 
         Caption         =   "You must Close the Conuters Before Logout"
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   540
         TabIndex        =   1
         Top             =   150
         Width           =   3375
      End
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   3660
      TabIndex        =   3
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmCounterReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Sub cmdNo_Click()
        Unload Me
    End Sub

    Private Sub cmdYes_Click()
        Unload Me
        frmMenu.DailyReports.Enabled = True
        frmMenu.Transactions.Enabled = False
        frmCounterReport.Show
        frmCounterReport.ZOrder (0)
    End Sub
    Private Sub Form_Load()
'        If frmReceiptsCounter.Visible = True Then
'            Unload frmReceiptsCounter
'        End If
        lblDate.Caption = Date
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        frmCounterReports.Visible = False
    End Sub
