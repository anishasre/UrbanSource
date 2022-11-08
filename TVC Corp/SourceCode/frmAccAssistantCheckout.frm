VERSION 5.00
Begin VB.Form frmAccAssistantCheckout 
   Caption         =   "CheckOut"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdCheckout 
         Caption         =   "Checkout"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmAccAssistantCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Created By Sajith Kumar
''Added On 5-07-12
''
    Private Sub cmdCheckout_Click()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
            
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
          mSQL = "update faEmpLog set dtCheckOut=getdate() where intID=" & gbUserID
          Rec.Open mSQL, mCnn
          Unload Me
          MsgBox "Checked Out"
        End If
    End Sub
