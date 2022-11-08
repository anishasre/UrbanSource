VERSION 5.00
Begin VB.Form frmDeleteJournal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Transaction Entries"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete "
      Height          =   420
      Left            =   1350
      TabIndex        =   2
      Top             =   1725
      Width           =   1965
   End
   Begin VB.TextBox txtJournalNo 
      Height          =   285
      Left            =   1260
      TabIndex        =   1
      Top             =   465
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Journal No"
      Height          =   240
      Left            =   285
      TabIndex        =   0
      Top             =   480
      Width           =   840
   End
End
Attribute VB_Name = "frmDeleteJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    If Val(txtJournalNo.Text) > 0 Then
        If MsgBox("Do you want to delete the Journal Entry?!", vbYesNo + vbDefaultButton2) = vbYes Then
            Dim mCn As New ADODB.Connection
            Dim objDB As New clsDB
            
            objDB.SetConnection mCn
            mCn.Execute "Delete From faTransactionChild Where intTransactionID  = " & Val(txtJournalNo)
            mCn.Execute "Delete From faTransactions Where intTransactionID  = " & Val(txtJournalNo)
            MsgBox "Deleted!", vbInformation
            
        End If
    End If
    
End Sub

Private Sub txtJournalNo_LostFocus()
    If Val(txtJournalNo) > 0 Then
        txtJournalNo.Text = Val(txtJournalNo)
    Else
        txtJournalNo.Text = ""
    End If
End Sub
