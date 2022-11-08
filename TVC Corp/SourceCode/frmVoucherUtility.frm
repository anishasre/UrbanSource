VERSION 5.00
Begin VB.Form frmVoucherUtility 
   BackColor       =   &H00EBF7F7&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher Utility"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete "
      Height          =   510
      Left            =   2685
      TabIndex        =   7
      Top             =   3675
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F3F3FA&
      Height          =   1665
      Left            =   780
      TabIndex        =   4
      Top             =   720
      Width           =   4950
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   1560
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   90
         Width           =   4920
      End
   End
   Begin VB.TextBox txtVoucherNo 
      Height          =   285
      Left            =   1905
      TabIndex        =   2
      Top             =   390
      Width           =   1830
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EBF7F7&
      Caption         =   "Change Date"
      Height          =   855
      Left            =   780
      TabIndex        =   0
      Top             =   2565
      Width           =   4935
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   405
         Left            =   3465
         TabIndex        =   6
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox txtNewDate 
         Height          =   285
         Left            =   1860
         TabIndex        =   1
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Date"
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
         TabIndex        =   8
         Top             =   345
         Width           =   810
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No:"
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
      Left            =   855
      TabIndex        =   3
      Top             =   405
      Width           =   1005
   End
End
Attribute VB_Name = "frmVoucherUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
    If Val(txtVoucherNo.Tag) <= 0 Then
        MsgBox " Please enter the Voucher No!", vbInformation
        txtVoucherNo.SetFocus
        Exit Sub
    End If
    If Not IsDate(txtNewDate) Then
        MsgBox " Please enter the new date to Change!", vbInformation
        txtNewDate.SetFocus
        Exit Sub
    End If
    
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    
    objDb.SetConnection mCnn
    mSQL = "Update faVouchers Set dtDate = '" & txtNewDate.Text & "' Where faVouchers.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    
    mSQL = "Update faTransactions Set dtTransactionDate = '" & txtNewDate.Text & "' Where faTransactions.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    
End Sub

Private Sub cmdDelete_Click()
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    
    objDb.SetConnection mCnn
    
    mSQL = "Delete From faVoucherChild Where faVoucherChild.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    
    mSQL = "Delete From faVoucherAddress Where faVoucherAddress.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    
    mSQL = "Delete From faVouchers Where faVouchers.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    
    mSQL = "Delete faTransactionChild From faTransactionChild Inner Join "
    mSQL = mSQL + " faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID  Where faTransactions.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    
    mSQL = "Delete From faTransactions Where faTransactions.intVoucherID = " & Val(txtVoucherNo.Tag)
    mCnn.Execute mSQL
    MsgBox "Successfully Deleted!", vbInformation
    
End Sub

Private Sub txtNewDate_LostFocus()
    txtNewDate.Text = Trim(txtNewDate)
    If Len(txtNewDate) Then
        txtNewDate.Text = CheckDateInMMM(txtNewDate)
    Else
        txtNewDate.Text = ""
    End If
End Sub

Private Sub txtVoucherNo_LostFocus()
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    
    txtVoucherNo.Text = Trim(txtVoucherNo.Text)
    If txtVoucherNo.Text = "" Then
        Exit Sub
    End If
    
    mSQL = "Select * From faVouchers Where intVoucherNo = '" & txtVoucherNo.Text & "'"
    objDb.SetConnection mCnn
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
    mSQL = ""
    If Not (Rec.BOF And Rec.EOF) Then
        txtVoucherNo.Tag = Rec!intVoucherID
        mSQL = mSQL + vbCrLf + "             Date :" & DdMmmYy(Rec!dtDate)
        mSQL = mSQL + vbCrLf + "       Voucher No :" & Rec!intVoucherNo
        mSQL = mSQL + vbCrLf + " Transaction Type :"
        mSQL = mSQL + vbCrLf + "     Total Amount :" & Format(Rec!fltAmount, "0.00")
        mSQL = mSQL + vbCrLf + "             Name :"
        mSQL = mSQL + vbCrLf + "          Counter :" & Rec!intCounterID
        txtDetails.Text = mSQL
    Else
        txtVoucherNo.Tag = ""
        txtVoucherNo.Text = ""
        txtDetails.Text = ""
    End If
    Rec.Close
End Sub



