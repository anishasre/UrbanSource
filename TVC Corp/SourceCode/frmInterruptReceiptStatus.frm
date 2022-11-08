VERSION 5.00
Begin VB.Form frmInterruptReceiptStatus 
   BackColor       =   &H00A65D3E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   675
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4770
      TabIndex        =   0
      Top             =   390
      Width           =   6990
   End
End
Attribute VB_Name = "frmInterruptReceiptStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
   
    '*********************************************************************************************'
    ' Form to alert the Approving Officer that Interrupt Receipt Request is pending for approval  '
    '*********************************************************************************************'
   
    Private Sub Form_Load()
        'Call Timer1_Timer
    End Sub

    Private Sub lblMessage_Click()
        lblMessage.Visible = False
        Unload frmInterruptedReceiptRequest
        frmInterruptedReceiptRequest.Show vbModal
    End Sub

    Private Sub Timer1_Timer()
        Dim mConTimer       As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim mSql            As String
        Dim RecTimer        As New ADODB.Recordset
        Dim mCountTimer     As Variant
        
        objDb.CreateNewConnection mConTimer, enuSourceString.Saankhya

        mCountTimer = ""
        mSql = "Select Count(*) As Count From faInterruptedRequests"
        mSql = mSql + " Where tnyStatus = 1"
        mSql = mSql + " And intTypeID = 1"
        RecTimer.Open mSql, mConTimer
        If Not (RecTimer.EOF And RecTimer.BOF) Then
            mCountTimer = RecTimer!count
        End If
        RecTimer.Close
        mConTimer.Close
        If mCountTimer <> 0 Then
            Timer2.Enabled = True
            If mCountTimer = 1 Then
                lblMessage.Caption = mCountTimer & " Interrupted Receipt request waiting for Approval"
            Else
                lblMessage.Caption = mCountTimer & " Interrupted Receipt requests waiting for Approval"
            End If
        Else
            lblMessage.Visible = False
            Timer2.Enabled = False
        End If
    End Sub

    Private Sub Timer2_Timer()
        If frmInterruptedReceiptRequest.Visible = False Then
            lblMessage.Visible = True
            Me.Top = frmMenu.Height - 2100
            Me.Width = frmMenu.Width - 200
            lblMessage.Left = Me.Width - 6990
            lblMessage.Top = 0
        End If
    End Sub
