VERSION 5.00
Begin VB.Form frmSearchExternalData 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRows 
      Height          =   4545
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   6045
   End
End
Attribute VB_Name = "frmSearchExternalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Option Explicit
    
    Private mvarTransactionTypeID As Long
    
    Private Sub FillList(mTransactionTypeID As Long)
        
        Dim objDB As New clsDB
        
        Dim mExtCnn As New ADODB.Connection
        Dim mConStr As String
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        Dim objTranType As New clsTransactionType
        
        
        objTranType.SetTransactionType (mTransactionTypeID)
        Select Case objTranType.ExternalApplicationID
            Case Is = AppID.Payroll
                
                'Select Case mTransactionTypeID
                '    Case Is < 10
                
                arrInput = Array(mTransactionTypeID)
                mConStr = objDB.GetConnectionString(2)
                objDB.SetExtDBConnection mExtCnn, mConStr
                Set Rec = objDB.ExecuteSP("spGetExtTransaction", arrInput, , False, mExtCnn)
                If Not (Rec.BOF And Rec.EOF) Then
                    While Not Rec.EOF
                        lstRows.AddItem Rec!vchExtTransactionCode
                        Rec.MoveNext
                    Wend
                End If
                Rec.Close
                
        End Select

    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then
            gbSearchStr = lstRows.Text
            Unload Me
        ElseIf KeyCode = vbKeyF4 Then
            Unload Me
        End If
    End Sub
    
    Private Sub Form_Load()
        Call FillList(mvarTransactionTypeID)
    End Sub
    
    Public Property Let TransactionTypeID(mData As Long)
        mvarTransactionTypeID = mData
    End Property
