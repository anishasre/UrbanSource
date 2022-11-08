VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateAccountHeadBalance 
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbOverall 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   1020
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   5430
      TabIndex        =   0
      Top             =   -75
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar pbCurrent 
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   330
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Label1"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   5775
   End
   Begin VB.Label lblOverall 
      Caption         =   "Label1"
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   5775
   End
End
Attribute VB_Name = "frmUpdateAccountHeadBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub RefreshOpeningBalance()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecTransactionChild As New ADODB.Recordset
    Dim mSQL As String
    Dim mOpentingAmt As Double
    Dim mCurrentBalance As Double
    
    objDB.SetConnection mCnn
    mSQL = "Select intAccountHeadID,fltOpeningBalance, fltCurrentBalance From faAccountHeads Order By intAccountHeadID"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    If Not Rec.EOF Then Rec.MoveLast: pbOverall.Max = Rec.RecordCount: pbOverall.Visible = True: Rec.MoveFirst
    While Not Rec.EOF
        mCurrentBalance = 0
        mOpentingAmt = IIf(IsNull(Rec!fltOpeningBalance), 0, Rec!fltOpeningBalance)
        mSQL = "Select faTransactionChild.intAccountHeadID, faTransactionChild.fltAmount, faTransactionChild.tinDebitOrCreditFlag, faTransactionChild.fltOpeningBalance From faTransactionChild  Inner Join faTransactions "
        mSQL = mSQL + " faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID "
        mSQL = mSQL + " Where intAccountHeadID = " & Rec!intAccountHeadID & " Order by dtTransactionDate, faTransactionChild.intTransactionID"
        RecTransactionChild.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        If Not RecTransactionChild.EOF Then RecTransactionChild.MoveLast: pbCurrent.Min = 0: pbCurrent.Max = RecTransactionChild.RecordCount: pbCurrent.Visible = True: RecTransactionChild.MoveFirst
        While Not RecTransactionChild.EOF
            RecTransactionChild!fltOpeningBalance = mOpentingAmt
            RecTransactionChild.Update
            If RecTransactionChild!tinDebitOrCreditFlag Then
                mOpentingAmt = mOpentingAmt + RecTransactionChild!fltAmount
            Else
                mOpentingAmt = mOpentingAmt - RecTransactionChild!fltAmount
            End If
            mCurrentBalance = mOpentingAmt
            If pbCurrent.Value < pbCurrent.MouseIcon Then pbCurrent.Value = pbCurrent.Value + 1
            RecTransactionChild.MoveNext
        Wend
        RecTransactionChild.Close
        Rec!fltCurrentBalance = mCurrentBalance
        Rec.Update
        pbOverall.Value = pbOverall.Value + 1
        Rec.MoveNext
    Wend
    Rec.Close
    Set mCnn = Nothing
End Sub


Private Sub cmdUpdate_Click()
    If MsgBox(" Do you want to update all records?", vbYesNo, "Saankhya") = vbYes Then
        cmdUpdate.Visible = False
        RefreshOpeningBalance
        'UpdateAllOpeningBalance
        MsgBox "All records updated", vbInformation, "Saankhya"
    End If
    Unload Me
End Sub

Public Sub UpdateAllOpeningBalance()
    cmdUpdate.Visible = False
    Dim RecAccHead As New ADODB.Recordset
    Dim mCon As New ADODB.Connection
    Dim objDB As New clsDB
    Dim mAccHeadCount As Long
    Dim mAccHeads As Variant
    Dim mLoop As Integer
    If objDB.SetConnection(mCon) Then
        RecAccHead.Open "SELECT intAccountHeadID,vchAccountHead FROM faAccountHeads", mCon
        If Not RecAccHead.EOF Then
            lblOverall.Visible = True
            pbOverall.Visible = True
            mAccHeads = RecAccHead.GetRows()
            mAccHeadCount = UBound(mAccHeads, 2) + 1
            pbOverall.Max = mAccHeadCount
            For mLoop = 0 To UBound(mAccHeads, 2)
                DoEvents
                UpdateOpeningBalance mAccHeads(0, mLoop), mAccHeads(1, mLoop)
                If pbOverall.Value < pbOverall.Max Then
                    pbOverall.Value = pbOverall.Value + 1
                    lblOverall.Caption = "Updating account heads " & CStr(CInt(pbOverall.Value / pbOverall.Max * 100)) & "% completed"
                End If
            Next mLoop
        End If
    End If
End Sub
 Private Sub UpdateOpeningBalance(ByVal mAccountHeadID As Integer, ByVal mAccHeadstr As String)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCon As New ADODB.Connection
        Dim mSQL As String
        Dim mVTransactions As Variant
        Dim mCurrentBalance As Double
        Dim mLoop As Long
        Dim mQuery As String
        Dim fltAmount As Double
        Dim mFirstTrans As Boolean
        Dim mTransCount As Long
        mFirstTrans = True
        If objDB.SetConnection(mCon) Then
            Rec.Open "Select intTransactionID,intSerialNo,fltAmount,tinDebitOrCreditFlag FROM FATRANSACTIONCHILD Where intTransactionID<> 0 and intAccountHeadID= " & mAccountHeadID, mCon
            If Not Rec.EOF Then
                 mVTransactions = Rec.GetRows
            End If
            
            If IsArray(mVTransactions) Then
                If Rec.State = 1 Then
                    Rec.Close
                End If
                lblCurrent.Visible = True
                pbCurrent.Visible = True
                mTransCount = UBound(mVTransactions, 2) + 1
                pbCurrent.Max = mTransCount
                For mLoop = 0 To UBound(mVTransactions, 2)
                    DoEvents
                    If Rec.State = 1 Then
                        Rec.Close
                    End If
                    If mFirstTrans = True Then
                        Rec.Open "Select fltOpeningBalance FROM faAccountHEads Where intAccountHeadID=" & mAccountHeadID, mCon
                        mFirstTrans = False
                    Else
                        Rec.Open "Select fltCurrentBalance FROM faAccountHEads Where intAccountHeadID=" & mAccountHeadID, mCon
                    End If
                    
                    If Not Rec.EOF Then
                        mCurrentBalance = IIf(IsNull(Rec.Fields(0)), 0#, Rec.Fields(0))
                    Else
                        mCurrentBalance = 0
                    End If
                    
                    mQuery = " Update faTransactionChild set fltOpeningBalance =" & mCurrentBalance & " where intTransactionID=" & mVTransactions(0, mLoop) & " and intSerialNo=" & mVTransactions(1, mLoop)
                    mCon.Execute mQuery
                    If mVTransactions(3, mLoop) = 1 Then
                        fltAmount = mVTransactions(2, mLoop)
                    Else
                        fltAmount = mVTransactions(2, mLoop) * (-1)
                    End If
                    mQuery = "Update faAccountHeads set fltCurrentBalance= " & mCurrentBalance + fltAmount & " Where intAccountHeadID= " & mAccountHeadID
                    mCon.Execute mQuery
                    If pbCurrent.Value < pbCurrent.Max Then
                        pbCurrent.Value = pbCurrent.Value + 1
                        lblCurrent.Caption = "Updating " & mAccHeadstr & CStr(CInt(pbCurrent.Value / pbCurrent.Max * 100)) & "% completed"
                    End If
                Next mLoop
            End If
        End If
    End Sub

Private Sub Form_Load()
   FormInitialize
End Sub
Private Sub FormInitialize()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    cmdUpdate.Left = (Me.Width - cmdUpdate.Width) / 2
    cmdUpdate.Top = (Me.Height - cmdUpdate.Height) / 2
    pbCurrent.Visible = False
    pbOverall.Visible = False
    lblCurrent.Visible = False
    lblOverall.Visible = False
End Sub

