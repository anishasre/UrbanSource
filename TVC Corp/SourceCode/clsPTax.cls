VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    Private mPTaxArrearHeadCode As String
    Private mPTaxCurrentHeadCode As String
    Private mPTaxAdvanceCollected As String
    Private mPTaxTransactionTypeID As Long
    Private Function CalculateFine(mFromDate As Date, mToDate As Date, mPTax As Double, mPTaxRate As Single) As Double
        Dim mNoOfMonths As Integer
        Dim mFineAmt As Single
        mNoOfMonths = DateDiff("m", mFromDate, mToDate)
        mFineAmt = mPTax * mPTaxRate / 100
        CalculateFine = mFineAmt
    End Function
    Public Function CalculatePTaxFine(numBuildingID As Double, numDemandID As Double) As Double
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim RecIDemand As New ADODB.Recordset
        Dim RecAdv As New ADODB.Recordset
        Dim mSQL As String
        
        Dim mAdvAmt As Double
        Dim mFineAmt As Double
        Dim mTotalFine As Double
        Dim mPTAmt As Double
        Dim mPTRate As Single
        Dim mFromDate As Date
        Dim mToDate As Date
        Dim mNote As String
        
        mAdvAmt = 0
        mFineAmt = 0
        mTotalFine = 0
        mPTAmt = 0
        mPTRate = 1
        
        objDB.SetConnection mCnn
        'mSQL = "Select numDemandID, dtOnDate, fltAmount From faIDemandChild Where tnyStatus = 0 And vchAccountHeadCode = '" & mPTaxArrearHeadCode & "' And numSubLedgerID = " & numBuildingID
        
        mSQL = ""
        mSQL = mSQL + " Select faIDemandChild.numDemandID, faIDemandChild.dtOnDate, faIDemandChild.fltAmount, numSubLedgerID"
        mSQL = mSQL + " From faIDemandChild Inner Join"
        mSQL = mSQL + " faIDemandTbl On faIDemandTbl.numDemandID = faIDemandChild.numDemandID"
        mSQL = mSQL + " Where faIDemandTbl.tnyStatus = 0 And faIDemandTbl.intTransactionTypeID = " & mPTaxTransactionTypeID
        mSQL = mSQL + " And faIDemandChild.vchAccountHeadCode = '" & mPTaxArrearHeadCode & "'"
        mSQL = mSQL + " And faIDemandTbl.numSubLedgerID = " & numBuildingID
        mSQL = mSQL + " And faIDemandTbl.numDemandID <= " & numDemandID
        
        RecIDemand.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        If Not (RecIDemand.EOF And RecIDemand.BOF) Then
            mSQL = ""
            mSQL = mSQL + " Select faIDemandChild.numDemandID, faIDemandChild.dtOnDate, faIDemandChild.fltAmount"
            mSQL = mSQL + " From faIDemandChild Inner Join"
            mSQL = mSQL + " faIDemandTbl On faIDemandTbl.numDemandID = faIDemandChild.numDemandID "
            mSQL = mSQL + " Where faIDemandTbl.tnyStatus = 0 And faIDemandTbl.intTransactionTypeID = " & mPTaxTransactionTypeID
            mSQL = mSQL + " And faIDemandChild.vchAccountHeadCode = '" & mPTaxAdvanceCollected & "' And faIDemandTbl.numSubLedgerID = " & RecIDemand!numSubLedgerID
            RecAdv.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        Else
            CalculatePTaxFine = 0
            Return
        End If
        While Not RecIDemand.EOF
            mPTAmt = RecIDemand!fltAmount
            mFromDate = RecIDemand!dtOnDate
            '->
            mNote = mNote + DdMmmYy(mFromDate) + "  PTax : " + Format(mPTAmt, "0.00") + vbCrLf
            While Not RecAdv.EOF
                If mAdvAmt <= 0 Then
                    mAdvAmt = RecAdv!fltAmount
                    mToDate = RecAdv!dtOnDate
                    '->
                    mNote = mNote + DdMmmYy(mToDate) + "   Adv : " + Format(mPTAmt, "0.00") + vbCrLf
                    GoTo CalculatFine:
                Else
CalculatFine:
                    mFineAmt = CalculateFine(mFromDate, mToDate, mPTAmt, mPTRate)
                    '->
                    mNote = mNote + Str(mFineAmt) & DdMmmYy(mFromDate) & "  " & DdMmmYy(mToDate) & Str(mPTAmt) & Str(mPTRate)
                    mTotalFine = mTotalFine + mFineAmt
                    If mAdvAmt >= mFineAmt Then
                        mAdvAmt = mAdvAmt - mFineAmt
                        mFineAmt = 0
                    Else
                        mFineAmt = mFineAmt - mAdvAmt
                        mAdvAmt = 0
                    End If
                    If mAdvAmt >= mPTAmt Then
                        mAdvAmt = mAdvAmt - mPTAmt
                        mPTAmt = 0
                    Else
                        mPTAmt = mPTAmt - mAdvAmt
                        mAdvAmt = 0
                    End If
                    If mAdvAmt > 0 Then
                        GoTo ReadNextDemand:
                    End If
                    If mPTAmt > 0 Then
                        mFromDate = mToDate
                    End If
                    RecAdv.MoveNext
                End If
            Wend
            If mPTAmt > 0 Then
                mToDate = gbTransactionDate
                mFineAmt = CalculateFine(mFromDate, mToDate, mPTAmt, mPTRate)
                mTotalFine = mTotalFine + mFineAmt
            End If
            
ReadNextDemand:
            RecIDemand.MoveNext
        Wend
        RecIDemand.Close
        Set RecIDemand = Nothing
        CalculatePTaxFine = mTotalFine
    End Function
    Private Sub Class_Initialize()
    
        Dim objTranType As New clsTransactionType
        mPTaxTransactionTypeID = objTranType.GetTransactionTypeID("Property Tax")
        mPTaxArrearHeadCode = "431100200"
        mPTaxCurrentHeadCode = "431100100"
        mPTaxAdvanceCollected = "350410101"
        Set objTranType = Nothing
    End Sub
