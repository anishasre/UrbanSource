VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChequeBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
    
    Private mvarChequeBookID    As Long
    Private mvarBookNo          As Long
    Private mvarPrefix          As String
    Private mvarStartingNo      As Long
    Private mvarEndingNo        As Long
    Private mvarIssuedDate      As Date
    Private mvarCurrentBookFlag As Boolean
    Private mvarClosedFlag      As Boolean
    Private mvarBankID          As Long
    Private mvarAccountHeadID   As Long
    Private mvarAccountNo       As String
    
    Public Property Get AccountHeadID() As Long
        AccountHeadID = mvarAccountHeadID
    End Property
    
    Public Property Let AccountHeadID(ByVal vData As Long)
        mvarAccountHeadID = vData
    End Property
    
    Public Property Get BankID() As Long
        BankID = mvarBankID
    End Property
    
    Public Property Let BankID(ByVal vData As Long)
        mvarBankID = vData
    End Property
    
    Public Property Let AccountNumber(ByVal vData As String)
        mvarAccountNo = vData
    End Property
    
    Public Property Get AccountNumber() As String
        AccountNumber = mvarAccountNo
    End Property
    
    Public Property Get ChequeBookID() As Long
        ChequeBookID = mvarChequeBookID
    End Property
    
    Public Property Let ChequeBookID(ByVal vData As Long)
        mvarChequeBookID = vData
    End Property
    
    Public Property Let BookNo(ByVal vData As Long)
        mvarBookNo = vData
    End Property
    
    Public Property Get BookNo() As Long
        BookNo = mvarBookNo
    End Property

    Public Property Let Prefix(ByVal vData As String)
        mvarPrefix = vData
    End Property
    
    Public Property Get Prefix() As String
        Prefix = mvarPrefix
    End Property

    Public Property Let StartingNo(ByVal vData As Long)
        mvarStartingNo = vData
    End Property
    
    Public Property Get StartingNo() As Long
        StartingNo = mvarStartingNo
    End Property

    Public Property Let EndingNo(ByVal vData As Long)
        mvarEndingNo = vData
    End Property
    
    Public Property Get EndingNo() As Long
        EndingNo = mvarEndingNo
        End Property

    Public Property Let IssuedDate(ByVal vData As Date)
        mvarIssuedDate = vData
    End Property
    
    Public Property Get IssuedDate() As Date
        IssuedDate = mvarIssuedDate
    End Property

    Public Property Let CurrentBookFlag(ByVal vData As Boolean)
        mvarCurrentBookFlag = vData
    End Property
    
    Public Property Get CurrentBookFlag() As Boolean
        CurrentBookFlag = mvarCurrentBookFlag
    End Property

    Public Property Let ClosedFlag(ByVal vData As Boolean)
        mvarClosedFlag = vData
    End Property
    
    Public Property Get CloaseFlag() As Boolean
        ClosedFlag = mvarClosedFlag
    End Property
    Private Sub ClearVariables()
            mvarChequeBookID = -1
            mvarBookNo = -1
            mvarPrefix = ""
            mvarStartingNo = -1
            mvarEndingNo = -1
            mvarIssuedDate = Date
            mvarCurrentBookFlag = False
            mvarClosedFlag = False
    End Sub
    Public Sub SetChequeBookInfo(mBankID As Long)
            
            Dim mCn As ADODB.Connection
            Dim rec As New ADODB.Recordset
            Dim mSQL As String
            
            Call ClearVariables
            
            mSQL = "SELECT * "
            mSQL = mSQL & " FROM faChequeBook LEFT JOIN "
            mSQL = mSQL & " faAccountHeads ON faChequeBook.intAccountHeadID = faAccountHeads.intAccountHeadID LEFT JOIN "
            mSQL = mSQL & " faBanks ON faBanks.intBankID = faChequeBook.intBankID "
            mSQL = mSQL & " WHERE faChequeBook.intBankID = " & mBankID
            
            
            rec.CursorLocation = adUseClient
            Set rec = GetRecordSet(mSQL)
            If rec.RecordCount > 0 Then
                '-----------------------------------------------'
                ' To over ride Null Value Errors if any         '
                '-----------------------------------------------'
                  On Error Resume Next
                '-----------------------------------------------'
                mvarChequeBookID = rec!intChequeBookID
                mvarAccountHeadID = rec!intAccountHeadID
                mvarBankID = rec!intBankID
                mvarBookNo = rec!intBookNo
                mvarPrefix = rec!vchPrefix
                mvarStartingNo = rec!intStartingNo
                mvarEndingNo = rec!intEndingNo
                mvarIssuedDate = rec!dtissuedDate
                mvarCurrentBookFlag = rec!tinCurrentBookFlag
                mvarClosedFlag = rec!tinClosedFlag
                On Error GoTo 0
            End If
            rec.Close
    End Sub
