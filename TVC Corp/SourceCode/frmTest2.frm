VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmTest2 
   Caption         =   "frmTest2"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   ScaleHeight     =   7890
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cochinFinal 
      Caption         =   "COCHIN Final"
      Height          =   540
      Left            =   6030
      TabIndex        =   16
      Top             =   5850
      Width           =   4275
   End
   Begin VB.CommandButton cmdCochTest 
      Caption         =   "COCHIN Test"
      Height          =   540
      Left            =   6030
      TabIndex        =   15
      Top             =   5280
      Width           =   4275
   End
   Begin VB.CommandButton cmdLFA2 
      Caption         =   " [2] LFA [2]"
      Height          =   525
      Left            =   165
      TabIndex        =   14
      Top             =   6885
      Width           =   4290
   End
   Begin VB.CommandButton Cochin 
      Caption         =   "COCHIN"
      Height          =   540
      Left            =   180
      TabIndex        =   13
      Top             =   6195
      Width           =   4275
   End
   Begin VB.CommandButton cmdLFA 
      Caption         =   "LFA"
      Height          =   525
      Left            =   180
      TabIndex        =   12
      Top             =   5595
      Width           =   4260
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   480
      Left            =   8805
      TabIndex        =   11
      Top             =   120
      Width           =   2070
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   465
      Left            =   5970
      TabIndex        =   10
      Top             =   165
      Width           =   2385
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   525
      Left            =   195
      TabIndex        =   9
      Top             =   4950
      Width           =   4260
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   4515
      Left            =   6030
      TabIndex        =   8
      Top             =   750
      Width           =   6000
      lastProp        =   500
      _cx             =   10583
      _cy             =   7964
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   0   'False
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Report"
      Height          =   525
      Left            =   180
      TabIndex        =   7
      Top             =   4260
      Width           =   4260
   End
   Begin VB.CommandButton cmdBankReconciliationAnalysis 
      Caption         =   "Bank Reconciliation Analysis"
      Height          =   525
      Left            =   180
      TabIndex        =   6
      Top             =   3630
      Width           =   4260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   840
      Left            =   4800
      TabIndex        =   5
      Top             =   1005
      Width           =   765
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DSN By Registry Entry"
      Height          =   525
      Left            =   180
      TabIndex        =   4
      Top             =   2730
      Width           =   4260
   End
   Begin VB.CommandButton cmdCreateDSN 
      Caption         =   "Create DSN"
      Height          =   525
      Left            =   180
      TabIndex        =   3
      Top             =   2100
      Width           =   4260
   End
   Begin VB.CommandButton cmdFixVouchersUsingTransactionsJournal 
      Caption         =   "Fix Vouchers Using Transactions - Journal Vouchers"
      Height          =   525
      Left            =   180
      TabIndex        =   2
      Top             =   1500
      Width           =   4260
   End
   Begin VB.CommandButton cmdFixVouchersUsingTransactions 
      Caption         =   "Fix Vouchers Using Transactions - Contra Vouchers"
      Height          =   525
      Left            =   180
      TabIndex        =   1
      Top             =   870
      Width           =   4260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grid in Tree View"
      Height          =   525
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   4260
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1335
      Left            =   4710
      TabIndex        =   17
      Top             =   6420
      Width           =   7125
      _cx             =   12568
      _cy             =   2355
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTest2.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



 Option Explicit

    Private Const REG_SZ = 1    'Constant for a string variable type.
    Private Const HKEY_LOCAL_MACHINE = &H80000002

    Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
       "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
       phkResult As Long) As Long

    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
       ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
       cbData As Long) As Long

    Private Declare Function RegCloseKey Lib "advapi32.dll" _
       (ByVal hKey As Long) As Long

Private Sub cmdBankReconciliationAnalysis_Click()
    '-------------------------------------------------------------'
    ' To Analyse Bank Reconciliation By Comparing Three tables    '
    '-------------------------------------------------------------'
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim mAmt As Double
        Dim mLoopControl As Long
        Dim RecBank As New ADODB.Recordset
        Dim Rec As New ADODB.Recordset
        
        mSql = "Select * From faBankReconciliationEntries Where tnyReconciled = 2 And intBankAccountHeadID = 1506 Order By numTockenID"
        objDB.SetConnection mCnn
        RecBank.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        FileInitialize
        While Not RecBank.EOF
            If Not IsNull(RecBank!intReconciliationID) Then
                mSql = " Select SUM(fltAmount) fltAmount From faVouchers Where tnyReconciled = 2 AND intKeyID1 = 1506 AND numTockenID = " & RecBank!intReconciliationID
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                If IsNumeric(RecBank!fltDrAmount) Then
                    mAmt = RecBank!fltDrAmount
                Else
                    mAmt = RecBank!fltCrAmount
                End If
                
                If IsNull(Rec!fltAmount) Or (mAmt <> Rec!fltAmount) Then
                    Print #gbFileNO, RecBank!intReconciliationID, mAmt, Rec!fltAmount
                    mSql = " Update faVouchers Set tnyReconciled = 9 Where tnyReconciled = 2 AND intKeyID1 = 1506 AND numTockenID = " & RecBank!intReconciliationID
                    mCnn.Execute mSql
                    mSql = " Update faBankReconciliationEntries SET tnyReconciled = 9 Where intBankAccountHeadID = 1506 AND intReconciliationID = " & RecBank!intReconciliationID
                    mCnn.Execute mSql
                End If
                Rec.Close
            End If
            RecBank.MoveNext
        Wend
        RecBank.Close
        Close #gbFileNO
        ShellPad
End Sub

    Private Sub cmdCochTest_Click()
       Dim xmlHttp As Object
       Set xmlHttp = CreateObject("MSXML2.XmlHttp")
       Dim mXmlString   As Variant
       Dim Rec  As New ADODB.Recordset
        Dim REcDet As New ADODB.Recordset
        Dim params
        Dim mCnt    As Integer
        params = "lbCode=G051105&year=2014-2015”"
        '//  doorNo~doorNo2~AssessmentNo~applicationNo~applicationStatus~ownerName~zoneId~wardId~ownerPin~ownerAddress
        '// 25/1~ ~ ~ ~ ~ ~9~ ~ ~
        ' OR
        '// 25/1~NA~NA~NA~NA~NA~9~NA~NA~NA

'''''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
'''''<propertyTaxVOes>
'''''<PropertyTaxVO>
'''''<actionCode>0</actionCode>
'''''<applicantName>Rajappan</applicantName>
'''''<applicationNo>CoC/PT/9/5/202/2015</applicationNo>
'''''<currentHalfYearPaid>false</currentHalfYearPaid>
'''''<doorNo>101</doorNo>
'''''<id>202</id>
'''''<ownerAddress>Sopanam,Menaka,Kochi,KLER,KL,IND</ownerAddress>
'''''<ownerPin>685474</ownerPin>
'''''<statusText>Property Tax Assessment form payment done</statusText>
'''''<wardNo>5</wardNo>
'''''<workFlowLevel>0</workFlowLevel>
'''''<zoneNo>9</zoneNo>
'''''</PropertyTaxVO>
'''''</propertyTaxVOes>


        Set Rec = New ADODB.Recordset
    
        params = "9~NA~25/1~NA~NA~NA~NA~NA~NA~NA~NA"
        xmlHttp.Open "POST", "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService/searchAssesmentDetails?searchParam=" & params, False
        
        'xmlHttp.Open "POST", "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService/getDemandDtls/" & 233, False
        
        xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
        xmlHttp.send
        mXmlString = xmlHttp.responseText
        'mXmlString = Replace(mXmlString, "UTF-8", "UTF-16")
        
        Dim oRS As ADODB.Recordset
        Dim oNode As Object 'MSXML2.IXMLDOMNode
        Dim oSubNodes As Object 'MSXML2.IXMLDOMSelection
        Dim oDoc As Object
        'Dim oDoc As MSXML2.DOMDocument26
        Set oDoc = CreateObject("MSXML2.DOMDocument")
        'Set oNode = CreateObject("MSXML2.IXMLDOMNode")
    '    Set oSubNodes = CreateObject("MSXML2.IXMLDOMSelection")
        
        oDoc.async = False
        oDoc.validateOnParse = False
        If Not oDoc.LoadXml(mXmlString) Then
            MsgBox "Error Loading"
            Exit Sub
        Else
            MsgBox "Sucess"
        End If
    
        Set oRS = New ADODB.Recordset
        Set oRS.ActiveConnection = Nothing
        oRS.CursorLocation = adUseClient
        oRS.LockType = adLockBatchOptimistic
    
        With oRS.Fields
             .Append "buildlingID", adInteger
            .Append "assementYear", adVarChar, 20
            .Append "zoneID", adInteger
            .Append "wardNo", adInteger
            .Append "wardName", adVarChar, 20
            .Append "doorNo1", adInteger
            .Append "doorNo2", adVarChar, 10
            .Append "ownerName", adVarChar, 10
            .Append "houseBuildingName", adVarChar, 50
            .Append "street", adVarChar, 50
            .Append "localplace", adVarChar, 50
            .Append "mainplace", adVarChar, 50
            
            .Append "post", adVarChar, 10
            .Append "district", adVarChar, 50
            .Append "pin", adVarChar, 20
            .Append "phone", adVarChar, 20
            .Append "ownerMobileNo", adVarChar, 20
        End With
       
        oRS.Open
        
         For Each oNode In oDoc.selectNodes("/propertyTaxVOss/PropertyTaxVO")
         'For Each oNode In oDoc.selectNodes("/PropertyTaxVo/demandRegisters")
            oRS.ADDNEW
            oRS.Fields("buildlingID").value = oNode.selectSingleNode("buildlingID").Text
            oRS.Fields("assementYear").value = oNode.selectSingleNode("assementYear").Text
            oRS.Fields("zoneID").value = oNode.selectSingleNode("zoneID").Text
            oRS.Fields("wardNo").value = oNode.selectSingleNode("wardNo").Text
            oRS.Fields("wardName").value = oNode.selectSingleNode("wardName").Text
            oRS.Fields("doorNo1").value = oNode.selectSingleNode("doorNo1").Text
            vsGrid.TextMatrix(mCnt, 0) = oRS.Fields("doorNo1").value
            oRS.Fields("doorNo2").value = oNode.selectSingleNode("doorNo2").Text
            oRS.Fields("ownerName").value = oNode.selectSingleNode("ownerName").Text
            vsGrid.TextMatrix(mCnt, 1) = oRS.Fields("ownerName").value
            oRS.Fields("houseBuildingName").value = oNode.selectSingleNode("houseBuildingName").Text
            oRS.Fields("street").value = oNode.selectSingleNode("street").Text
            oRS.Fields("localplace").value = oNode.selectSingleNode("localplace").Text
            oRS.Fields("mainplace").value = oNode.selectSingleNode("mainplace").Text
            oRS.Fields("post").value = oNode.selectSingleNode("post").Text
            oRS.Fields("district").value = oNode.selectSingleNode("district").Text
            oRS.Fields("pin").value = oNode.selectSingleNode("pin").Text
            oRS.Fields("phone").value = oNode.selectSingleNode("phone").Text
            oRS.Fields("ownerMobileNo").value = oNode.selectSingleNode("ownerMobileNo").Text
            vsGrid.TextMatrix(mCnt, 2) = oRS.Fields("ownerMobileNo").value
            mCnt = mCnt + 1
        Next
            
'        Dim mXmlStream As New ADODB.Stream
'        mXmlStream.Open
'        mXmlStream.WriteText mXmlString
'        mXmlStream.Flush
'        mXmlStream.Position = 0
'        'Rec.Fields.Append "ID", adInteger
'        Rec.Open mXmlStream
'        mXmlStream.Close
        
       
        'MsgBox mXmlString 'xmlHttp.responseText

    End Sub
    
'''Public Sub LoadDocument()
'''    Dim xDoc As MSXML.DOMDocument
'''    Set xDoc = New MSXML.DOMDocument
'''    xDoc.validateOnParse = False
'''    If xDoc.Load("C:\My Documents\sample.xml") Then
'''       ' The document loaded successfully.
'''       ' Now do something intersting.
'''       DisplayNode xDoc.childNodes, 0
'''    Else
'''       ' The document failed to load.
'''       ' See the previous listing for error information.
'''    End If
'''End Sub

''Public Sub DisplayNode(ByRef Nodes As MSXML.IXMLDOMNodeList, _
''   ByVal Indent As Integer)
''
''   Dim xNode As MSXML.IXMLDOMNode
''   Indent = Indent + 2
''
''   For Each xNode In Nodes
''      If xNode.nodeType = NODE_TEXT Then
''         Debug.Print Space$(Indent) & xNode.parentNode.nodeName & _
''            ":" & xNode.nodeValue
''      End If
''
''      If xNode.hasChildNodes Then
''         DisplayNode xNode.childNodes, Indent
''      End If
''   Next xNode
''End Sub


'''''
'''''Public Function ConvertXMLtoRecordset(ByVal voNL As IXMLDOMNodeListByVal, sTableName As String) As ADODB.Recordset
'''''Dim oTableNode As IXMLDOMNode
'''''Dim oRecordNode As IXMLDOMNode
'''''Dim oFieldNode As IXMLDOMNode
'''''Dim oNodeList As IXMLDOMNodeList
'''''Dim oRS As ADODB.Recordset
'''''Dim sXPath As String
'''''Dim lLength As Long
'''''
'''''' Create Recordset using the xsd schema
''''''Set oTableNode = voNL.Item(1).selectNodes("//" vsTableName).Item(0) ' selectSingleNode(sXPath)
'''''Set oRS = New ADODB.Recordset
'''''If oTableNode Is Nothing Then
'''''Set ConvertXMLtoRecordset = Nothing
'''''Exit Function
'''''End If
'''''
'''''' Iterate trough all fields
'''''For Each oFieldNode In oTableNode.childNodes
'''''If Not oFieldNode.Attributes Is Nothing Then
'''''' Retrieve Max Length
'''''lLength = 0
'''''' Find all records of current field
'''''sXPath = "//" & vsTableName & "/" & oFieldNode.baseName
'''''Set oNodeList = voNL.Item(1).selectNodes(sXPath)
'''''
'''''' Iterate trough all records
'''''For Each oRecordNode In oNodeList
'''''If Len(oRecordNode.Text) > lLength Then
'''''lLength = Len(oRecordNode.Text)
'''''End If
'''''Next
'''''
'''''' Add Field
'''''On Error Resume Next
'''''If lLength = 0 Then lLength = 1
'''''Call oRS.Fields.Append(oFieldNode.baseName, adVarCharlLength)
'''''End If
'''''Next
'''''
'''''' Add the data to the empty Recordset
'''''sXPath = "//" & vsTableName
'''''Set oNodeList = voNL.Item(1).selectNodes(sXPath)
'''''
'''''Call oRS.Open
'''''
'''''' Iterate trough all records
'''''For Each oRecordNode In oNodeList
'''''' Add Record
'''''Call oRS.ADDNEW
'''''
'''''' Iterate trough all fields of current record
'''''For Each oFieldNode In oRecordNode.childNodes
'''''If Len(oFieldNode.baseName) > 0 Then
'''''' Set value
'''''oRS.Fields(oFieldNode.baseName) = oFieldNode.Text
'''''End If
'''''Next
'''''Next
'''''
'''''' Return the Recordset
'''''If Not (oRS.BOF And oRS.EOF) Then Call oRS.MoveFirst
'''''Set ConvertXMLtoRecordset = oRS
'''''End Function
Private Sub cmdCreateDSN_Click()
''    Dim mDSNDetails As tDSNAttrib, mError As String
''    With mDSNDetails
''        .Database = "DB_Finance"
''        .Driver = "SQL Server"
''        .Server = "(local)"
''        .TrustedConnection = True   'True=Use NT authentication::False need user name and pwd
''        .PassWord = "FAUser"
''        .UserID = "FAUser"
''        .Dsn = "dsnFA"
''        .Description = "Saankhya "
''        .Type = eDBType.ServerBased
''        .SystemDSN = True           'Create a System DSN
''    End With
''
''    mError = DSNCreate(mDSNDetails)
''    If Trim(mError) = "" Then
''        MsgBox "Dsn Created"
''    Else
''        MsgBox mError
''        Debug.Print mError
''    End If
End Sub

Private Sub cmdFixVouchersUsingTransactions_Click()
    Dim objDB As New clsDB
    Dim RecTran As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim mLoopCount As Integer
    '-------------------------------------------------------'
    ' faVoucher                                             '
    '-------------------------------------------------------'
    Dim mintVoucherID_1                As Double
    '@intLocalBodyID_2  [int],
    '@intTransactionID_3    [bigint],
    Dim mintTransactionTypeID_4        As Long
    Dim mtnyVoucherTypeID_5            As Integer
    Dim mintVoucherNo_6                As Variant
    Dim mintBookNo_7                   As Variant
    Dim mdtDate_8                      As Variant
    Dim mfltAmount_9                   As Variant
    Dim mintInstrumentTypeID_10        As Variant
    Dim mvchInstrumentNo_11            As Variant
    Dim mdtInstrumentDate_12           As Variant
    Dim mvchDescription_13             As Variant
    Dim mnumZoneID_14                  As Variant
    Dim mnumWardID_15                  As Variant
    Dim mintDoorNoP1_16                As Variant
    Dim mvchDoorNoP2_17                As Variant
    Dim mvchDoorNoP3_18                As Variant
    Dim mintUserID_19                  As Variant
    Dim mintCounterID_20               As Variant
    Dim mnumSubLedgerID_21             As Variant
    Dim mintKeyID1_22                  As Variant
    Dim mintKeyID2_23                  As Variant
    Dim mintExternalApplicationID_24   As Variant
    Dim mintExternalModuleID_25        As Variant
    Dim mintFinancialYearID_26         As Variant
    
    Dim mvchBank_33                    As Variant
    Dim mvchBankPlace_34               As Variant
    Dim mintFundID_35                  As Variant
    Dim mRefNo                         As Variant
    Dim mRoundOff                      As Variant
    Dim mAdvAmtAdj                     As Variant
    
    
    
    '-------------------------------------------------------'
    ' faVoucher Child
    '-------------------------------------------------------'
    'Dim mintVoucherID_1       As Double  '
    Dim mintLocalBodyID_2       As Long
    Dim mintSlNo_3              As Long
    Dim mintAccountHeadID_4     As Long
    Dim mtnyDebitOrCredit_5     As Byte
    Dim mintYearID_6            As Long
    Dim mtnyPeriodID_7          As Variant
    Dim mtnyArrearFlag_8        As Variant
    Dim mnumDemandID_9          As Variant
    Dim mfltAmount_10           As Variant
    
    
    
    objDB.SetConnection mCnn
    mSql = "Select * From faTransactions Inner Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID AND faTransactionChild.intSerialNo = 1 Where intGroupID = 30  Order By faTransactions.intTransactionID"

    RecTran.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
    If (RecTran.BOF And RecTran.EOF) Then
        Exit Sub
    End If
    
    While Not RecTran.EOF
        
        mSql = "Select * From faTransactions  Inner Join faVouchers On faVouchers.intVoucherID = faTransactions.intVoucherID"
        mSql = mSql + " Where intGroupID = 30 And tnyVoucherTypeID = 30 And faTransactions.intTransactionID = " & RecTran!intTransactionID & " And faVouchers.vchDescription = faTransactions.vchNarration"
        mSql = mSql + " Order By faTransactions.intTransactionID "
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        If (Rec.BOF And Rec.EOF) Then
                
                '@intVoucherID_1     [bigint],
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                
                mintTransactionTypeID_4 = RecTran!intTransactionTypeID
                mtnyVoucherTypeID_5 = 30
                mintVoucherNo_6 = Null
                mintBookNo_7 = ""
                mdtDate_8 = RecTran!dtTransactionDate
                mfltAmount_9 = RecTran!fltAmount
                mintInstrumentTypeID_10 = RecTran!intTransactionTypeID
                mvchInstrumentNo_11 = Null 'RecTran!vchInstruementNo
                mdtInstrumentDate_12 = RecTran!dtTransactionDate
                mvchDescription_13 = RecTran!vchNarration
                mnumZoneID_14 = Null
                mnumWardID_15 = Null
                mintDoorNoP1_16 = Null
                mvchDoorNoP2_17 = Null
                mvchDoorNoP3_18 = Null
                mintUserID_19 = 9999
                mintCounterID_20 = 11
                mnumSubLedgerID_21 = Null ' mBuildingID ' Changed by Aiby on 10-Dec-2008 From Kollam Corp.
                mintKeyID1_22 = RecTran!intAccountHeadID
                mintKeyID2_23 = Null
                mintExternalApplicationID_24 = AppID.Saankhya
                mintExternalModuleID_25 = 0
                mintFinancialYearID_26 = gbFinancialYearID
                mvchBank_33 = Null
                mvchBankPlace_34 = Null
                mintFundID_35 = 1
                mRefNo = Null
                mRoundOff = Null
                mAdvAmtAdj = Null
                '========================================='
                ' BEGIN TRANSACTION                       '
                '-----------------------------------------'
                'mCnn.BeginTrans
                'On Error GoTo ErrorRollBack:
                '========================================='
                
                arrInput = Array( _
                -1, _
                gbLocalBodyID, _
                Null, _
                mintTransactionTypeID_4, _
                mtnyVoucherTypeID_5, _
                mintVoucherNo_6, _
                mintBookNo_7, _
                mdtDate_8, _
                mfltAmount_9, _
                mintInstrumentTypeID_10, _
                mvchInstrumentNo_11, _
                mdtInstrumentDate_12, _
                mvchDescription_13, _
                mnumZoneID_14, _
                mnumWardID_15, _
                mintDoorNoP1_16, _
                mvchDoorNoP2_17, _
                mvchDoorNoP3_18, _
                mintUserID_19, _
                mintCounterID_20, _
                mnumSubLedgerID_21, _
                mintKeyID1_22, mintKeyID2_23, mintExternalApplicationID_24, _
                mintExternalModuleID_25, mintFinancialYearID_26, gbShiftID, 1, 0, _
                mvchBank_33, mvchBankPlace_34, mintFundID_35, gbSeatID, gbSessionID, mRefNo, mRoundOff, mAdvAmtAdj)
                
                objDB.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID_1 = arrOutPut(0, 0)
                    
                Else
                    GoTo ErrorRollBack:
                End If
                
                'NOTE:- Fetching Transaction Child Data By TransactionID
                mSql = "Select * From faTransactionChild Where intSerialNo <> 1 AND intTransactionID  = " & RecTran!intTransactionID
                RecChild.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                If Not (RecChild.BOF And RecChild.EOF) Then
                    While Not RecChild.EOF
                            mLoopCount = mLoopCount + 1
                            mintLocalBodyID_2 = gbLocalBodyID
                            mintSlNo_3 = mLoopCount
                            mintAccountHeadID_4 = RecChild!intAccountHeadID
                            mtnyDebitOrCredit_5 = RecChild!tinDebitOrCreditFlag
                            mintYearID_6 = gbFinancialYearID
                            mtnyPeriodID_7 = Null
                            mtnyArrearFlag_8 = Null
                            mnumDemandID_9 = Null
                            mfltAmount_10 = RecChild!fltAmount
                            
                            Set arrInput = Nothing
                            arrInput = Array( _
                            mintVoucherID_1, _
                            mintLocalBodyID_2, _
                            mintSlNo_3, _
                            mintAccountHeadID_4, _
                            mtnyDebitOrCredit_5, _
                            mintYearID_6, _
                            mtnyPeriodID_7, _
                            mtnyArrearFlag_8, _
                            mnumDemandID_9, _
                            mfltAmount_10 _
                            )
                            objDB.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                        
                        RecChild.MoveNext
                    Wend
                End If
                RecChild.Close
                
                'NOTES:-Updating Voucher ID in Transaction Table
                RecTran!intVoucherID = mintVoucherID_1
                RecTran.Update
                
                
                'mCnn.CommitTrans
                GoTo NextRec:
ErrorRollBack:
                'mCnn.RollbackTrans
                
                
        End If
NextRec:
        Rec.Close
        RecTran.MoveNext
    Wend
    
    
    
    
    
End Sub

Private Sub cmdFixVouchersUsingTransactionsJournal_Click()

    Dim objDB As New clsDB
    Dim RecTran As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim mLoopCount As Integer
    '-------------------------------------------------------'
    ' faVoucher                                             '
    '-------------------------------------------------------'
    Dim mintVoucherID_1                As Double
    '@intLocalBodyID_2  [int],
    '@intTransactionID_3    [bigint],
    Dim mintTransactionTypeID_4        As Long
    Dim mtnyVoucherTypeID_5            As Integer
    Dim mintVoucherNo_6                As Variant
    Dim mintBookNo_7                   As Variant
    Dim mdtDate_8                      As Variant
    Dim mfltAmount_9                   As Variant
    Dim mintInstrumentTypeID_10        As Variant
    Dim mvchInstrumentNo_11            As Variant
    Dim mdtInstrumentDate_12           As Variant
    Dim mvchDescription_13             As Variant
    Dim mnumZoneID_14                  As Variant
    Dim mnumWardID_15                  As Variant
    Dim mintDoorNoP1_16                As Variant
    Dim mvchDoorNoP2_17                As Variant
    Dim mvchDoorNoP3_18                As Variant
    Dim mintUserID_19                  As Variant
    Dim mintCounterID_20               As Variant
    Dim mnumSubLedgerID_21             As Variant
    Dim mintKeyID1_22                  As Variant
    Dim mintKeyID2_23                  As Variant
    Dim mintExternalApplicationID_24   As Variant
    Dim mintExternalModuleID_25        As Variant
    Dim mintFinancialYearID_26         As Variant
    
    Dim mvchBank_33                    As Variant
    Dim mvchBankPlace_34               As Variant
    Dim mintFundID_35                  As Variant
    Dim mRefNo                         As Variant
    Dim mRoundOff                      As Variant
    Dim mAdvAmtAdj                     As Variant
    
    
    
    '-------------------------------------------------------'
    ' faVoucher Child
    '-------------------------------------------------------'
    'Dim mintVoucherID_1       As Double  '
    Dim mintLocalBodyID_2       As Long
    Dim mintSlNo_3              As Long
    Dim mintAccountHeadID_4     As Long
    Dim mtnyDebitOrCredit_5     As Byte
    Dim mintYearID_6            As Long
    Dim mtnyPeriodID_7          As Variant
    Dim mtnyArrearFlag_8        As Variant
    Dim mnumDemandID_9          As Variant
    Dim mfltAmount_10           As Variant
    
    
    
    objDB.SetConnection mCnn
    mSql = "Select * From faTransactions Inner Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID AND faTransactionChild.intSerialNo = 1 Where intGroupID = 40  Order By faTransactions.intTransactionID"

    RecTran.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
    If (RecTran.BOF And RecTran.EOF) Then
        Exit Sub
    End If
    
    While Not RecTran.EOF
        
        mSql = "Select * From faTransactions  Inner Join faVouchers On faVouchers.intVoucherID = faTransactions.intVoucherID"
        mSql = mSql + " Where intGroupID = 40 And tnyVoucherTypeID = 40 And faTransactions.intTransactionID = " & RecTran!intTransactionID & " And faVouchers.vchDescription = faTransactions.vchNarration"
        mSql = mSql + " Order By faTransactions.intTransactionID "
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        If (Rec.BOF And Rec.EOF) Then
                
                '@intVoucherID_1     [bigint],
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                
                mintTransactionTypeID_4 = RecTran!intTransactionTypeID
                mtnyVoucherTypeID_5 = 40
                mintVoucherNo_6 = Null
                mintBookNo_7 = ""
                mdtDate_8 = RecTran!dtTransactionDate
                mfltAmount_9 = RecTran!fltAmount
                mintInstrumentTypeID_10 = RecTran!intTransactionTypeID
                mvchInstrumentNo_11 = Null 'RecTran!vchInstruementNo
                mdtInstrumentDate_12 = RecTran!dtTransactionDate
                mvchDescription_13 = RecTran!vchNarration
                mnumZoneID_14 = Null
                mnumWardID_15 = Null
                mintDoorNoP1_16 = Null
                mvchDoorNoP2_17 = Null
                mvchDoorNoP3_18 = Null
                mintUserID_19 = 9999
                mintCounterID_20 = 11
                mnumSubLedgerID_21 = Null ' mBuildingID ' Changed by Aiby on 10-Dec-2008 From Kollam Corp.
                mintKeyID1_22 = RecTran!intAccountHeadID
                mintKeyID2_23 = Null
                mintExternalApplicationID_24 = AppID.Saankhya
                mintExternalModuleID_25 = 0
                mintFinancialYearID_26 = gbFinancialYearID
                mvchBank_33 = Null
                mvchBankPlace_34 = Null
                mintFundID_35 = 1
                mRefNo = Null
                mRoundOff = Null
                mAdvAmtAdj = Null
                '========================================='
                ' BEGIN TRANSACTION                       '
                '-----------------------------------------'
                'mCnn.BeginTrans
                'On Error GoTo ErrorRollBack:
                '========================================='
                
                arrInput = Array( _
                -1, _
                gbLocalBodyID, _
                Null, _
                mintTransactionTypeID_4, _
                mtnyVoucherTypeID_5, _
                mintVoucherNo_6, _
                mintBookNo_7, _
                mdtDate_8, _
                mfltAmount_9, _
                mintInstrumentTypeID_10, _
                mvchInstrumentNo_11, _
                mdtInstrumentDate_12, _
                mvchDescription_13, _
                mnumZoneID_14, _
                mnumWardID_15, _
                mintDoorNoP1_16, _
                mvchDoorNoP2_17, _
                mvchDoorNoP3_18, _
                mintUserID_19, _
                mintCounterID_20, _
                mnumSubLedgerID_21, _
                mintKeyID1_22, mintKeyID2_23, mintExternalApplicationID_24, _
                mintExternalModuleID_25, mintFinancialYearID_26, gbShiftID, 1, 0, _
                mvchBank_33, mvchBankPlace_34, mintFundID_35, gbSeatID, gbSessionID, mRefNo, mRoundOff, mAdvAmtAdj)
                
                objDB.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID_1 = arrOutPut(0, 0)
                    
                Else
                    GoTo ErrorRollBack:
                End If
                
                'NOTE:- Fetching Transaction Child Data By TransactionID
                mSql = "Select * From faTransactionChild Where intSerialNo <> 1 AND intTransactionID  = " & RecTran!intTransactionID
                RecChild.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                If Not (RecChild.BOF And RecChild.EOF) Then
                    While Not RecChild.EOF
                            mLoopCount = mLoopCount + 1
                            mintLocalBodyID_2 = gbLocalBodyID
                            mintSlNo_3 = mLoopCount
                            mintAccountHeadID_4 = RecChild!intAccountHeadID
                            mtnyDebitOrCredit_5 = RecChild!tinDebitOrCreditFlag
                            mintYearID_6 = gbFinancialYearID
                            mtnyPeriodID_7 = Null
                            mtnyArrearFlag_8 = Null
                            mnumDemandID_9 = Null
                            mfltAmount_10 = RecChild!fltAmount
                            
                            Set arrInput = Nothing
                            arrInput = Array( _
                            mintVoucherID_1, _
                            mintLocalBodyID_2, _
                            mintSlNo_3, _
                            mintAccountHeadID_4, _
                            mtnyDebitOrCredit_5, _
                            mintYearID_6, _
                            mtnyPeriodID_7, _
                            mtnyArrearFlag_8, _
                            mnumDemandID_9, _
                            mfltAmount_10 _
                            )
                            objDB.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                        
                        RecChild.MoveNext
                    Wend
                End If
                RecChild.Close
                
                'NOTES:-Updating Voucher ID in Transaction Table
                RecTran!intVoucherID = mintVoucherID_1
                RecTran.Update
                
                
                'mCnn.CommitTrans
                GoTo NextRec:
ErrorRollBack:
                'mCnn.RollbackTrans
                
                
        End If
NextRec:
        Rec.Close
        RecTran.MoveNext
    Wend
    
    


End Sub



Private Sub cmdLFA_Click()
    Dim mArrIN As Variant
    Dim mArrOut As Variant
    Dim mUrl   As String
    Dim client1 As New MSSOAPLib.SoapClient
    Dim objSOAP As Variant
    Dim clnt As New SoapClient30
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim mSql  As String
    Dim mLoop As Integer
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya

    '--------------'
    ' Web Service  '
    '--------------'

    'mUrl = "http://172.16.2.220:8080/CalculatorWSApplication/NewWebService"
    
    mUrl = "http://117.239.77.103:9081/eGovWebServices/SaankhyaIntegrationService"
    'mUrl = "http://172.16.1.112:8080/webservices/accountService"
    'mArrIN = Array(mReqInboxID, gbLBID, txtRequisition.Text)
    
    Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
    objSOAP.MSSoapInit mUrl + "?WSDL"
    
    'SoapClient.ConnectorProperty("AuthName") = "username"
    'SoapClient.ConnectorProperty("AuthPassword") = "userpwd"
        
    Call objSOAP.methodWithNoParameter
    
    
    
    Dim objSOA As New MSSOAPLib30.SoapClient30
    objSOA.MSSoapInit mUrl + "?WSDL"
    
    'objsoa.ClientProperty(
    'objSoapClient.ConnectorProperty ("AuthUser")
    'objSOA.MSSoapInit2
    
    Dim X As Integer
    'X = objSOAP.addition(10, 20)
    
    objSOAP.submitNone
    Call objSOAP.submitNoneValues("ikm")
    
'    Dim objac(0) As objSOAP.accounts
'    objac(0).Amount = 100
'    objac(0).LBCode = "G1002"
'    objac(0).headofAccount = "450100101"
'
'
''    Dim objac As New Collection
''
''    Dim mAc As uAcc
''    mAc.LBCode = "G1001"
''    mAc.HeadCode = "450450250"
''    mAc.Amount = 50000
''
''
''    objAC.Add mAc
''    objAC.Add mAc
''
'    Dim mArr(3, 0) As Variant
'
'    'ReDim Preserve mArr(3, 0)
'    mArr(0, 0) = "G101"
'    mArr(1, 0) = "450450250"
'    mArr(2, 0) = "5000"
    
    'ReDim Preserve mArr(3, 1)
    'mArr(0, 1) = 101
    'mArr(1, 1) = "450450350"
    'mArr(2, 1) = 65000
    
    '    Dim uAr As uAcc
    '    uAr.LBCode = "50000"
    '    uAr.HeadCode = "450450250"
    '    uAr.Amount = "G101"
    '    mArr(0) = Mar

'objSOAP.SubmitAccountin2dArray (mArr)
    
    'Call objSOAP.submitParameter
    
    
    
'    For mLoop = 1 To vsGrid.Rows - 1
'        If vsGrid.TextMatrix(mLoop, 1) <> "" Then
'          mArrOut = objSOAP.SyncUpdateSyncFlagToRequisitionInbox(val(vsGrid.TextMatrix(mLoop, 14)), gbLBID, vsGrid.TextMatrix(mLoop, 1))
'
'          mSQL = "Update faRequisitionInbox set  tnyStage=2 WHERE intID=" & val(vsGrid.TextMatrix(mLoop, 14)) & "  "
'          objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
'        End If
'    Next mLoop
End Sub

Private Sub cmdLFA2_Click()
        Dim xmlHttp As Object
        Set xmlHttp = CreateObject("MSXML2.XmlHttp")
        Dim Para As String
        Dim X As Integer
        
        Para = ""
        Para = "lbCode=TVM&year=2015"
        For X = 1 To 24
            Para = Para + "&ielarpItem=L&codeNo=310000000&description=Panchayat Fund&schedule=B-1&amount=" & (5000 + X)
        Next
        xmlHttp.Open "GET", "http://202.88.240.97:8084/idms/home.action?b=" & Para, False
        xmlHttp.send
        MsgBox xmlHttp.responseText
End Sub

Private Sub Cochin_Click()
    Dim mArrIN As Variant
    Dim mArrOut As Variant
    Dim mUrl   As String
    Dim client1 As New MSSOAPLib.SoapClient
    Dim objSOAP As Variant
    Dim clnt As New SoapClient30
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim mSql  As String
    Dim mLoop As Integer
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya

    '--------------'
    ' Web Service  '
    '--------------'
    
    
    mUrl = "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService/{any integer varible}"
    mUrl = "http://117.239.77.103:9081/eGovWebServices/SaankhyaIntegrationService"
    'mUrl = "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService"
    'mArrIN = Array(mReqInboxID, gbLBID, txtRequisition.Text)
    
    Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
    objSOAP.MSSoapInit mUrl + "?WSDL"
    Dim objFrom As New FormTwoDtlsVO
    
    objFrom.id = 1001
    objFrom.doorNo = "101"
    objFrom.doorNo2 = "A"
    objFrom.applicantName = "Hello World"
    objFrom.ownerPin = "695017"
    objFrom.wardNo = "1"
    Call objSOAP.methodWithIntegerParameter(1)
    'Call objSOAP.searchAssesmentDetails '1'
    
   Set mArrOut = objSOAP.methodWithIntegerParameter(1)
    MsgBox (mArrOut)
    
    
End Sub

Private Sub cochinFinal_Click()
     Dim xmlHttp As Object
       Set xmlHttp = CreateObject("MSXML2.XmlHttp")
       Dim mXmlString   As Variant
       Dim Rec  As New ADODB.Recordset
        Dim REcDet As New ADODB.Recordset
        Dim params
        params = "lbCode=G051105&year=2014-2015”"
        '//  doorNo~doorNo2~AssessmentNo~applicationNo~applicationStatus~ownerName~zoneId~wardId~ownerPin~ownerAddress
        '// 25/1~ ~ ~ ~ ~ ~9~ ~ ~
        ' OR
        '// 25/1~NA~NA~NA~NA~NA~9~NA~NA~NA


        Set Rec = New ADODB.Recordset
        params = "9~NA~25/1~NA~NA~NA~NA~NA~NA~NA~NA"

        xmlHttp.Open "POST", "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService/searchAssesmentDetails?searchParam=params", False
        xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
        xmlHttp.send
        mXmlString = xmlHttp.responseText
        mXmlString = Replace(mXmlString, "UTF-8", "UTF-16")
        Dim mCnt    As Integer
        Dim oRS As ADODB.Recordset
        Dim oNode As Object 'MSXML2.IXMLDOMNode
        Dim oSubNodes As Object 'MSXML2.IXMLDOMSelection
        Dim oDoc As Object
        'Dim oDoc As MSXML2.DOMDocument26
        Set oDoc = CreateObject("MSXML2.DOMDocument")
        'Set oNode = CreateObject("MSXML2.IXMLDOMNode")
    '    Set oSubNodes = CreateObject("MSXML2.IXMLDOMSelection")
        
        oDoc.async = False
        oDoc.validateOnParse = False
        If Not oDoc.LoadXml(mXmlString) Then
            MsgBox "Error Loading"
            Exit Sub
        Else
            MsgBox "Sucess"
        End If
    
        Set oRS = New ADODB.Recordset
        Set oRS.ActiveConnection = Nothing
        oRS.CursorLocation = adUseClient
        oRS.LockType = adLockBatchOptimistic
    
        With oRS.Fields
            .Append "actionCode", adVarChar, 10
            .Append "ownerName", adVarChar, 200
            .Append "applicationNo", adVarChar, 200
            .Append "currentHalfYearPaid", adBoolean
            .Append "doorNo", adVarChar, 50
            .Append "id", adInteger
            .Append "ownerAddress", adVarChar, 200
            .Append "ownerPin", adVarChar, 10
            .Append "statusText", adVarChar, 200
            .Append "wardNo", adInteger
            .Append "workFlowLevel", adInteger
            .Append "zoneNo", adInteger
           
        End With
       mCnt = 0
       vsGrid.Rows = 2
        oRS.Open
         For Each oNode In oDoc.selectNodes("/propertyTaxVOes/PropertyTaxVO")
            oRS.ADDNEW
            mCnt = mCnt + 1
            vsGrid.Rows = vsGrid.Rows + 1
            'oRS.Fields("ownerName").value = oNode.selectSingleNode("ownerName").Text
    '        oRS.Fields("applicationNo").value = oNode.selectSingleNode("applicationNo").Text
    '        oRS.Fields("currentHalfYearPaid").value = oNode.selectSingleNode("currentHalfYearPaid").Text
            'oRS.Fields("doorNo").value = oNode.selectSingleNode("doorNo").Text
            vsGrid.TextMatrix(mCnt, 0) = oNode.selectSingleNode("doorNo").Text
            'vsGrid.TextMatrix(mCnt, 2) = oNode.selectSingleNode("ownerName").Text
        Next
        
        

End Sub

Private Sub Command1_Click()
    smdSubTotal
    
    
End Sub
 

 

Private Sub smdSubTotal()

'Dim i As Integer
'
'Dim vbab As String
'
'    fg.Rows = 102
'
'    fg.OutlineCol = 0
'
'    fg.OutlineBar = flexOutlineBarSimpleLeaf
'
'    fg.AllowUserResizing = flexResizeColumns
'
'    fg.Editable = flexEDKbdMouse
'
'    Dim intOutlinelevel
'
'    intOutlinelevel = 0
'
'    fg.Rows = 1
'
'    fg.FixedRows = 1
'
'    For i = 1 To 100
'
'            If i = 1 Then
'
'                fg.AddItem "Rent"
'
'                fg.IsSubtotal(i) = True
'
'                fg.RowOutlineLevel(i) = intOutlinelevel
'
'                intOutlinelevel = intOutlinelevel + 1
'
'            ElseIf Len(CStr(i)) > 1 Then
'
'                fg.AddItem "" & vbab & "Rent" & CStr(i)
'
'                If i Mod 10 = 0 Then
'
'                    fg.IsSubtotal(i) = True
'
'                End If
'
'                fg.RowOutlineLevel(i) = intOutlinelevel
'
'            Else
'
'                fg.AddItem "" & vbab & "" & vbab & "Rent" & CStr(i)
'
'                'fg.IsSubtotal(i) = True
'
'                fg.RowOutlineLevel(i) = 2
'
'            End If
'
'    Next i

End Sub

Private Sub Command2_Click()
    Dim DataSourceName As String
   Dim DatabaseName As String
   Dim Description As String
   Dim DriverPath As String
   Dim DriverName As String
   Dim LastUser As String
   Dim Regional As String
   Dim Server As String

   Dim lResult As Long
   Dim hKeyHandle As Long

   'Specify the DSN parameters.

   DataSourceName = "dsnFAs"
   DatabaseName = "DB_Finance"
   Description = "Testing"
   DriverPath = "SQLSRV32.dll"
   LastUser = "FAUser"
   Server = "(local)"
   DriverName = "SQL Server"

   'Create the new DSN key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DataSourceName, hKeyHandle)

   'Set the values of the new DSN key.

   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
      ByVal DatabaseName, Len(DatabaseName))
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
      ByVal Description, Len(Description))
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
      ByVal DriverPath, Len(DriverPath))
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
      ByVal LastUser, Len(LastUser))
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
      ByVal Server, Len(Server))

   'Close the new DSN key.

   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)

   
                        
End Sub

Private Sub Command3_Click()
            Dim rptFileName As String
            Dim arrInput As Variant
            Set arrInput = Nothing
            Dim Rpt As New CRAXDRT.Report
            Dim mApp As New CRAXDRT.Application
            Dim mLoop As Long
            
            
            'mvarRptFileName = App.Path & "..\Reports\rptLedgerView.rpt"
            Debug.Print App.Path & "\Reports\rptLedgerView.rpt"
            
            rptFileName = App.Path & "\Reports\rptGEN-40.rpt"
            
            Screen.MousePointer = vbHourglass
            crvReport.DisplayToolbar = True
            
            Set Rpt = Nothing
            mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
            Set Rpt = mApp.OpenReport(rptFileName, 1)
            If IsArray(arrInput) Then
                For mLoop = LBound(arrInput) To UBound(arrInput)
                    Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
                    Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
                Next mLoop
            End If
            Screen.MousePointer = vbDefault
            
            crvReport.ReportSource = Rpt
            crvReport.Refresh
            'crvReport.Left = 0
            'crvReport.Top = 0
            
            crvReport.ViewReport
            crvReport.Zoom (2)
End Sub

Private Sub Command4_Click()
    Dim Rec As New ADODB.Recordset
    Dim oStream As New ADODB.Stream
    Dim strRecordset As String
    oStream.Open
    Rec.Save oStream, adPersistXML
    oStream.Position = 0
    strRecordset = oStream.ReadText
    'return the XML string representation of your recordset
    'RecordsetToXML = strRecordset
End Sub

Private Sub Command5_Click()
Dim mArrIN As Variant
        Dim mArrOut As Variant
        Dim mUrl   As String
        Dim client1 As New MSSOAPLib.SoapClient
        Dim objSOAP As Variant
        Dim clnt As New SoapClient30
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSql  As String
        Dim mReqIDCheck As Boolean
     
        Dim KEY As String
        Dim KEY2 As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
        '--------------'
        ' Web Service  '
        '--------------'
        mReqIDCheck = False
        mUrl = "http://localhost/SaankhyaService/SaankhyaService.asmx" 'gbDefaultUrlForRequisition
        mArrIN = Array(gbLBID, gbFinancialYearID)
        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        objSOAP.MSSoapInit mUrl + "?WSDL"
        
        Dim obj As Object
        Set obj = objSOAP.Test2
     
        mArrOut = objSOAP.SyncRequisitionInboxToLB(gbLBID, gbFinancialYearID)
        Dim mXmlStream As New ADODB.Stream
        mXmlStream.Open
        mXmlStream.WriteText mArrOut
        mXmlStream.Position = 0
        
        Dim Rec     As New ADODB.Recordset
        Dim RecID   As New ADODB.Recordset
        
        Rec.Open mXmlStream
        mXmlStream.Close
End Sub

