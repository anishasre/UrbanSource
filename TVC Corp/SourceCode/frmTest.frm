VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPast 
      Height          =   510
      Left            =   1440
      TabIndex        =   39
      Top             =   5535
      Width           =   1185
   End
   Begin VB.CommandButton cmdIRR 
      Caption         =   "IR Register"
      Height          =   465
      Left            =   7830
      TabIndex        =   38
      Top             =   810
      Width           =   915
   End
   Begin VB.CommandButton cmdPTax 
      Caption         =   "PTax"
      Height          =   510
      Left            =   2475
      TabIndex        =   37
      Top             =   270
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1920
      Left            =   2160
      TabIndex        =   36
      Top             =   2160
      Width           =   1380
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   8955
      TabIndex        =   35
      Text            =   "0.0.0.0"
      Top             =   2955
      Width           =   2475
   End
   Begin RichTextLib.RichTextBox txtInsertQRY 
      Height          =   2040
      Left            =   45
      TabIndex        =   34
      Top             =   6345
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   3598
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmTest.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTBL 
      Height          =   375
      Left            =   4815
      TabIndex        =   33
      Top             =   5310
      Width           =   1995
   End
   Begin VB.TextBox txtQRY 
      Height          =   1005
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   5265
      Width           =   4875
   End
   Begin VB.CommandButton cmdInsertQRY 
      Caption         =   "QRY Insert"
      Height          =   420
      Left            =   4815
      TabIndex        =   31
      Top             =   5760
      Width           =   1365
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   525
      Left            =   2640
      TabIndex        =   30
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   330
      Width           =   1635
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   450
      Left            =   8970
      TabIndex        =   28
      Top             =   2340
      Width           =   2340
   End
   Begin VB.CommandButton cmdGetMac 
      Caption         =   "Get Mac"
      Height          =   510
      Left            =   8895
      TabIndex        =   27
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Generate Insert Script"
      Height          =   540
      Left            =   8910
      TabIndex        =   26
      Top             =   1635
      Width           =   2430
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9810
      TabIndex        =   24
      Top             =   4170
      Width           =   1560
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Transactions Vs Vouchers"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8910
      TabIndex        =   23
      Top             =   960
      Width           =   2430
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   330
      Top             =   5070
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\IKM\SaankhyaDoubleEntry\Bank.xls"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\IKM\SaankhyaDoubleEntry\Bank.xls"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7185
      TabIndex        =   22
      Text            =   "ØÞCcÞ çØÞËíxíæÕÏV"
      Top             =   3060
      Width           =   1365
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Find Difference in Transaction Child and Voucher Child"
      Height          =   585
      Left            =   9255
      TabIndex        =   21
      Top             =   4605
      Width           =   2475
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Fixing Receipt Voucher"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6870
      TabIndex        =   19
      Top             =   3825
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortVoucherNo 
      Caption         =   "Sort Voucher No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4590
      TabIndex        =   18
      Top             =   4815
      Width           =   1950
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Major Heads"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4575
      TabIndex        =   17
      Top             =   4275
      Width           =   1950
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Major Head wiseReceipt && Payments"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2505
      TabIndex        =   16
      Top             =   4800
      Width           =   1950
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Receipts && Payments"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2490
      TabIndex        =   15
      Top             =   4320
      Width           =   1965
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   450
      TabIndex        =   14
      Top             =   4320
      Width           =   1680
   End
   Begin VB.CommandButton cmdDifferenceInVoucher 
      Caption         =   "Fixing Payment Voucher"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6870
      TabIndex        =   13
      Top             =   4335
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintReceipt 
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5370
      TabIndex        =   12
      Top             =   3195
      Width           =   1380
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRV 
      Height          =   2085
      Left            =   3840
      TabIndex        =   11
      Top             =   675
      Width           =   3915
      lastProp        =   500
      _cx             =   6906
      _cy             =   3678
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   255
      Width           =   3660
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   9
      Top             =   3900
      Width           =   1635
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   8
      Top             =   3450
      Width           =   1635
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2970
      Width           =   1605
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3555
      TabIndex        =   6
      Top             =   3660
      Width           =   1695
   End
   Begin VB.TextBox txtReceiptNo 
      Height          =   285
      Left            =   3570
      TabIndex        =   5
      Top             =   3300
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   465
      TabIndex        =   4
      Top             =   2430
      Width           =   1635
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   465
      TabIndex        =   3
      Top             =   1890
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   450
      TabIndex        =   2
      Top             =   1350
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   465
      TabIndex        =   1
      Top             =   825
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Voucher Vs Transactions"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8910
      TabIndex        =   0
      Top             =   315
      Width           =   2430
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   165
      Left            =   9285
      TabIndex        =   25
      Top             =   4230
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ReceiptNo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2655
      TabIndex        =   20
      Top             =   3300
      Width           =   870
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CntrlKey As Boolean

'*****************************************************************************************
'* Application ID           :                                                            *
'* Application Name         : Saankhya Double Entry ( NMA )                              *
'* Screen id                : Payments                                                   *
'* Version No               : Ver 2.0.0                                                  *
'* Form Designed By         :                                                            *
'* Created on               :                                                            *
'* Coded By                 :                                                            *
'* Coded on                 :                                                            *
'* Reviewed By              :                                                            *
'* Reviewed on              :                                                            *
'* Purpose                  :                                                            *
'*                          :                                                            *
'*                          :                                                            *
'* Name of Database         : DB_Finance                                                 *
'* DSN                      : dsnFA ( UserName=FAUser; PWD=FAUser )                      *
'* Name of Table(s)         :                                                            *
'* Look up Table(s)         :                                                            *
'*                          :                                                            *
'*                          :                                                            *
'* Stored Procedures        :                                                            *
'*                          :                                                            *
'*                          :                                                            *
'*=======================================================================================*
    
    Dim mTest As New Collection
    Private objAcc As New clsAccounts
    Dim mCon As ADODB.Connection
    Function ExcelData(sFileNameAndPath As String, sSheetName As String) As Variant
   '-- return a 2-D array if no error happens
   Dim xlApp As Object
   Dim wb As Object
   
   '-- may need error handler
   Set xlApp = CreateObject("Excel.Application")
   Set wb = xlApp.Workbooks.Open(sFileNameAndPath)
   ExcelData = wb.Worksheets(sSheetName).UsedRange
   wb.Close False
   Set wb = Nothing
   xlApp.Quit
   Set xlApp = Nothing
End Function

Sub ReadDataFromExcel()
   
End Sub
    Public Function Read_Excel(ByVal sFile As String) As ADODB.Recordset
    
          On Error GoTo fix_err
          Dim Rec As ADODB.Recordset
          Set Rec = New ADODB.Recordset
          Dim mCn As String
    
          Rec.CursorLocation = adUseClient
          Rec.CursorType = adOpenKeyset
          Rec.LockType = adLockBatchOptimistic
    
          mCn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sFile
          Rec.Open "SELECT * FROM [CORPO$]", mCn
          Set Read_Excel = Rec
          Set Rec = Nothing
          Exit Function
fix_err:
          Debug.Print err.Description + " " + _
                      err.Source, vbCritical, "Import"
          err.Clear
    End Function

'Private Sub PrintReceipt_ForNewFormat(intVoucherID As Double)
'' NEW FORMAT FOR  SAANKHYA SOOCHIKA Modified on 11-Oct-2011 (Aiby)
''        gbFileNO = FreeFile
''        gbFileName = "C:\Report.txt"
''        Open gbFileName For Output As #gbFileNO
''        Print #gbFileNO, Chr$(27) + Chr$(80)
''        Print #gbFileNO, String(136, "-")
''        Close #gbFileNO
''        Shell "Print " & gbFileName
''------------------------------------------------------------------------------------------------------------'
''-----------------------------------------Printing in 17 CPI-------------------------------------------------'
''------------------------------------------------------------------------------------------------------------'
'        Dim objDb As New clsDB
'        Dim mCnn As New ADODB.Connection
'        Dim Rec As New ADODB.Recordset
'        Dim mSql As String
'        Dim mLoop As Long
'        Dim mstrYear As String
'        Dim mCount As Long
'        Dim objCounter As New clsCounter
'        Dim objUser As New clsUser
'        Dim mName As String
'        Dim mChequeNo As String
'        Dim mStrInWard As String
'        Dim mRupees As String
'        Dim mStr1 As String
'        Dim mStr2 As String
'        Dim mInwardNo As String
'
'        'PrinterInit
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        If Len(Dir(gbFileName)) Then
'            Kill gbFileName
'        End If
'
'        'FileInitialize
'''''        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
'''''        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
'''''        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
'''''        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
'''''        mSql = mSql + " Left Join faPeriodicity On  faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
'''''        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
'''''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        objDb.SetConnection mCnn
'        Rec.CursorLocation = adUseClient
'        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic
'
'''''''        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
'''''''            If Rec.RecordCount > 9 Then
'''''''                Rec.Close
'''''''                Call PrintSummaryReceiptPTax(intVoucherID)
'''''''                Exit Sub
'''''''            End If
'''''''        End If
'        Open gbFileName For Output As #gbFileNO
'
'        Print #gbFileNO, Chr$(27) + Chr$(80); ' Set to 10 CPI
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        Print #gbFileNO, Tab(3); gbBold; gbDoubleWidth; "RECEIPT"; Tab(31); gbLBName; " Panchayat"; gbDoubleWidthOff
''        Select Case Rec!intInstrumentTypeID
''        Case Is = 1
''            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
''        Case Is = 4
''            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
''            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
''            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
''        Case Is = 5
''            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
''            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
''            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
''        Case Else
''            Print #gbFileNO,
''        End Select
'
'        If Not (Rec.EOF And Rec.BOF) Then
'            ' Line 6
'            'Print #gbFileNO, ; gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(65); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff;
'            ' Changed for KMBR By Cijith Sreedharan
'            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
'                If mKMBRFlag Or mSoochikaConnected Then
'                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
'                    'Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'                Else
'                    'Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'                End If
'            Else
'                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate));
'                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'            End If
'
'            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
'            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
'            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
'            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
'            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
'
'            Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff
'
'            'Changed for Sujith by Aiby - 24-Mar-2009
'
''            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
''            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
'
'            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(63); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
''            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
'            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(63); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
''            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
'            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
'            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
'
'            ' --------------------------------------------------------------------------------- '
'            ' To Print Check Number and DD Number Printing Phone Number is Commented
'            ' --------------------------------------------------------------------------------- '
'            Select Case Rec!intInstrumentTypeID
'            Case Is = 1
'                'Print #gbFileNO,
'            Case Is = 4
'                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                If Not IsNull(Rec!dtInstrumentDate) Then
'                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'                End If
'                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
'            Case Is = 5
'                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                If Not IsNull(Rec!dtInstrumentDate) Then
'                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'                End If
'                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
'            Case Else
'                'Print #gbFileNO,
'            End Select
'            Print #gbFileNO, ; gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
'            Print #gbFileNO, Tab(15); PadR(mChequeNo, 30);
'            Print #gbFileNO, Tab(57); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
'            Print #gbFileNO, Tab(72); PadR(mChequeNo, 32);
'            ' Line 15 Next
'            'Changed its Possition- Requested by Sujith on 24-Mar-2009
'            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
'
'            'Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
'            If Not (IsNull(Rec!vchRefNo)) Then
'                Print #gbFileNO, Tab(106); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", PadR(Rec!vchRefNo, 28))
'            Else
'                Print #gbFileNO,
'            End If
'                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
''                If Len(mStr1) < 47 Then
''                    mStr1 = mStr1 & String(47 - Len(mStr1), " ")
''                Else
''                    mStr1 = PadR(mStr1, 46)
''                End If
''                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
''                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
''                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
'            Print #gbFileNO, PadR(mStr1, 46); Tab(57); PadR(mStr1, 78)
'            'Print #gbFileNO,
'
'            ' Line 18 Next
'
'
'
'            Dim RecPTAX         As New ADODB.Recordset
'            Dim mStartingYear   As Integer
'            Dim mStartingPeriod As Integer
'            Dim mEndingYear     As Integer
'            Dim mEndingPeriod   As Integer
'            Dim mNarration      As String
'
'            mStartingYear = 2100
'
'
'
'            'If Rec!intTransactionTypeID = gbTransactionTypePTax Then
'            If Rec.RecordCount > 9 Then
'                mSql = "Select faVoucherChild.intAccountHeadID,Sum(fltAmount) As Amount,vchAccountHeadCode,vchAlias,tnyArrearFlag From faVoucherChild"
'                mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
'                mSql = mSql + " Where intVoucherID =" & intVoucherID '& Rec!intVoucherID
'                mSql = mSql + " Group By faVoucherChild.intAccountHeadID,vchAccountHeadCode,vchAlias,tnyArrearFlag"
'                mSql = mSql + " Order By tnyArrearFlag Desc,vchAccountHeadCode Desc"
'                RecPTAX.Open mSql, mCnn
'                While Not RecPTAX.EOF
'                    mLoop = mLoop + 1
'                    Print #gbFileNO, IIf(IsNull(RecPTAX!vchAccountHeadCode), "", RecPTAX!vchAccountHeadCode);
'                    Print #gbFileNO, Tab(37); PadL(Format(RecPTAX!amount, "0.00"), 9);
'                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
'                    Print #gbFileNO, Tab(58); PadR(RecPTAX!vchAlias, 46);
'                    Print #gbFileNO, Tab(127); PadL(Format(RecPTAX!amount, "0.00"), 9)
'                    RecPTAX.MoveNext
'                Wend
'                RecPTAX.Close
'                While Not Rec.EOF
'                    If mStartingYear > Rec!intYearID Then
'                        mStartingYear = Rec!intYearID
'                        mStartingPeriod = Rec!tnyPeriodID
'                    End If
'                    If mEndingYear < Rec!intYearID Then
'                        mEndingYear = Rec!intYearID
'                    End If
'                    mEndingPeriod = Rec!tnyPeriodID
'                    Rec.MoveNext
'                Wend
'                'Rec.Close
'                Rec.MoveFirst
'                Print #gbFileNO,
'                mLoop = mLoop + 1
'                mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
'                Print #gbFileNO, mNarration; Tab(54); mNarration
'                mLoop = mLoop + 1
'
'                mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
'                If mStartingPeriod = 1 Then
'                    mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
'                ElseIf mStartingPeriod = 2 Then
'                    mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
'                Else
'                    mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
'                End If
'
'                If mEndingPeriod = 1 Then
'                    mNarration = mNarration & " Ist Hf )"
'                ElseIf mEndingPeriod = 2 Then
'                    mNarration = mNarration & " IInd Hf )"
'                Else
'                    mNarration = mNarration & ")"
'                End If
'                mLoop = mLoop + 1
'                Print #gbFileNO, mNarration; Tab(52); mNarration
'            Else
''               GoTo LB ' To print Property Tax containing less than 9 rows
''           End If
''            Else
'
''LB:
'                mLoop = 0
'                Rec.MoveFirst
'                While Not Rec.EOF
'                    mLoop = mLoop + 1
'                    '==================================================================='
'                    ' Counter Foil
'                    '==================================================================='
'                    Print #gbFileNO, IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode);
'                    If Not IsNull(Rec!intYearID) Then
'                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
'                    Else
'                        mstrYear = ""
'                    End If
'                    Select Case Rec!tnyPeriodID
'                        Case Is = 1: Print #gbFileNO, Tab(12); mstrYear & "/1Hf";
'                        Case Is = 2: Print #gbFileNO, Tab(12); mstrYear & "/2Hf";
'                        Case Is = 3: Print #gbFileNO, Tab(12); mstrYear & "/F";
'                        Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
'
'                    End Select
'
'                    If Rec!intYearID < gbFinancialYearID Then
'                        Print #gbFileNO, Tab(26); PadL(Format(Rec!fltAmount, "0.00"), 9);
'                    Else
'                        Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
'                    End If
'
'                    '==================================================================='
'                    ' Receipt Area
'                    '==================================================================='
'                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
'                    Print #gbFileNO, Tab(58); PadR(Rec!vchAlias, 46);
'                    If Not IsNull(Rec!intYearID) Then
'                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
'                    Else
'                        mstrYear = ""
'                    End If
'                    Select Case Rec!tnyPeriodID
'                        Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
'                        Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
'                        Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
'                        Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
'                    End Select
'
'                    If Rec!intYearID < gbFinancialYearID Then
'                        Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                    Else
'                        Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                    End If
'                    Rec.MoveNext
'                Wend
'            End If
'            Rec.MoveFirst
'
'            For mCount = mLoop + 1 To 9
'                Print #gbFileNO,
'            Next mCount
'            If Rec!fltAdvAmtAdj > 0 Then
'                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 15); Tab(54); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 20);
'            Else
''                Print #gbFileNO,'Commented By Vinod
'            End If
'            Print #gbFileNO, Tab(25); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(116); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
'
'            Print #gbFileNO, Tab(25); "Total :"; Tab(36); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
'            Print #gbFileNO, Tab(116); "Total :"; Tab(128); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
'
'            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
'            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
'
'            mRupees = Rupees(Rec!TotalAmt)
'            If Len(mRupees) < 186 Then
'                mRupees = mRupees + String(185 - Len(mRupees), " ")
'            End If
'            'Print #gbFileNO, Tab(12); Left(mRupees, 34);
'            Print #gbFileNO, Tab(54); Left(mRupees, 75)
'
'            'Print #gbFileNO, Tab(12); mID$(mRupees, 33, 34);
'            'Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)
'
'            'Print #gbFileNO,'Commented By Vinod
'            Dim mInward As String
'            If Not IsNull(Rec!numInwardNo) Then
'                mInward = Rec!numInwardNo
'            Else
'                mInward = ""
'            End If
'
'            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
'            Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
'
'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
'                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
'            Else
'                Print #gbFileNO,
'            End If
'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
'                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
''            Else
''                Print #gbFileNO,
'            End If
'
'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 46 Then
'                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
'            Else
'                Print #gbFileNO,
'            End If
'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
'                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
''            Else
''                Print #gbFileNO,
'            End If
'
'
''             objCounter.SetCounter (Rec!intCounterID)
''            If objCounter.CounterID > 0 Then
''                Print #gbFileNO, Tab(30); objCounter.CounterNo;
''                Print #gbFileNO, Tab(67); objCounter.CounterNo & " : " & objCounter.CounterDescription
''            End If
''            objUser.SetUser (Rec!intUserID)
''            If objUser.UserID > -1 Then
''                Print #gbFileNO, Tab(27); objUser.UserName;
''                Print #gbFileNO, Tab(67); objUser.UserName
''            End If
'
'            objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                objUser.SetUser (Rec!intUserID)
'                If objUser.UserID > -1 Then
'                    Print #gbFileNO, Tab(27); objCounter.CounterNo; Tab(31); objUser.UserName;
'                    Print #gbFileNO, Tab(66); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(93); objUser.UserName
'                End If
'            End If
'
'
'            'Print #gbFileNO,
'        End If
'
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
'finishprinting:
'        Close #gbFileNO
'        'ShellPad
'        Shell "Print " & gbFileName
'        'Kill gbFileName
'
'End Sub
'

'
'
'    Private Sub PrintReceipt(intVoucherID As Double)
'
'
'        Dim objDb As New clsDB
'        Dim mCnn As New ADODB.Connection
'        Dim Rec As New ADODB.Recordset
'        Dim mSql As String
'        Dim mLoop As Long
'        Dim mstrYear As String
'        Dim mCount As Long
'        Dim objCounter As New clsCounter
'        Dim objUser As New clsUser
'        Dim mName As String
'        Dim mChequeNo As String
'
'
'        'PrinterInit
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        If Len(Dir(gbFileName)) Then
'            Kill gbFileName
'        End If
'        Open gbFileName For Output As #gbFileNO
'        'FileInitialize
'        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
'        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
'        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
'        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
'        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
'        objDb.SetConnection mCnn
'        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
'
'        Print #gbFileNO,
'        Print #gbFileNO,
'        Print #gbFileNO,
'
'        Select Case Rec!intInstrumentTypeID
'
'        Case Is = 1
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
'        Case Is = 4
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
'            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Is = 5
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
'            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Else
'            Print #gbFileNO,
'        End Select
'
'        If Not (Rec.EOF And Rec.BOF) Then
'            ' Line 6
'            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(120); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(65); IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'
'            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
'            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
'            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
'            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
'            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
'
'            Print #gbFileNO, Tab(15); Style(mName, True); Tab(65); Style(mName, True)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(65); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(65); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(65); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(65); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
'            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
'            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
'
'            ' --------------------------------------------------------------------------------- '
'            ' To Print Check Number and DD Number Printing Phone Number is Commented
'            ' --------------------------------------------------------------------------------- '
'            Select Case Rec!intInstrumentTypeID
'            Case Is = 1
'                Print #gbFileNO,
'            Case Is = 4
'                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
'            Case Is = 5
'                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
'            Case Else
'                Print #gbFileNO,
'            End Select
'
'            ' Line 15 Next
'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
'            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(55); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
'            Print #gbFileNO,
'            Print #gbFileNO,
'
'            ' Line 18 Next
'
'            Rec.MoveFirst
'            While Not Rec.EOF
'                mLoop = mLoop + 1
'
'                '==================================================================='
'                ' Counter Foil
'                '==================================================================='
'                Print #gbFileNO, IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode);
'                If Not IsNull(Rec!intYearID) Then
'                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
'                Else
'                    mstrYear = ""
'                End If
'                Select Case Rec!tnyPeriodID
'                    Case Is = 1: Print #gbFileNO, Tab(12); mstrYear & "/1Hf";
'                    Case Is = 2: Print #gbFileNO, Tab(12); mstrYear & "/2Hf";
'                    Case Is = 3: Print #gbFileNO, Tab(12); mstrYear & "/F";
'                    Case Else:   Print #gbFileNO, Tab(12); mstrYear;
'
'                End Select
'
'                If Rec!intYearID < gbFinancialYearID Then
'                    Print #gbFileNO, Tab(27); PadL(Format(Rec!fltAmount, "0.00"), 9);
'                Else
'                    Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
'                End If
'
'
'                '==================================================================='
'                ' Receipt Area
'                '==================================================================='
'                Print #gbFileNO, Tab(48); PadL(CStr(mLoop), 2);
'                Print #gbFileNO, Tab(56); PadR(Rec!vchAlias, 41);
'                If Not IsNull(Rec!intYearID) Then
'                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
'                Else
'                    mstrYear = ""
'                End If
'                Select Case Rec!tnyPeriodID
'                    Case Is = 1: Print #gbFileNO, Tab(98); mstrYear & "/1Hf";
'                    Case Is = 2: Print #gbFileNO, Tab(98); mstrYear & "/2Hf";
'                    Case Is = 3: Print #gbFileNO, Tab(98); mstrYear & "/F";
'                    Case Else:   Print #gbFileNO, Tab(98); mstrYear;
'                End Select
'
'                If Rec!intYearID < gbFinancialYearID Then
'                    Print #gbFileNO, Tab(109); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                Else
'                    Print #gbFileNO, Tab(126); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                End If
'                'Print #gbFileNO, Tab(26); PadL(Trim(str(mLoop)), 3); Tab(31); Rec!vchAccountHeadCode; Tab(40); PadR(IIf(IsNull(Rec!vchAlias), "", Rec!vchAlias), 20); Rec!tnyPeriodID; Tab(70); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                Rec.MoveNext
'            Wend
'            Rec.MoveFirst
'
'            For mCount = mLoop + 1 To 10
'                Print #gbFileNO,
'            Next mCount
'            Print #gbFileNO,
'
'            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
'            Print #gbFileNO, Tab(117); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
'
'            Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
'            Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
'            Print #gbFileNO,
'            If Not IsNull(Rec!numInwardNo) Then
'                mInward = Rec!numInwardNo
'            Else
'                mInward = ""
'            End If
'            Print #gbFileNO, mInward; ; Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 40); Tab(61); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
'
'            Print #gbFileNO,
'            objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                Print #gbFileNO, Tab(11); objCounter.CounterNo;
'                Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
'            End If
'            objUser.SetUser (Rec!intUserID)
'            If objUser.UserID > -1 Then
'                Print #gbFileNO, Tab(11); objUser.UserName;
'                Print #gbFileNO, Tab(61); objUser.UserName
'            End If
'        End If
'
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'
'        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
'
'        Close #gbFileNO
'        'ShellPad
'        Shell "Print " & gbFileName
'        'Kill gbFileName
'    End Sub
'
'
'


Private Sub cmd_Click()

End Sub

Private Sub cmdBudgetTest_Click()
    Dim objBudget As New clsBudgetCentre
    Dim Rec As New ADODB.Recordset
    objBudget.FunctionaryID = 2
    objBudget.FunctionID = 2
    objBudget.SetBudgetAccountHead 336
End Sub

Private Sub cmdDifferenceInVoucher_Click()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecTr As New ADODB.Recordset
    Dim mSQl As String
    Dim objDB As New clsDB
    Dim mInput As Variant
    Dim mOutput As Variant
    Dim objAc As New clsAccounts
    Dim mVoucherNo As Double
    Dim Recv As New ADODB.Recordset
    
    
    FileInitialize
    objDB.SetConnection mCnn
    
    Set RecTr = objDB.ExecuteSP("spTmpFetchVouchers", , , , mCnn, adCmdStoredProc)

    While Not RecTr.EOF
        objAc.SetAccountID (RecTr!intKeyID1)
        If objAc.AccountHeadID > 0 Then
            mInput = Array(objAc.AccountCode, 20, mVoucherNo)
            Set Recv = objDB.ExecuteSP("spTmpGetVoucherNo", mInput, mOutput, , mCnn, adCmdStoredProc)
            Print #gbFileNO, objAc.AccountCode;
            If IsArray(mOutput) Then
                Print #gbFileNO, mOutput(0, 0); "  "
                mSQl = "SELECT * FROM faVouchers Where tnyVoucherTypeID = 20 And intVoucherID = " & RecTr!intVoucherID
                Rec.Open mSQl, mCnn, adOpenKeyset, adLockPessimistic
                If Not (Rec.EOF And Rec.BOF) Then
                    Rec!intVoucherNo = mOutput(0, 0)
                    Rec.Update
                End If
                Rec.Close
                
            Else
                Print #gbFileNO,
            End If
        Else
            Print #gbFileNO, "Account Head Not Found "; RecTr!intKeyID1
        End If
        RecTr.MoveNext
    Wend
    
    
    RecTr.Close
    Close #gbFileNO
    ShellPad
    
End Sub

Private Sub cmdGetMac_Click()
    Dim mMac As String
    mMac = GetMacAddress
    MsgBox mMac
    Debug.Print mMac
    
    txtIP.Text = GetMeMacAddressOf(Trim(txtIP.Text))
    
End Sub

    Private Sub cmdInsertQRY_Click()
        On Error GoTo last
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSQl As String
        Dim mCount As Integer
        Dim mTotalStr As String
        Dim mData As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCount = 0
        txtInsertQRY.Text = ""
        mTotalStr = ""
        If Trim(txtTBL.Text) = "" Or txtQRY = "" Then
            MsgBox "Fill All Requred Fields Table Name & Query"
            Exit Sub
        End If
        Rec.Open txtQRY.Text, mCnn
        mSQl = "INSERT INTO " & txtTBL.Text & " ( "
        While mCount < Rec.Fields.count
            mSQl = mSQl & Rec.Fields(mCount).Name & ","
            mCount = mCount + 1
        Wend
        mSQl = Left(mSQl, Len(mSQl) - 1) & ") VALUES("
        
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                mCount = 0
                mData = ""
                For mCount = 0 To Rec.Fields.count - 1
                    If Left(Rec.Fields(mCount).Name, 2) = "dt" Then
                        mData = mData & "'" & Format(Rec.Fields(mCount), "dd/mmm/yyyy") & "',"
                    Else
                        mData = mData & "'" & Rec.Fields(mCount) & "',"
                    End If
                Next
                mData = Left(mData, Len(mData) - 1) & ")"
                mTotalStr = mTotalStr + mSQl + mData + vbNewLine
                Rec.MoveNext
            Wend
        End If
        txtInsertQRY.Text = mTotalStr
        Rec.Close
        Exit Sub
last:
        MsgBox err.Description, vbInformation
    End Sub

    Private Sub cmdIRR_Click()
        frmInterruptedReceiptRegister.Show
    End Sub

Private Sub cmdPrintReceipt_Click()
    'Call PrintReceipt_ForNewFormat(val(txtReceiptNo))
    'Call PrintReceipt(val(txtReceiptNo))
End Sub

Private Sub cmdPTAX_Click()
frmPTaxCalculator.Show
End Sub

Private Sub cmdSortVoucherNo_Click()
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecTran As New ADODB.Recordset
    
    Dim mSQl As String
    Dim varInPut As Variant
    Dim varOutPut As Variant
    
    mSQl = "Select * From faVouchers WHERE tnyVoucherTypeID = 20 Order By dtDate"
    objDB.SetConnection mCnn
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    FileInitialize
    While Not Rec.EOF
        mSQl = "Select * From faTransactions Where intGroupID = 20 AND intVoucherID = " & Rec!intVoucherID
        RecTran.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
        If RecTran.EOF And RecTran.BOF Then
            Print #gbFileNO, Rec!intVoucherID ', Rec!intVoucherNo, Rec!dtDate, Rec!vchInstrumentNo, Rec!fltAmount, Rec!vchDescription
        End If
        RecTran.Close
        Rec.MoveNext
    Wend
    Rec.Close
    Close #gbFileNO
    ShellPad
    
End Sub

Private Sub cmdTest_Click()
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSQl As String
    Dim mData As String
    Dim mTblName As String
    mTblName = InputBox("Enter the Name of Table", "To Generate Insert Script", "")
    
    objDB.SetConnection mCnn
    Rec.Open "SELECT column_name,data_type FROM information_schema.columns WHERE table_name = '" & mTblName & "'", mCnn, adOpenDynamic, adLockOptimistic
    If Not (Rec.BOF And Rec.EOF) Then
        
        FileInitialize
        mSQl = "INSERT INTO " & mTblName & " ( "
        While Not Rec.EOF
            mSQl = mSQl & Rec.Fields(0).value & ","
            Rec.MoveNext
        Wend
        mSQl = Left(mSQl, Len(mSQl) - 1) & ") VALUES("
        'Print #gbFileNO, mSQL
        
        Rec.Close
        Dim mCount As Integer
        
        Rec.Open mTblName, mCnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
        If Not (Rec.BOF And Rec.EOF) Then
            
            While Not Rec.EOF
                mCount = 0
                mData = ""
                For mCount = 0 To Rec.Fields.count - 1
                    If Left(Rec.Fields(mCount).Name, 2) = "dt" Then
                        mData = mData & "'" & Format(Rec.Fields(mCount), "dd/mmm/yyyy") & "',"
                    Else
                        mData = mData & "'" & Rec.Fields(mCount) & "',"
                    End If
                Next
                mData = Left(mData, Len(mData) - 1) & ")"
                Print #gbFileNO, mSQl + mData
                Rec.MoveNext
            Wend
        End If
        Close #gbFileNO
        ShellPad
    End If
    
    

    'objDB.SetConnection mCnn
    'Rec.Open "SELECT column_name,data_type FROM information_schema.columns WHERE table_name = '" & mTblName & "'", mCnn, adOpenDynamic, adLockOptimistic
    'If Not (Rec.BOF And Rec.EOF) Then
    '
    '    FileInitialize
    '    mSQL = "INSERT INTO " & mTblName & " ( "
    '    While Not Rec.EOF
    '        mSQL = mSQL & Rec.Fields(0).Value & ","
    '        Rec.MoveNext
    '    Wend
    '    mSQL = Left(mSQL, Len(mSQL) - 1) & ") VALUES("
    '
    '    Rec.MoveFirst
    '    While Not Rec.EOF
    '
    '        mSQL = mSQL & "''+ ISNULL(CONVERT(VarChar(20), " & Rec.Fields(0).Value & "),''Null'') +'',"
    '        Rec.MoveNext
    '    Wend
    '    mSQL = Left(mSQL, Len(mSQL) - Len("+'',")) & "+'' )'' FROM " & mTblName
    '    Print #gbFileNO, mSQL
    '
    '
    '
    '    Close #gbFileNO
    '    ShellPad
    'End If


End Sub

Private Sub Command1_Click()
   Dim Rec As New ADODB.Recordset
   Dim RecTr As New ADODB.Recordset
   Dim mCnn As New ADODB.Connection
   Dim objDB As New clsDB
   Dim mSQl As String
   Dim mLastTransactionID As Long
   Dim mNarration As String
   
   objDB.SetConnection mCnn
   'mSQL = "SELECT intVoucherID, dtDate, intVoucherNo, fltAmount, intTransactionTypeID, vchDescription  From faVouchers where tnyVoucherTypeID = 40 AND tnyCancelFlag <> 1 Order By  intVoucherID  "
   mSQl = "SELECT intVoucherID, dtDate, intVoucherNo, fltAmount, intTransactionTypeID, vchDescription  From faVouchers where tnyVoucherTypeID = 40  Order By  intVoucherID  "
   Rec.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic
   FileInitialize
   mLastTransactionID = 0
   While Not Rec.EOF
        Print #gbFileNO, Rec!intVoucherID, Rec!dtDate, Rec!intVoucherNo, Rec!fltAmount,
        mSQl = "Select faTransactions.intTransactionID, dtTransactionDate, intTransactionTypeID, intGroupID, fltAmount, faTransactions.vchNarration  From faTransactions Inner Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID WHERE intSerialNo = 1 AND intGroupID = 40 "
        mSQl = mSQl + " AND dtTransactionDate = '" & DdMmmYy(Rec!dtDate) & "' AND fltAmount = " & Rec!fltAmount & " AND faTransactions.intTransactionID > " & mLastTransactionID
        RecTr.Open mSQl, mCnn, adOpenForwardOnly, adLockOptimistic
        If Not (RecTr.BOF And RecTr.EOF) Then
            Print #gbFileNO, RecTr!intTransactionID, RecTr!fltAmount
            If Trim(Rec!vchDescription) <> Trim(RecTr!vchNarration) Then
                Print #gbFileNO, Rec!vchDescription
                Print #gbFileNO, RecTr!vchNarration
            Else
                mLastTransactionID = RecTr!intTransactionID
                mCnn.Execute "Update faTransactions Set intVoucherID = " & Rec!intVoucherID & " WHERE intTransactionID = " & RecTr!intTransactionID
            End If
        Else
            Print #gbFileNO, Rec!intVoucherID, Rec!dtDate, Rec!intVoucherNo, Rec!fltAmount,
            Print #gbFileNO, , , "Not found"
        End If
        RecTr.Close
        Rec.MoveNext
        
   Wend
   Close #gbFileNO
   ShellPad
End Sub

Private Sub Command10_Click()
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQl As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    mSQl = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "   OPENING BALANCE"
    Print #gbFileNO, "==================="
    While Not Rec.EOF
        If Rec!fltOpeningBalance <> 0 Then
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
            Print #gbFileNO, Tab(90); PadL(Format(Rec!fltOpeningBalance, "00"), 12)
            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "00")
            
        End If
        Rec.MoveNext
    Wend
    Rec.Close
    
    
    mSQl = " Select Distinct faAccountHeads.intAccountHeadID,faAccountHeads.intMajorAccountHeadID,"
    mSQl = mSQl + " faAccountHeads.intMinorAccountHeadID , faAccountHeads.vchAccountHeadCode ,"
    mSQl = mSQl + " faAccountHeads.vchAccountHead, vchMinorAccountHeadCode, vchMinorAccountHead,"
    mSQl = mSQl + " vchMajorAccountHeadCode , vchMajorAccountHead, intOperating "
    mSQl = mSQl + " From faTransactionChild Inner Join"
    mSQl = mSQl + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
    mSQl = mSQl + " faAccountHeads ON faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID Inner Join"
    mSQl = mSQl + " faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID Inner Join"
    mSQl = mSQl + " faMajorAccountHeads On faMajorAccountHeads.intMajorAccountHeadID = faAccountHeads.intMajorAccountHeadID"
    mSQl = mSQl + " Where faTransactions.intGroupID In (10,20) AND faAccountHeads.intMajorAccountHeadID <>40"
    mSQl = mSQl + " Order By intOperating, faAccountHeads.intMajorAccountHeadID, faAccountHeads.intMinorAccountHeadID,"
    mSQl = mSQl + " faAccountHeads.vchAccountHeadCode "
    
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   RECEIPTS"
    Print #gbFileNO, "======================================================"
    
    Print #gbFileNO,
    Print #gbFileNO, "  Operating Payments"
    Print #gbFileNO, "------------------------------------------------------"
    
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!CR <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                        Print #gbFileNO, "------------------------------------------------------"
                        Print #gbFileNO, "  Non-Operating Receipts"
                        Print #gbFileNO, "------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; Rec!vchAccountHead;
                Print #gbFileNO, Tab(90); PadL(Format(RecBalance!CR, "00"), 12)
                mTotalDr = mTotalDr + Format(RecBalance!CR, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.MoveFirst
    mFlag = False
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   PAYMENTS"
    Print #gbFileNO, "======================================================"
    
    Print #gbFileNO,
    Print #gbFileNO, "  Operating Payments"
    Print #gbFileNO, "------------------------------------------------------"
                    
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Dr <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                        Print #gbFileNO, "------------------------------------------------------"
                        Print #gbFileNO, "  Non-Operating Payments"
                        Print #gbFileNO, "------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; Rec!vchAccountHead;
                Print #gbFileNO, Tab(118); PadL(Format(RecBalance!Dr, "00"), 12)
                mTotalCr = mTotalCr + Format(RecBalance!Dr, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close
    
    mSQl = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   CLOSING BALANCE "
    Print #gbFileNO, "==================="
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Balance > 0 Then
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
            Print #gbFileNO, Tab(118); PadL(Format(RecBalance!Balance, "00"), 12)
            'mTotalDr = mTotalDr + Format(RecBalance!Balance, "00")
            ElseIf RecBalance!Balance < 0 Then
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
            Print #gbFileNO, Tab(90); PadL(Format(RecBalance!Balance, "00"), 12)
            'mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
            End If
            mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
        End If
        RecBalance.Close
        
        Rec.MoveNext
    Wend
    Rec.Close
    
    
    Print #gbFileNO, Tab(72); "============================================================"
    Print #gbFileNO, Tab(90); PadL(Format(mTotalDr, "00"), 12); Tab(118); PadL(Format(mTotalCr, "00"), 12);
    Print #gbFileNO, Tab(72); "============================================================"
    Close #gbFileNO
    ShellPad





















'
'    FileInitialize
'
'    objDB.SetConnection mCnn
'
'    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
'    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'
'    Print #gbFileNO, "   OPENING BALANCE"
'    Print #gbFileNO, "==================="
'    While Not Rec.EOF
'        If Rec!fltOpeningBalance <> 0 Then
'            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'            Print #gbFileNO, Tab(50); PadL(Format(Rec!fltOpeningBalance, "00"), 12)
'            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "00")
'
'        End If
'        Rec.MoveNext
'    Wend
'    Rec.Close
'
'
'    mSQL = " Select Distinct faAccountHeads.intAccountHeadID,faAccountHeads.intMajorAccountHeadID,"
'    mSQL = mSQL + " faAccountHeads.intMinorAccountHeadID , faAccountHeads.vchAccountHeadCode ,"
'    mSQL = mSQL + " faAccountHeads.vchAccountHead, vchMinorAccountHeadCode, vchMinorAccountHead,"
'    mSQL = mSQL + " vchMajorAccountHeadCode , vchMajorAccountHead, intOperating "
'    mSQL = mSQL + " From faTransactionChild Inner Join"
'    mSQL = mSQL + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
'    mSQL = mSQL + " faAccountHeads ON faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID Inner Join"
'    mSQL = mSQL + " faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID Inner Join"
'    mSQL = mSQL + " faMajorAccountHeads On faMajorAccountHeads.intMajorAccountHeadID = faAccountHeads.intMajorAccountHeadID"
'    mSQL = mSQL + " Where faTransactions.intGroupID In (10,20) AND faAccountHeads.intMajorAccountHeadID <>40"
'    mSQL = mSQL + " Order By intOperating, faAccountHeads.intMajorAccountHeadID, faAccountHeads.intMinorAccountHeadID,"
'    mSQL = mSQL + " faAccountHeads.vchAccountHeadCode "
'
'    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO, "   RECEIPTS"
'    Print #gbFileNO, "======================================================"
'
'    Print #gbFileNO,
'    Print #gbFileNO, "  Operating Payments"
'    Print #gbFileNO, "------------------------------------------------------"
'
'    While Not Rec.EOF
'        varInput = Array(Rec!intAccountHeadID)
'        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInput, , , mCnn, adCmdStoredProc)
'        If Not (RecBalance.EOF And RecBalance.BOF) Then
'            If RecBalance!Cr <> 0 Then
'                If Not mFlag Then
'                    If Rec!intOperating = 1 Then
'                        Print #gbFileNO, "------------------------------------------------------"
'                        Print #gbFileNO, "  Non-Operating Receipts"
'                        Print #gbFileNO, "------------------------------------------------------"
'                        mFlag = True
'                    End If
'                End If
'                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'                Print #gbFileNO, Tab(50); PadL(Format(RecBalance!Cr, "00"), 12)
'                mTotalDr = mTotalDr + Format(RecBalance!Cr, "00")
'            End If
'        End If
'        RecBalance.Close
'        Rec.MoveNext
'    Wend
'    Rec.MoveFirst
'    mFlag = False
'
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO, "   PAYMENTS"
'    Print #gbFileNO, "======================================================"
'
'    Print #gbFileNO,
'    Print #gbFileNO, "  Operating Payments"
'    Print #gbFileNO, "------------------------------------------------------"
'
'    While Not Rec.EOF
'        varInput = Array(Rec!intAccountHeadID)
'        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInput, , , mCnn, adCmdStoredProc)
'        If Not (RecBalance.EOF And RecBalance.BOF) Then
'            If RecBalance!Dr <> 0 Then
'                If Not mFlag Then
'                    If Rec!intOperating = 1 Then
'                        Print #gbFileNO, "------------------------------------------------------"
'                        Print #gbFileNO, "  Non-Operating Payments"
'                        Print #gbFileNO, "------------------------------------------------------"
'                        mFlag = True
'                    End If
'                End If
'                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'                Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Dr, "00"), 12)
'                mTotalCr = mTotalCr + Format(RecBalance!Dr, "00")
'            End If
'        End If
'        RecBalance.Close
'        Rec.MoveNext
'    Wend
'    Rec.Close
'
'    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
'    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO, "   CLOSING BALANCE "
'    Print #gbFileNO, "==================="
'    While Not Rec.EOF
'        varInput = Array(Rec!intAccountHeadID)
'        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInput, , , mCnn, adCmdStoredProc)
'        If Not (RecBalance.EOF And RecBalance.BOF) Then
'            If RecBalance!Balance > 0 Then
'            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'            Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Balance, "00"), 12)
'            'mTotalDr = mTotalDr + Format(RecBalance!Balance, "00")
'            ElseIf RecBalance!Balance < 0 Then
'            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'            Print #gbFileNO, Tab(50); PadL(Format(RecBalance!Balance, "00"), 12)
'            'mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
'            End If
'            mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
'        End If
'        RecBalance.Close
'
'        Rec.MoveNext
'    Wend
'    Rec.Close
'
'
'    Print #gbFileNO, Tab(50); "============================================================"
'    Print #gbFileNO, Tab(50); PadL(Format(mTotalDr, "00"), 12); Tab(74); PadL(Format(mTotalCr, "00"), 12);
'    Print #gbFileNO, Tab(50); "============================================================"
'    Close #gbFileNO
'    ShellPad
'

End Sub

Private Sub Command11_Click()
    
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQl As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    'mSQL = "Select * From faMinorAccountHeads where intMajorAccountHeadID = 40 Order By vchMinorAccountHeadCode"
    
    mSQl = "Select faAccountHeads.intGroupID, Sum(fltOpeningBalance) fltOpeningBalance From faAccountHeads Inner Join"
    mSQl = mSQl + " faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID"
    mSQl = mSQl + " Where faAccountHeads.intMajorAccountHeadID = 40"
    mSQl = mSQl + " Group By faAccountHeads.intGroupID"

    
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "   OPENING BALANCE   "
    Print #gbFileNO, "====================="
    While Not Rec.EOF
        If Rec!fltOpeningBalance <> 0 Then
            If Rec!intGroupID = 1 Then
                Print #gbFileNO, "  Cash";
            Else
                Print #gbFileNO, "  Bank";
            End If
            Print #gbFileNO, Tab(50); PadL(Format(Rec!fltOpeningBalance, "00"), 12)
            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "00")
        End If
        Rec.MoveNext
    Wend
    Rec.Close
    
    mSQl = "        Select Distinct faAccountHeads.intMajorAccountHeadID, vchMajorAccountHeadCode , "
    mSQl = mSQl + " vchMajorAccountHead,intOperating "
    mSQl = mSQl + " From faTransactionChild Inner Join"
    mSQl = mSQl + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
    mSQl = mSQl + " faAccountHeads ON faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID Inner Join"
    mSQl = mSQl + " faMajorAccountHeads On faMajorAccountHeads.intMajorAccountHeadID = faAccountHeads.intMajorAccountHeadID"
    mSQl = mSQl + " Where faTransactions.intGroupID In (10,20) AND faAccountHeads.intMajorAccountHeadID <> 40"
    mSQl = mSQl + " Order By intOperating, faAccountHeads.intMajorAccountHeadID"
    
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   RECEIPTS"
    Print #gbFileNO, '"======================================================"
    Print #gbFileNO, "  Operating Receipts"
    Print #gbFileNO, '"------------------------------------------------------"
    
    While Not Rec.EOF
        varInPut = Array(Rec!intMajorAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceMajorHeadWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!CR <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                        Print #gbFileNO, '"------------------------------------------------------"
                        Print #gbFileNO, "  Non-Operating Receipts"
                        Print #gbFileNO, '"------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                    
                'Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; PadL(Rec!vchMajorAccountHead, 30);
                Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; Rec!vchMajorAccountHead;
                Print #gbFileNO, Tab(50); PadL(Format(RecBalance!CR, "00"), 12)
                mTotalDr = mTotalDr + Format(RecBalance!CR, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.MoveFirst
    mFlag = False
    
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   PAYMENTS"
    Print #gbFileNO, '"======================================================"
    Print #gbFileNO, "     Operating Payments"
    Print #gbFileNO, '"------------------------------------------------------"
                    
    While Not Rec.EOF
        varInPut = Array(Rec!intMajorAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceMajorHeadWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Dr <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                    
                        Print #gbFileNO, '"------------------------------------------------------"
                        Print #gbFileNO, "     Non-Operating Payments"
                        Print #gbFileNO, ' "------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                    
                'Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; PadL(Rec!vchMajorAccountHead, 30);
                Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; Rec!vchMajorAccountHead;
                Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Dr, "00"), 12)
                mTotalCr = mTotalCr + Format(RecBalance!Dr, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close
    
    Dim mClosingCash As Double
    Dim mClosingBank As Double
    
    mSQl = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By intGroupID,vchAccountHeadCode"
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   CLOSING BALANCE "
    Print #gbFileNO, '"==================="
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If Rec!intGroupID = 1 Then
                mClosingCash = mClosingCash + RecBalance!Balance
            ElseIf Rec!intGroupID = 2 Then
                mClosingBank = mClosingBank + RecBalance!Balance
            End If
            mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close
    
    Print #gbFileNO, " Cash";
    Print #gbFileNO, Tab(74); PadL(Format(mClosingCash, "00"), 12)
    
    Print #gbFileNO, " Bank";
    Print #gbFileNO, Tab(74); PadL(Format(mClosingBank, "00"), 12)

    
    Print #gbFileNO, Tab(50); "============================================================"
    Print #gbFileNO, Tab(50); PadL(Format(mTotalDr, "00"), 12); Tab(74); PadL(Format(mTotalCr, "00"), 12);
    Print #gbFileNO, Tab(50); "============================================================"
    Close #gbFileNO
    ShellPad

    
End Sub

Private Sub Command12_Click()
    
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQl As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    mSQl = "Select * From faMajorAccountHeads"
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "=================================================================================="
    Print #gbFileNO, "   Major Account Heads"
    Print #gbFileNO, "=================================================================================="
    While Not Rec.EOF
        Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; "RP-"; Trim(str(Rec!intMajorAccountHeadID)); Tab(18); Rec!vchMajorAccountHead
        Rec.MoveNext
    Wend
    Rec.Close
    Close #gbFileNO
    ShellPad
    
End Sub

Private Sub Command13_Click()
        Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecTr As New ADODB.Recordset
    Dim mSQl As String
    Dim objDB As New clsDB
    Dim mInput As Variant
    Dim mOutput As Variant
    Dim objAc As New clsAccounts
    Dim mVoucherNo As Double
    Dim Recv As New ADODB.Recordset
    
    
    FileInitialize
    objDB.SetConnection mCnn
    Set RecTr = objDB.ExecuteSP("spTmpFetchVouchers", , , , mCnn, adCmdStoredProc)
    
    
    While Not RecTr.EOF
        mInput = Array((RecTr!intCounterID), (RecTr!intKeyID1), 1)
        objDB.ExecuteSP "spTmpGetBookNo", mInput, mOutput, , mCnn, adCmdStoredProc
        If IsArray(mOutput) Then
            
            Print #gbFileNO, mOutput(0, 0); "  "; mOutput(1, 0),
            mSQl = "SELECT * FROM faVouchers Where tnyVoucherTypeID = 10 And tnyStatus = 1 AND intVoucherID = " & RecTr!intVoucherID
            Rec.Open mSQl, mCnn, adOpenKeyset, adLockPessimistic
            If Not (Rec.EOF And Rec.BOF) Then
                Rec!intBookNo = mOutput(0, 0)
                Rec!intVoucherNo = mOutput(1, 0)
                Rec.Update
            Else
                Print #gbFileNO, "Not found VoucherID "
            End If
            Rec.Close
        Else
            Print #gbFileNO, "Book No not generated!"
        End If
        RecTr.MoveNext
    Wend
    
    
    
    RecTr.Close
    Close #gbFileNO
    ShellPad
    

    
    
End Sub



Private Sub Command14_Click()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecTr As New ADODB.Recordset
    Dim RecVr As New ADODB.Recordset
    
    Dim mSQl As String
    Dim mTotalAmt As Double
    Dim mVoucherTotal As Double
    
    FileInitialize
    objDB.SetConnection mCnn
    mSQl = "Select * From faVouchers Where tnyCancelFlag = 0 AND dtDate = '" & txtDate & "'"
    Rec.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic
    If Not (Rec.BOF And Rec.EOF) Then
        While Not Rec.EOF
            GoTo Skip
            mSQl = "Select * From faTransactionChild Inner Join faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID Where intVOucherID = " & Rec!intVoucherID
            RecTr.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic
            If Not (RecTr.BOF And RecTr.EOF) Then
                mTotalAmt = 0
                While Not RecTr.EOF
                    '----------------------------'
                    ' Checking KeyID
                    '----------------------------'
                    If RecTr!intSerialNo = 1 Then
                        If Rec!intKeyID1 <> RecTr!intAccountHeadID Then
                            'Print #gbFileNO, "KeyID NOT matching";
                        Else
                            'Print #gbFileNO, "                  ";
                        End If
                    Else
                    
                    '----------------------------'
                    '
                    '----------------------------'
                    If RecTr!intAccountHeadID <> Rec!intAccountHeadID Then
                        'Print #gbFileNO, "VID : " & Rec!intVoucherID, Rec!intAccountHeadID; Rec!fltAmount, " Missing Head"
                    End If
                    mTotalAmt = mTotalAmt + RecTr!fltAmount
                    End If ' intSerialNo
                    RecTr.MoveNext
                    If RecTr.EOF Then
                        If Rec!TotalAmount <> mTotalAmt Then
                            Print #gbFileNO, "VID = "; Rec!intVoucherID, Rec!TotalAmount; "="; "Total Amount Not Matching  : "; mTotalAmt
                        End If
                    End If
                    
                Wend
            End If
            RecTr.Close
Skip:
            mSQl = "Select * From faVoucherChild Where intVoucherID = " & Rec!intVoucherID
            RecVr.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic
            mVoucherTotal = 0
            If Not (RecVr.BOF And RecVr.EOF) Then
                While Not RecVr.EOF
                    mVoucherTotal = mVoucherTotal + RecVr!fltAmount
                    RecVr.MoveNext
                Wend
            End If
            RecVr.Close
            If Rec!fltAmount <> (mVoucherTotal - Rec!fltAdvAmtAdj + Rec!fltRoundOff) Then
                Print #gbFileNO, "VID = "; Rec!intVoucherID, Rec!fltAmount; "="; "Total Amount Not Matching  : "; mVoucherTotal
            End If
            
            Rec.MoveNext
        Wend
        
        
        
    End If
    Close #gbFileNO
    ShellPad
    
    
End Sub

Private Sub Command15_Click()
       Dim Rec As New ADODB.Recordset
   Dim RecTr As New ADODB.Recordset
   Dim mCnn As New ADODB.Connection
   Dim objDB As New clsDB
   Dim mSQl As String
   Dim mLastTransactionID As Integer
   Dim mNarration As String
   
   objDB.SetConnection mCnn
   'mSQL = "SELECT * From faTransactions where intGroupID = 10  Order By  intTransactionID  "
   mSQl = "Select faTransactions.intTransactionID, dtTransactionDate, intTransactionTypeID, intGroupID, fltAmount, faTransactions.vchNarration  From faTransactions Inner Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID WHERE intSerialNo = 1 AND intGroupID = 10  ORDER By faTransactions.intTransactionID"
   Rec.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic
   FileInitialize
   mLastTransactionID = 0
   While Not Rec.EOF
        Print #gbFileNO, Rec!intTransactionID, Rec!dtTransactionDate, Rec!intVoucherID, Rec!fltAmount,
        mSQl = "Select *  From faVouchers Inner Join faVoucherChild On faVoucherChild.intVoucherID = faVouchers.intVoucherID WHERE  tnyVoucherTypeID = 10 "
        mSQl = mSQl + " AND dtDate = '" & DdMmmYy(Rec!dtTransactionDate) & "' AND faVouchers.fltAmount = " & Rec!fltAmount & " AND faVouchers.intVoucherID > " & mLastTransactionID
        RecTr.Open mSQl, mCnn, adOpenForwardOnly, adLockOptimistic
        If Not (RecTr.BOF And RecTr.EOF) Then
            'Print #gbFileNO, RecTr!intVoucherID, RecTr!fltAmount
            If Trim(RecTr!vchDescription) <> Trim(Rec!vchNarration) Then
                Print #gbFileNO, RecTr!vchDescription
                Print #gbFileNO, Rec!vchNarration
            Else
                
                
            End If
            mLastTransactionID = RecTr!intVoucherID
            'mCnn.Execute "Update faTransactions Set intVoucherID = " & Rec!intVoucherID & " WHERE intTransactionID = " & RecTr!intTransactionID
            
        Else
            Print #gbFileNO, Rec!intTransactionID, Rec!dtTransactionDate, Rec!intVoucherID, Rec!fltAmount,
            Print #gbFileNO, , , "Not found"
        End If
        RecTr.Close
        Rec.MoveNext
        
   Wend
   Close #gbFileNO
   ShellPad

End Sub

Private Sub Command16_Click()

    Dim strRegConStr As String
    Dim dbCon As Object
    Dim objcn As Object 'Web32CR.clsGen1
    Set objcn = CreateObject("Web32CR.clsGen1")
    strRegConStr = objcn.gen_cnset("DSNSevanaRegn", 0)

    dbCon.Open strRegConStr

 

End Sub

    Private Sub Command17_Click()
        Dim mRowCnt As Integer
        Dim voucher As uVoucher
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim aryIn As Variant
        Dim mCounterID As Integer
        mCounterID = val(InputBox("Counter ID ", , "1"))
        On Error GoTo err:
        
        For mRowCnt = 1 To 50
            With voucher
                .dtDate_8 = Date
                .dtInstrumentDate_12 = Date
                .fltAdvAmtAdj = 0
                .fltAmount_9 = mRowCnt
                .fltRoundOff = 0
                .intBookNo_7 = 0
                .intCounterID_20 = mCounterID
                .intDoorNoP1_16 = 1
                .intExternalApplicationID_24 = 1
                .intExternalModuleID_25 = 1
                .intFinancialYearID_26 = 2009
                .intFundID_35 = 1
                .intInstrumentTypeID_10 = 1
                .intKeyID1_22 = 1
                .intKeyID2_23 = 1
                .intLocalBodyID_2 = 167
                .intSessionID = 1
                .intTransactionID_3 = 1
                .intTransactionTypeID_4 = 9999
                .intUserID_19 = 1
                .intVoucherID_1 = -1
                .intVoucherNo_6 = Null
                .numInwardNo = mRowCnt
                .numLocationID = 1
                .numSeatID = 1
                .numSubLedgerID_21 = 1
                .numWardID_15 = 1
                .numZoneID_14 = 1
                .tnyCancelFlag_29 = 0
                .tnyPrintFlag_28 = 0
                .tnyShiftID_27 = 0
                .tnyStatus_32 = 0
                .tnyVoucherTypeID_5 = 10
                .vchBank_33 = Null
                .vchBankPlace_34 = Null
                .vchDescription_13 = Null
                .vchDoorNoP2_17 = Null
                .vchDoorNoP3_18 = Null
                .vchInstrumentNo_11 = Null
                .vchRefNo = Null
                
                objDB.SetConnection mCnn
                
                aryIn = Array(.intVoucherID_1, _
                                .intLocalBodyID_2, _
                                .intTransactionID_3, _
                                .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
                                .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
                                .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
                                .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
                                .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
                                .intKeyID2_23, .intExternalApplicationID_24, _
                                .intExternalModuleID_25, .intFinancialYearID_26, _
                                .tnyShiftID_27, .tnyPrintFlag_28, _
                                .tnyCancelFlag_29, .vchBank_33, _
                                .vchBankPlace_34, .intFundID_35, _
                                .numSeatID, .intSessionID, _
                                .vchRefNo, .fltRoundOff, _
                                .fltAdvAmtAdj, .numInwardNo, _
                                .tnyStatus_32, .numLocationID)
                                
                objDB.ExecuteSP "spSaveVoucher", aryIn, , , mCnn, adCmdStoredProc
                
            End With
        Next
        MsgBox "Testing Success machaaaaa SUCCESS!!!! Hands off Mr. AIBY MOHANDAS! U R GREAT", vbInformation
        Exit Sub
        
err:
        MsgBox "Error" & Error$, vbInformation
    End Sub

Private Sub Command18_Click()
        ''Dim mCnn As New ADODB.Connection
        ''Dim Rec As New ADODB.Recordset
        ''Dim mSQL As String
        ''Dim objDb As New clsDB
        ''Dim i As Integer
        ''Dim J As Integer
        ''Dim mStr As String
        ''
        ''objDb.SetConnection mCnn
        ''
        ''For i = 1000 To 1500
        ''    mSQL = "Select * From faTransactionTypeChild  "
        ''    mSQL = mSQL + " Where intTransactionTypeID = " & i & " and intOrder = 0 "
        ''    mSQL = mSQL + " Order By intTransactionTypeID,intOrder "
        ''    J = 1
        ''    Rec.Open mSQL, mCnn
        ''    While Not (Rec.EOF Or Rec.BOF)
        ''        mStr = "Update faTransactionTypeChild Set intOrder = " & J & " Where intID = " & Rec!intID
        ''        mCnn.Execute mSQL
        ''        J = J + 1
        ''        Rec.MoveNext
        ''    Wend
        ''    If Rec.State = 1 Then Rec.Close
        ''Next
        ''MsgBox "Test"
End Sub

Private Sub Command2_Click()
    
    frmSugSaleofTender.Visible = True
End Sub

Private Sub Command3_Click()
        Dim mArrIn As Variant
        Dim mArrOut As Variant
        Dim mUrl   As String
        Dim client1 As New MSSOAPLib.SoapClient
        Dim objSOAP As Variant
        Dim clnt As New SoapClient30
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQl  As String
        Dim mReqIDCheck As Boolean
     
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
        '--------------'
        ' Web Service  '
        '--------------'
        mReqIDCheck = False
        mUrl = gbDefaultUrlForRequisition
        mArrIn = Array(gbLBID, gbFinancialYearID)
        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        objSOAP.mssoapinit mUrl + "?WSDL"

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

Private Sub Command5_Click()

    Call MenuManager
    
End Sub

Private Sub Command6_Click()
    Dim mSQl As String
    Dim Recv As New ADODB.Recordset
    Dim Rect As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mTotal As Single
    Dim mTotalAmt As Single
    Dim mID As Long
    
    FileInitialize
        Print #gbFileNO,
        mSQl = " Select faTransactionChild.* From faTransactionChild Inner Join "
        mSQl = mSQl + " faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID"
        mSQl = mSQl + " Where intGroupID = 40 Order By faTransactionChild.intTransactionID,intSerialNo "
        objDB.SetConnection mCnn
        Rect.Open mSQl, mCnn, adOpenDynamic, adLockOptimistic
        While Not Rect.EOF
            If mID <> Rect!intTransactionID Then
                If Int(mTotal) <> Int(mTotalAmt) Then
                    Print #gbFileNO, mID, mTotal, mTotalAmt, (mTotalAmt - mTotal)
                End If
                mID = Rect!intTransactionID
                mTotal = 0
                If Rect!intSerialNo = 1 Then
                    mTotalAmt = Format(Rect!fltAmount, "0.00")
                End If
            End If
            If Rect!intSerialNo <> 1 Then
                mTotal = mTotal + Format(Rect!fltAmount, "0.00")
            End If
            Rect.MoveNext
        Wend
    Close #gbFileNO
    ShellPad
    
End Sub


Private Sub Command7_Click()
    MsgBox RoundOffAdjustment(InputBox("Amount"))
End Sub

Private Sub Command8_Click()

            Dim rptFileName As String
            Dim arrInput As Variant
            Set arrInput = Nothing
            Dim Rpt As New CRAXDRT.Report
            Dim App As New CRAXDRT.Application
            Dim mLoop As Long
            
        
            rptFileName = "D:\My Projects\IKM\SaankhyaDoubleEntry\SourceCode\Reports\rptJournal.rpt"
           
            
'            If IsArray(mvarInputParameters) Then
'                arrInput = mvarInputParameters
'            End If

            Screen.MousePointer = vbHourglass
            CRV.Left = 0
            CRV.Top = 0
            CRV.DisplayToolbar = True
            CRV.Zoom 1
            CRV.Height = 5500
            CRV.Width = Me.Width
            CRV.EnableExportButton = True
            Set Rpt = Nothing
            App.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
            Set Rpt = App.OpenReport(rptFileName, 1)
            CRV.Container = Frame1
            CRV.ReportSource = Rpt
            CRV.Refresh
            CRV.ViewReport
            
            
End Sub

Private Sub Command9_Click()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQl As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    mSQl = "Select * from faAccountHeads Where intMajorAccountHeadID = 40"
    objDB.SetConnection mCnn
    Rec.Open mSQl, mCnn, adOpenKeyset, adLockOptimistic
    FileInitialize
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            
            
            While Not RecBalance.EOF
                Print #gbFileNO, RecBalance.Fields(0).value;
                Print #gbFileNO, Tab(10); RecBalance.Fields(1).value;
                Print #gbFileNO, Tab(25); RecBalance.Fields(2).value;
                Print #gbFileNO, Tab(40); RecBalance.Fields(3).value;
                RecBalance.MoveNext
            Wend
            Print #gbFileNO, Tab(58); "@"; Rec!fltOpeningBalance; " - "; Rec!vchAccountHeadCode; " "; Rec!vchAccountHead
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Close #gbFileNO
    ShellPad
End Sub
'Private Function CalculateFineforPTaxKLM(mYearID As Integer, mPeriodID As Integer, mPTax As Double) As Double
'        '==============================================================================='
'        ' Modified By : Aiby                                                            '
'        '             : For Kollam  Corporation                                         '
'        '                                                                               '
'        '==============================================================================='
'        Dim dtFromDt As Variant
'        Dim mNoOfMonths As Long
'        Dim mAmount     As Double
'        mFineAmt = 0
'        dtFromDt = DateSerial(mYearID, 11, 1)
'        If mYearID = gbFinancialYearID Then
'            CalculateFineforPTaxKLM = 0
'            Exit Function
'        End If
'
'        If mYearID < 2006 Then
'            If mYearID = 2005 Then
'                If mPeriodID > 1 Then
'                    GoTo Skip:
'                End If
'            End If
'            mNoOfMonths = (2005 - mYearID) * 24 + 10
'            If mPeriodID = 2 Then
'                mNoOfMonths = mNoOfMonths - 12
'            End If
'            'mNoOfMonths = mNoOfMonths + (gbFinancialYearID - 2005) * 12
'
'            mYearID = 2005
'            mPeriodID = 2
'        End If
'Skip:
'        mNoOfMonths = mNoOfMonths + (gbFinancialYearID - mYearID) * 12
'        If mPeriodID = 2 Then
'            mNoOfMonths = mNoOfMonths - 6
'        End If
'        If Month(gbTransactionDate) > 3 Then
'            mNoOfMonths = mNoOfMonths + Month(gbTransactionDate) - 3
'        Else
'            mNoOfMonths = mNoOfMonths + Month(gbTransactionDate) + 9
'        End If
'        mNoOfMonths = mNoOfMonths - 1
'        CalculateFineforPTaxKLM = mNoOfMonths
'    End Function

Private Sub Text1_LostFocus()
    Text1.Text = CheckDateInMMM(Text1.Text)
End Sub

Private Sub smdSubTotal()
'
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

Private Function CalculateFineforPTax1(mYearID As Integer, mPeriodID As Integer, mPTax As Double) As Double
'        '==============================================================================='
'        ' Modified By : Aiby                                                            '
'        '             : Befor Changing Fine Calculation for Calicut Corporation         '
'        '                                                                               '
'        '==============================================================================='
'        Dim dtFromDt As Variant
'        Dim mNoOfMonths As Integer
'        Dim mAmount     As Double
'
'        dtFromDt = DateSerial(mYearID, 11, 1)
'        If mYearID < 2004 Then
'        'If dtFromDt < DateSerial(2005, 8, 1) Then
'            mNoOfMonths = DateDiff("m", dtFromDt, DateSerial(2005, 8, 1))
'            mAmount = (mPTax * 2 / 100) * mNoOfMonths
'            mFineAmt = mAmount + mFineAmt
'            dtFromDt = DateSerial(2005, 8, 1)
'            GoTo CalculateRest:
'        Else
'CalculateRest:
'            mNoOfMonths = DateDiff("m", dtFromDt, gbTransactionDate) + 1
'            mAmount = mFineAmt + (mPTax * 1 / 100) * mNoOfMonths
'            If mAmount < 0 Then mAmount = mAmount * -1
'            mFineAmt = mAmount + mFineAmt
'        End If
'        CalculateFineforPTax1 = mFineAmt
    End Function
    
Private Sub txtDate_LostFocus()
    If Trim(txtDate.Text) <> "" Then
        txtDate.Text = CheckDateInMMM(txtDate.Text)
    Else
        txtDate.Text = ""
    End If
End Sub

'
'Private Sub txtPast_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Or KeyAscii = 45 Then
'            PressTabKey
'            Exit Sub
'        End If
'        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
'        Else
'            KeyAscii = 0
'        End If
'
'End Sub
Private Sub txtPast_LostFocus()
    Dim Total As Double
    Dim mTaxRate As Single
    mTaxRate = val(txtPast.Text)
    If (mTaxRate - Int(mTaxRate)) > 0 Then
            mTaxRate = Int(mTaxRate) + 1
          '  mHYrTaxInFraction = True
        End If
  
        txtPast.Text = Format(mTaxRate, "0.00")
        'txtTaxRate.Text = Format(val(txtTaxRate), "#0")
        txtPast.Text = Format(val(txtPast), "0.00")
        
End Sub
