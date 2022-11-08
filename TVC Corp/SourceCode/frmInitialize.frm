VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInitialize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "frmInitialize"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6840
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar PgrBar 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblInitialize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Initializing..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mTimer  As Integer
    Dim mYearID As Integer
    Dim mOpeningYearID As Integer
    Dim mExtractionState As Boolean
    Dim mControlVariableForProgressBar As Integer
    Dim mStartFlag As Boolean
    
    Private Sub Command1_Click()
        Timer1.Enabled = True

    End Sub
    
    Private Sub Form_Activate()
        ''Me.Top = 700
        ''Me.Left = frmMenu.Width - 7000

    End Sub

    Private Sub Form_DblClick()
        Unload Me
    End Sub

    Private Sub Form_Load()
    '        If gbCounterSectionID <> gbJSKSectionID Then
    '            PgrBar.value = 0
    '            lblInitialize.Visible = False
    '            PgrBar.Max = 100
    '            mExtractionState = False
    '            Call ExtractData
    '        End If
    End Sub
    Private Sub DoExtraction()
            PgrBar.value = 0
            lblInitialize.Visible = False
            PgrBar.Max = 100
            mExtractionState = False
            Call ExtractData
       End Sub
    Private Sub CheckProgressBar()
        PgrBar.Max = 10000 + 1
        While PgrBar.value < PgrBar.Max
            PgrBar.value = PgrBar.value + 1
        Wend
     End Sub
    
    Private Sub Timer1_Timer()
        If mStartFlag = False And mControlVariableForProgressBar > 5 Then
            mStartFlag = True
            Call DoExtraction
        End If
       
        If mTimer = 0 Then
            mTimer = 1
            lblInitialize.Visible = True
            'Exit Sub
        ElseIf mTimer = 1 Then
            mTimer = 0
            lblInitialize.Visible = False
            'Exit Sub
        End If
        
        If mExtractionState Then
            Timer1.Enabled = False
            Unload Me
        End If
        
        If mControlVariableForProgressBar < 20 Then
            mControlVariableForProgressBar = mControlVariableForProgressBar + 1
        Else
            mControlVariableForProgressBar = 0
            If PgrBar.value < PgrBar.Max Then
                PgrBar.value = PgrBar.value + 1
            Else
                PgrBar.value = 1
            End If
        End If
'        If PgrBar.value < PgrBar.Max Then
'                PgrBar.value = PgrBar.value + 1
'            Else
'                PgrBar.value = 1
'        End If
    
    End Sub
    
    Private Function ExtractData()
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim msql                As String
        Dim mRowCnt             As Integer
        Dim mArrIn              As Variant
        Dim mExtractedYearID    As Integer
    
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCnn.CommandTimeout = 1000000000
        
        'MARK AS SESSION STARTED - SO NO OTHER INSTANCE OF APP WILL RUN THE SAME ROUTINE
        Call UpdateLastExtractedDate
        
        'GET LAST EXTRACTED YEAR FROM CONFIG TABLE
        mExtractedYearID = GetExtractedYearID
        
        
        'NOTE:IF EXTRACTION IS NOT STARTED
        '     FIND OPENING FINANCIAL YEAR FROM OPENING JV
        If mExtractedYearID = 0 Then
            Call GetOpeningFinYear
            If mOpeningYearID > 2005 Then
                mExtractedYearID = mOpeningYearID
            End If
        End If
        
        
        'START YEARLY EXTRACTION
        While mExtractedYearID < 2014 And mExtractedYearID <> 0
            
            mYearID = mExtractedYearID + 1
            msql = " DELETE FROM faDailyExtracts WHERE  intFinancialYearID=" & mYearID & ""
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            
            msql = ""
            msql = "UPDATE faDailyIndex SET tnyExtractFlag=NULL, tnySyncFlag=NULL WHERE intFinYearID=" & mYearID & " "
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            
            mArrIn = Array(mYearID)
            objDb.ExecuteSP "spDailyExtract_Opening", mArrIn, , , mCnn, adCmdStoredProc
            objDb.ExecuteSP "spDailyExtractByYear", mArrIn, , , mCnn, adCmdStoredProc
            
            mExtractedYearID = mExtractedYearID + 1
            msql = " UPDATE faConfig SET ExtractedYearID = " & mExtractedYearID
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            If PgrBar.value < PgrBar.Max - 2 Then
                PgrBar.value = PgrBar.Max - 1
            End If
        Wend
        mExtractionState = True
        mCnn.Close
        
        
    End Function
    
    Private Function GetOpeningFinYear()    'FUNCTION TO GET THE OPENING YEAR ID

    Dim mCnn                As New ADODB.Connection
    Dim objDb               As New clsDB
    Dim Rec                 As New ADODB.Recordset
    Dim msql                As String

   
    objDb.SetConnection mCnn
    
    msql = " SELECT   dtDate, intFinancialYearID mYear FROM faVouchers  WHERE intTransactionTypeID=3000"
    
    Rec.Open msql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        mYearID = Rec!mYear
        mOpeningYearID = Rec!mYear
    End If
    Rec.Close
    'mYearID = mYearID + 1
   
    End Function

    Private Function UpdateLastExtractedDate()     'FUNCTION TO GET dtLastExtractedDate
    
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim msql                As String
        Dim dtLastExtractedDate     As Variant
    
        objDb.SetConnection mCnn
        msql = " UPDATE faConfig SET dtLastExtractedDate=GETDATE()"
        objDb.ExecuteSP msql, , , , mCnn, adCmdText
    End Function
    Private Function GetLastExtractedDate()     'FUNCTION TO GET dtLastExtractedDate
    
    Dim mCnn                As New ADODB.Connection
    Dim objDb               As New clsDB
    Dim Rec                 As New ADODB.Recordset
    Dim msql                As String
    Dim dtLastExtractedDate     As Variant
    
    objDb.SetConnection mCnn
    
    msql = " SELECT dtLastExtractedDate   FROM faConfig"
    Rec.Open msql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        If Rec!dtLastExtractedDate <> "" Then
            dtLastExtractedDate = DdMmYy(Rec!dtLastExtractedDate)
        Else
            dtLastExtractedDate = ""
        End If
    Rec.Close
    If dtLastExtractedDate = "" Then
        msql = " UPDATE faConfig SET dtLastExtractedDate=GETDATE()"
        objDb.ExecuteSP msql, , , , mCnn, adCmdText
    End If
    End If
    End Function

        
    Private Function GetExtractedYearID() As Integer    'FUNCTION TO GET ExtractedYearID
    
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim msql                As String
        Dim mExtractedYearID    As Variant
            
        objDb.SetConnection mCnn
        msql = " SELECT ExtractedYearID   FROM faConfig"
        Rec.Open msql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mExtractedYearID = IIf(IsNull(Rec!ExtractedYearID), 0, Rec!ExtractedYearID)
        End If
        Rec.Close
        GetExtractedYearID = mExtractedYearID
    End Function
