VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRefreshDCB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEMAND COLLECTION BALANCE"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRefreshDCB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9960
      TabIndex        =   12
      Top             =   6540
      Width           =   1140
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10035
      TabIndex        =   10
      Top             =   855
      Width           =   1140
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11265
      TabIndex        =   9
      Top             =   -45
      Width           =   11265
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   45
      TabIndex        =   2
      Top             =   675
      Width           =   11190
      Begin VB.ComboBox cmbFinancialYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   2085
      End
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label Label3 
         Caption         =   "FINANCIAL YEAR"
         Height          =   225
         Left            =   405
         TabIndex        =   7
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
         Height          =   225
         Left            =   4320
         TabIndex        =   4
         Top             =   270
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   11280
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   45
         ScaleHeight     =   15
         ScaleWidth      =   10410
         TabIndex        =   8
         Top             =   45
         Width           =   10410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   75
         TabIndex        =   1
         Top             =   210
         Width           =   75
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5040
      Left            =   45
      TabIndex        =   5
      Top             =   1395
      Width           =   11625
      _cx             =   20505
      _cy             =   8890
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRefreshDCB.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Editable        =   2
      ShowComboButton =   0
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
   Begin MSComctlLib.ProgressBar PgrBar 
      Height          =   255
      Left            =   3060
      TabIndex        =   11
      Top             =   6615
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmRefreshDCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mExtractionState As Boolean

    Private Sub FillCombo()
        Dim mCnn  As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select  intFinancialYear,intFinancialYearID from faFinancialYear where intFinancialYear between 2013 and 2014"
        PopulateList cmbFinancialYear, mSql, , , True, True, enuSourceString.Saankhya
    End Sub
    Private Sub fillMonthCombo()
        cmbMonth.AddItem "April"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 4
        cmbMonth.AddItem "May"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 5
        cmbMonth.AddItem "June"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 6
        cmbMonth.AddItem "July"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 7
        cmbMonth.AddItem "August"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 8
        cmbMonth.AddItem "September"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 9
        cmbMonth.AddItem "October"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 10
        cmbMonth.AddItem "November"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 11
        cmbMonth.AddItem "December"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 12
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
    End Sub

    Private Sub cmbMonth_Click()
        If cmbFinancialYear.ListIndex < 0 Then
            MsgBox "Select FinancialYear", vbInformation
            cmbMonth.ListIndex = -1
            Exit Sub
        Else
            Call FillGrid
            cmdVerify.Enabled = True
        End If
    End Sub

    Private Sub cmdPrint_Click()
        frmDCBReport.Show
        Unload Me
    End Sub

    Private Sub cmdVerify_Click()
        Dim mCnn    As New ADODB.Connection
        Dim mSql    As String
        Dim objDb   As New clsDB
        Dim arInput As Variant

        If cmbFinancialYear.ListIndex < 0 Then
            MsgBox "Year Not Selected", vbInformation
            Exit Sub
        End If
        If cmbMonth.ListIndex < 0 Then
            MsgBox "Month Not Selected", vbInformation
            Exit Sub
        End If
        
        arInput = Array(cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex), cmbMonth.ItemData(cmbMonth.ListIndex))
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCnn.Execute " Delete  from faMonthlyDCB Where intYearID =" & val(cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)) & " And intMonthID = " & cmbMonth.ItemData(cmbMonth.ListIndex)
        'objDb.ExecuteSP "spExtractMonthlyDCBNEW", arInput, , , mCnn, adCmdStoredProc
        
        objDb.ExecuteSP "spExtactDCBHeadWise", arInput, , , mCnn, adCmdStoredProc
       
        MsgBox "Refreshed!!!", vbInformation
        cmdVerify.Enabled = False
        mCnn.Close
        Call FillGrid
        
    End Sub
    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
    End Sub
    Private Sub Form_Load()
        Call FillCombo
        Call fillMonthCombo
        Call ExtractDCB
    End Sub
    Private Sub CheckProgressBar()
        PgrBar.Max = 10000 + 1
        While PgrBar.value < PgrBar.Max
            PgrBar.value = PgrBar.value + 1
        Wend
    End Sub
    Private Sub FillGrid()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objDb   As New clsDB
        Dim mLoop   As Integer
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya

          mSql = "SELECT  * ,numCurrentDemand+numArrearBalance AmountAccured FROM  faMonthlyDCB"
          mSql = mSql + " INNER JOIN faDCBHeads ON faMonthlyDCB.intDCBID=faDCBHeads.intID"
          mSql = mSql + " INNER JOIN faAccountHeads ON  faAccountHeads.intAccountHeadID=faMonthlyDCB.intDCBHeadID"
          
         If cmbFinancialYear.ListIndex > -1 Then
            mSql = mSql + " where  intYearID=  " & cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex) & ""

         End If

            mSql = mSql + " AND  intMonthID=  " & cmbMonth.ItemData(cmbMonth.ListIndex) & ""

        Rec.Open mSql, mCnn
        vsGrid.Rows = 1
        If Not (Rec.BOF And Rec.EOF) Then
            While Not (Rec.EOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!vchAccountHead
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!numPreYearAmt), "", Rec!numPreYearAmt)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!AmountAccured), "", Rec!AmountAccured)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!numCollectionPreMonth), "", Rec!numCollectionPreMonth)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!numCollectionCurrentMonth), "", Rec!numCollectionCurrentMonth)

                Rec.MoveNext
            Wend
        End If
        Rec.Close
        mCnn.Close
    End Sub
''    Private Sub FillGrid()
''        Dim mCnn    As New ADODB.Connection
''        Dim Rec     As New ADODB.Recordset
''        Dim mSql    As String
''        Dim objDb   As New clsDB
''        Dim mLoop   As Integer
''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
''
''          mSql = "SELECT  * FROM  faMonthlyDCB"
''          mSql = mSql + " INNER JOIN faDCBHeads ON faMonthlyDCB.intID=faDCBHeads.intID"
''
''         If cmbFinancialYear.ListIndex > 0 Then
''            mSql = mSql + " where  intYearID=  " & cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex) & ""
''
''         End If
''
''            mSql = mSql + " AND  intMonthID=  " & cmbMonth.ItemData(cmbMonth.ListIndex) & ""
''
''        Rec.Open mSql, mCnn
''        vsGrid.Rows = 1
''        If Not (Rec.BOF And Rec.EOF) Then
''            While Not (Rec.EOF)
''                vsGrid.Rows = vsGrid.Rows + 1
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!vchItem
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!fltPreActualdemand), "", Rec!fltPreActualdemand)
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!fltAmountAccured), "", Rec!fltAmountAccured)
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!fltCollectionUptoPreMnth), "", Rec!fltCollectionUptoPreMnth)
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!fltCollectionCurrent), "", Rec!fltCollectionCurrent)
''
''                Rec.MoveNext
''            Wend
''        End If
''        Rec.Close
''        mCnn.Close
''    End Sub
''    Private Function CheckExtractStatus(mYearID As Integer, mMonthID As Integer) As Boolean
''        Dim mCnn    As New ADODB.Connection
''        Dim Rec     As New ADODB.Recordset
''        Dim mSQl    As String
''        Dim objDB   As New clsDB
''
''        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
''        mSQl = "SELECT  dtDate  FROM  faMonthlyDCB WHERE intYearID=" & mYearID & " AND intMonthID=" & mMonthID
''        Rec.Open mSQl, mCnn
''        If Not (Rec.BOF And Rec.EOF) Then
''            CheckExtractStatus = True   'ALREADY EXTRACTED
''        Else
''            CheckExtractStatus = False
''            Call FillDetails
''        End If
''        Rec.Close
''        mCnn.Close
''    End Function
''    Private Function FillDetails()
''        Dim mCnn    As New ADODB.Connection
''        Dim Rec     As New ADODB.Recordset
''        Dim mSQl    As String
''        Dim objDB   As New clsDB
''        Dim arInput As Variant
''
''        arInput = Array(cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex), cmbMonth.ItemData(cmbMonth.ListIndex))
''        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
''        Set Rec = objDB.ExecuteSP("spGetExtractMonthlyDCBDetails", arInput, , , mCnn, adCmdStoredProc)
''        vsGrid.Rows = 1
''        If Not (Rec.BOF And Rec.EOF) Then
''            While Not (Rec.EOF)
''                vsGrid.Rows = vsGrid.Rows + 1
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!vchItem
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!fltDemandCurrent), "", Rec!fltDemandCurrent)
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!fltDemandArrear), "", Rec!fltDemandArrear)
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!fltCollectionCurrent), "", Rec!fltCollectionCurrent)
''                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!fltCollectionArrear), "", Rec!fltCollectionArrear)
''
''                Rec.MoveNext
''            Wend
''        End If
''             Rec.Close
''        mCnn.Close
''
''  End Function

   Private Function GetDCBYearID() As Integer    'FUNCTION TO GET intDCBYearID
    
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim mSql                As String
        Dim mDCBYearID          As Variant
            
        objDb.SetConnection mCnn
        mSql = " SELECT intDCBYearID   FROM faConfig"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mDCBYearID = IIf(IsNull(Rec!intDCBYearID), 0, Rec!intDCBYearID)
        End If
        Rec.Close
        GetDCBYearID = mDCBYearID
    End Function



  Private Function ExtractDCB()
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim mSql                As String
        Dim mRowCnt             As Integer
        Dim arInput              As Variant
        Dim mDCBYearID          As Integer
        Dim mMonthID            As Integer
    
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCnn.CommandTimeout = 1000000000
        
        PgrBar.value = 0
        PgrBar.Max = 100
        mExtractionState = False
       
        mDCBYearID = GetDCBYearID
        If mDCBYearID = 0 Then
            mDCBYearID = 2013
        ElseIf mDCBYearID = 2014 Then
            PgrBar.Visible = False
            Exit Function
        End If
        
        'START YEARLY EXTRACTION
        While mDCBYearID <= 2014 And mDCBYearID <> 0
            mCnn.Execute " Delete  from faMonthlyDCB Where intYearID = " & mDCBYearID
            For mMonthID = 1 To 12
                arInput = Array(mDCBYearID, mMonthID)
                
                'objDb.ExecuteSP "spExtractMonthlyDCBNEW", arInput, , , mCnn, adCmdStoredProc
                objDb.ExecuteSP "spExtactDCBHeadWise", arInput, , , mCnn, adCmdStoredProc
            Next
            
            mSql = " UPDATE faConfig SET intDCBYearID = " & mDCBYearID
            objDb.ExecuteSP mSql, , , , mCnn, adCmdText
            mDCBYearID = mDCBYearID + 1
            If PgrBar.value < PgrBar.Max - 2 Then
                PgrBar.value = PgrBar.Max - 1
            End If
        Wend
        mExtractionState = True
        mCnn.Close
     
    End Function
     
