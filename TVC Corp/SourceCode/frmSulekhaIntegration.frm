VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSulekhaIntegration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sulekha Project Search"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   Icon            =   "frmSulekhaIntegration.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFund 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   585
      Width           =   1605
   End
   Begin VB.CommandButton cmdSourceOfFund 
      BackColor       =   &H00F8FFF9&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3165
      TabIndex        =   1
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox txtProjectNameMal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "ML-TTRevathi"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6015
      TabIndex        =   3
      Top             =   585
      Width           =   3585
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4470
      TabIndex        =   4
      Top             =   1005
      Width           =   1020
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   75
      TabIndex        =   5
      Top             =   1350
      Width           =   9825
      _cx             =   17330
      _cy             =   7276
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSulekhaIntegration.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6015
      TabIndex        =   2
      Top             =   285
      Width           =   3585
   End
   Begin VB.TextBox txtProjectNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   285
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fund"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1110
      TabIndex        =   9
      Top             =   615
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Project Name in Malayalam"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3600
      TabIndex        =   8
      Top             =   600
      Width           =   2385
   End
   Begin VB.Label lblProjectNameEng 
      Caption         =   "Project Name in English"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3900
      TabIndex        =   7
      Top             =   285
      Width           =   2085
   End
   Begin VB.Label lblProjectNo 
      AutoSize        =   -1  'True
      Caption         =   "Project No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   615
      TabIndex        =   6
      Top             =   285
      Width           =   915
   End
End
Attribute VB_Name = "frmSulekhaIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '*********************************************************************************************'
    '              Form to list the Project Details after Project Synchronization                 '
    '*********************************************************************************************'
    Function AutoWordWrap(vs As VSFlexGrid)
        With vs
            .AutoSizeMode = flexAutoSizeRowHeight
            .WordWrap = True
            .AutoSize 0, .Cols - 1
'            .Cell(flexcpAlignment, 1, 1, .Rows - 1, .Cols - 1) = 0
        End With
    End Function
    Private Sub FillEstAmount()
        Dim mRowCount   As Double
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim objDb       As New clsDB
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mRowCount = 1
        For mRowCount = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mRowCount, 5) <> "" And vsGrid.TextMatrix(mRowCount, 6) <> "" Then
                mSql = "Select Sum(fltEstAmt) As Amount From suEstimation"
                mSql = mSql + " Where decProjectID = " & val(vsGrid.TextMatrix(mRowCount, 5))
                mSql = mSql + " And intYearID = " & val(vsGrid.TextMatrix(mRowCount, 6))
                Rec.Open mSql, mCnn
                While Not Rec.EOF
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
                    Rec.MoveNext
                Wend
                Rec.Close
            End If
        Next
    End Sub
    Private Sub FillvsGrid(Rec As ADODB.Recordset)
        Dim mRowCount   As Double
        
        mRowCount = 1
        vsGrid.Rows = 1
        vsGrid.Clear 1, 1
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!chvProjectName), "", Rec!chvProjectName)
            vsGrid.Cell(flexcpFontName, mRowCount, 2) = "ML-TTRevathi"
            vsGrid.Cell(flexcpFontSize, mRowCount, 2) = 11
            'vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!EstSum), "", Rec!EstSum)
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!SptAmt), "", Rec!SptAmt)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!ProjectID), "", Rec!ProjectID)
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
        AutoWordWrap vsGrid
    End Sub
    Private Sub cmdSearch_Click()
        Dim mCnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mProjectNo      As String
        Dim mProjectName    As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If txtProjectNo.Text = "" Then
            mProjectNo = "%"
        Else
            mProjectNo = txtProjectNo.Text
        End If
        If txtProjectName.Text = "" Then
            mProjectName = "%"
        Else
            mProjectName = txtProjectName.Text
        End If
        
        mSql = "Select suProjectDetails.decProjectID[ProjectID],chvProjectSlNo,chvProjectName,chvProjectNameEnglish,Sum(fltEstAmt) As EstSum,Sum(fltAmount) As SptAmt,suProjectDetails.intYearID From suProjectDetails"
        mSql = mSql + " Left Join suEstimation On suProjectDetails.decProjectID = suEstimation.decProjectID"
        mSql = mSql + " Left Join suExpenditures On suEstimation.decProjectID = suExpenditures.decProjectID And suEstimation.intFundID = suExpenditures.intFundID"
        mSql = mSql + " Where suProjectDetails.chvProjectSlNo LIKE '%" & mProjectNo & "%'"
        mSql = mSql + " And suProjectDetails.chvProjectNameEnglish LIKE '" & mProjectName & "%'"
        If txtProjectName.Text <> "" Then
            mSql = mSql + " And suProjectDetails.chvProjectName LIKE '" & txtProjectNameMal.Text & "%'"
        End If
        If txtFund.Tag <> "" Then
            mSql = mSql + " And suEstimation.intFundID  = " & txtFund.Tag
        End If
            mSql = mSql + " And  suProjectDetails.intYearID = " & gbFinancialYearID & " "
        
        mSql = mSql + " Group By suProjectDetails.decProjectID,chvProjectSlNo,chvProjectName,chvProjectNameEnglish,suProjectDetails.intYearID"
        mSql = mSql + " Order by chvProjectSlNo"
        Rec.Open mSql, mCnn
        Call FillvsGrid(Rec)
        Rec.Close
        Call FillEstAmount
    End Sub
    
    Private Sub cmdSourceOfFund_Click()
        On Error GoTo err
        gbSearchStr = ""
        gbSearchID = -1
        txtFund.Text = ""
        txtFund.Tag = ""
        frmSearchMasters.SQLQry = "Select * From suSourceOfFund"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If Not gbSearchStr = "" Then
            txtFund.Text = gbSearchStr
            txtFund.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Activate()
        'Me.Left = 0
        'Me.Top = 0
        Me.Height = 6000
        Me.Width = 10050
    End Sub

    Private Sub Form_Load()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya

        vsGrid.SelectionMode = flexSelectionByRow
        mSql = "Select suProjectDetails.decProjectID[ProjectID],chvProjectSlNo,chvProjectName,chvProjectNameEnglish,Sum(fltEstAmt) As EstSum,Sum(fltAmount) As SptAmt,suProjectDetails.intYearID From suProjectDetails"
        mSql = mSql + " Left Join suEstimation On suProjectDetails.decProjectID = suEstimation.decProjectID"
        mSql = mSql + " Left Join suExpenditures On suEstimation.decProjectID = suExpenditures.decProjectID And suEstimation.intFundID = suExpenditures.intFundID"
        mSql = mSql + " Where suProjectDetails.intYearID = " & gbFinancialYearID & " "
        mSql = mSql + " Group By suProjectDetails.decProjectID,chvProjectSlNo,chvProjectName,chvProjectNameEnglish,suProjectDetails.intYearID"
        mSql = mSql + " Order by chvprojectSlNo"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
        Rec.Close
        End If
        Call FillEstAmount
    End Sub
    
    Private Sub vsGrid_DblClick()
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim objDb       As New clsDB
        Dim mRowCount   As Double
'        Dim m           As Integer
'        Dim mSqlFund    As String
'        Dim RecFund     As New ADODB.Recordset
'        Dim mCnnFund    As New ADODB.Connection
        
        '*********************************************************************************************'
        '              Procedure to list the Fundwise Details of a particular Project                 '
        '*********************************************************************************************'
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        objDb.CreateNewConnection mCnnFund, enuSourceString.Saankhya
        
        If vsGrid.Row <> 0 And vsGrid.TextMatrix(vsGrid.Row, 5) <> "" Then
            frmEstimationDetails.txtProjectNo.Text = vsGrid.TextMatrix(vsGrid.Row, 0)
            frmEstimationDetails.txtProjectNo.Tag = val(vsGrid.TextMatrix(vsGrid.Row, 5)) ' Project ID
            
            frmEstimationDetails.txtProjectNameEng.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            frmEstimationDetails.txtProjectName.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            frmEstimationDetails.vsGrid.Clear 1, 1
            frmEstimationDetails.vsGrid.Rows = 1
            mRowCount = 1
            mSql = "Select suEstimation.intID[EstimationID],suEstimation.intFundID[FundID],suEstimation.decProjectID,vchSourceFundShortName,fltEstAmt,Sum(fltAmount) As Amount From suEstimation"
            mSql = mSql + " Inner Join suSourceOfFund On suEstimation.intFundID = suSourceOfFund.intSourceFundID"
            mSql = mSql + " Left Join suExpenditures On suEstimation.decProjectID = suExpenditures.decProjectID"
            mSql = mSql + " Where suEstimation.decProjectID = " & val(vsGrid.TextMatrix(vsGrid.Row, 5))
            mSql = mSql + " And suEstimation.intYearID = " & val(vsGrid.TextMatrix(vsGrid.Row, 6))
            mSql = mSql + " Group By suEstimation.intID,suEstimation.intFundID,suEstimation.decProjectID,vchSourceFundShortName,fltEstAmt"
            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
                'For m = 0 To Rec.Fields.Count - 1
'                    If Left(Rec.Fields(m).Name, 3) = "flt" And Rec.Fields(m).Name <> "fltTotal" And IsNull(Rec.Fields(m)) = False And Rec.Fields(m) <> 0 Then
            While Not Rec.EOF
            frmEstimationDetails.vsGrid.AddItem ""
            frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchSourceFundShortName), "", Rec!vchSourceFundShortName)
            
'                        mSqlFund = "Select intSourceFundID From suSourceOfFund"
'                        mSqlFund = mSqlFund + " Where vchSourceFundShortName ='" & frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 0) & "'"
'                        RecFund.Open mSqlFund, mCnnFund
'                        If Not (Rec.EOF And Rec.BOF) Then
                
'                        End If
'                        RecFund.Close
            frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!fltEstAmt), "", Rec!fltEstAmt)
            frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
            frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!EstimationID), "", Rec!EstimationID)
            frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 4) = vsGrid.TextMatrix(vsGrid.Row, 5)
            frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!FundID), "", Rec!FundID)
            Rec.MoveNext
            mRowCount = mRowCount + 1
'                    End If
                'Next
'            End If
            Wend
            Rec.Close
'            mSQL = "Select * From suEstimation"
'            mSQL = mSQL + " Left Join suSourceofFund on suEstimation.intFundID = suSourceofFund.intFundID"
'            mSQL = mSQL + " Left Join suExpenditures On suEstimation.decProjectID = suExpenditures.decProjectID And suEstimation.intFundID = suExpenditures.intFundID"
'            mSQL = mSQL + " Where suEstimation.decProjectID = " & vsGrid.TextMatrix(vsGrid.Row, 5)
'            Rec.Open mSQL, mCnn
'            While Not Rec.EOF
'                frmEstimationDetails.vsGrid.AddItem ""
'                frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchFundName), "", Rec!vchFundName)
'                frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!fltEstAmt), "", Rec!fltEstAmt)
'                frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'                frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!intID), "", Rec!intID)
'                frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 4) = vsGrid.TextMatrix(vsGrid.Row, 5)
'                frmEstimationDetails.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!intFundID), "", Rec!intFundID)
'                mRowCount = mRowCount + 1
'                Rec.MoveNext
'            Wend
            frmEstimationDetails.Show vbModal
        End If
    End Sub
