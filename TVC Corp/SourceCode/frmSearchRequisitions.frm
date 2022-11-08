VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearchRequisitions 
   Caption         =   "SearchRequisitions"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRpTo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtRpFrom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtRequistionNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7530
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3150
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   9900
      _cx             =   17462
      _cy             =   5556
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchRequisitions.frx":0000
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
   Begin MSComCtl2.DTPicker dtpkrFromDate 
      Height          =   315
      Left            =   2580
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483638
      Format          =   64815105
      CurrentDate     =   39612
   End
   Begin MSComCtl2.DTPicker dtpkrToDate 
      Height          =   315
      Left            =   5325
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483638
      Format          =   64815105
      CurrentDate     =   39612
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requistion NO:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5985
      TabIndex        =   2
      Top             =   135
      Width           =   1530
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DateTo"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   1
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSearchRequisitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    If mPreviousYearMode Then
            txtDateFrom.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            txtDateTo.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
        Else
            txtDateFrom.Text = DdMmmYy(gbEndingDate)
            txtDateTo.Text = DdMmmYy(gbEndingDate)
        End If
        Call FillGrid
End Sub

Private Sub cmdsearch_Click()
'If Trim(txtRpFrom.Text) = "" Then
'            MsgBox "From date is Mandatory"
'            Exit Sub
'        End If
'        If Trim(txtRpTo.Text) = "" Then
'            MsgBox "To date is Mandatory"
'            Exit Sub
'        End If
        Call FillGrid
End Sub

 Private Sub FillGrid()
            Dim objdb       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim mRow As Integer
            Dim mSql As String
            Dim arInput As Variant
            Dim mCnnSulekha    As New ADODB.Connection
            If Not (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                MsgBox "Connection To Plan[Sulekha] Module not found", vbCritical
                Exit Sub
            End If
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'            If txtRequistionNo.Text = "" Then
'                txtRequistionNo.Text = ""
'            End If
'
'
'            arInput = Array(txtDatefrom.Text, txtDateTo.Text, txtRequistionNo.Text)
'
'            Set Rec = objDB.ExecuteSP("spGetRequisitionDetails", arInput, , , mCnn, adCmdStoredProc)

            mSql = "        SELECT intID,vchRequisitionNo,dtRequisitionDate,faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID,"
            mSql = mSql + " faSubSidiaryAccountHeads.vchName,suSourceOfFund.intSourceFundID,suSourceOfFund.vchSourceFundName,"
            mSql = mSql + " faTransactionCategory.intCategoryID,faTransactionCategory.vchTransactionCategory "
            mSql = mSql + " FROM faAllotments "
            mSql = mSql + " INNER JOIN faSubSidiaryAccountHeads on faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID=faAllotments.intImplementingOfficersID"
            mSql = mSql + " INNER JOIN suSourceOfFund on suSourceOfFund.intSourceFundID=faAllotments.intSourceID"
            mSql = mSql + " LEFT JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faAllotments.intFundCategoryID where "
            If txtRequistionNo.Text <> "" Then
                  mSql = mSql + "  vchRequisitionNo like '" & txtRequistionNo.Text & "%'"
                  
            Else
                If txtRpFrom.Text <> "" And txtRpTo.Text <> "" Then
                  mSql = mSql + " dtRequisitionDate BETWEEN '" & txtRpFrom.Text & "'AND '" & txtRpTo.Text & "'"
                Else
                  MsgBox "Please enter date"
                End If
            End If
            
            If mPreviousYearMode = 0 Then
                mSql = mSql + "AND tnyStatus <> 2 "
            Else
                If mPreviousYearTaskID = 3 Then
                    mSql = mSql + "AND tnyStatus = 0 "
                Else
                    mSql = mSql + "AND tnyStatus <> 2 "
                End If
            End If

            
            
            
            

            
            'mSQL = mSQL + " WHERE dtRequisitionDate BETWEEN '" & txtDatefrom & "' And '" & txtDateTo & "'"

            

           
            
'            If txtIMPOName.Text <> "" Then
'                mSQL = mSQL + " And IsNull(faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID,0)=" & txtIMPOName.Tag
'            End If
'            If txtSourceofFund.Text <> "" Then
'                mSQL = mSQL + " And IsNull(suSourceOfFund.intSourceFundID,0)=" & txtSourceofFund.Tag
'            End If
'            If txtCatogory.Text <> "" Then
'                mSQL = mSQL + " And IsNull(faTransactionCategory.intCategoryID,0)=" & txtCatogory.Tag
'            End If
            mSql = mSql + " Order by dtRequisitionDate "
           Rec.Open mSql, mCnn
           vsGrid.Clear
           mRow = 1
          ' vsGrid.Clear 1, 1
           vsGrid.Rows = 1
           vsGrid.TextMatrix(0, 1) = "Requisition No"
           vsGrid.TextMatrix(0, 2) = "Date"
           vsGrid.TextMatrix(0, 3) = "Implimenting Officer"
           vsGrid.TextMatrix(0, 4) = "Source of Fund"
           vsGrid.TextMatrix(0, 5) = "Category"
           
           If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                      vsGrid.Rows = vsGrid.Rows + 1
                      vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!intID), "", Rec!intID)
                      vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
                      vsGrid.TextMatrix(mRow, 2) = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
                      vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                      vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                      vsGrid.TextMatrix(mRow, 5) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                      Rec.MoveNext
                      mRow = mRow + 1
                Wend
               
               
           End If
    End Sub

Private Sub dtpRPFrom_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    'txtRpFrom.Text = CheckDateInMMM(dtpkrFromDate.Value)
End Sub

Private Sub dtpkrFromDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    txtRpFrom.Text = CheckDateInMMM(dtpkrFromDate.Value)
End Sub

Private Sub dtpkrFromDate_CloseUp()
If CDate(dtpkrFromDate.Value) Then
        If CDate(dtpkrToDate.Value) Then
                If CDate(dtpkrFromDate.Value) > CDate(dtpkrFromDate.Value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpkrFromDate.Value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtRpFrom.Text = CheckDateInMMM(dtpkrFromDate.Value)
        End If
        txtRpFrom.Text = DdMmmYy(dtpkrFromDate.Value)

End Sub

Private Sub dtpkrToDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    txtRpTo.Text = CheckDateInMMM(dtpkrToDate.Value)
End Sub

Private Sub dtpkrToDate_CloseUp()
If CDate(dtpkrToDate.Value) Then
            If CDate(dtpkrFromDate.Value) Then
                If CDate(dtpkrFromDate.Value) > CDate(dtpkrToDate.Value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpkrToDate.Value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtRpTo.Text = CheckDateInMMM(dtpkrToDate.Value)
        End If
        txtRpTo.Text = DdMmmYy(dtpkrToDate.Value)
End Sub

Private Sub Form_Load()
    If mPreviousYearMode Then
            txtDateFrom.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            txtDateTo.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
        Else
            txtRpFrom.Text = DdMmmYy(gbTransactionDate)
            txtRpTo.Text = DdMmmYy(gbTransactionDate)
        End If
End Sub

'Private Sub txtRpFrom_LostFocus()
'    Dim mDate As Date
'        If IsDate(txtRpFrom) Then
'            txtRpFrom.Text = DdMmmYy(txtRpFrom.Text)
'        Else
'            txtRpFrom.Text = CheckDateInMMM(txtRpFrom)
'        End If
'
'
'        If IsDate(txtRpFrom) Then
'            mDate = txtRpFrom
'        End If
'
'End Sub

Private Sub txtDateTo_LostFocus()
'  Dim mDate As Date
'        If IsDate(txtDateTo) Then
'            mDate = txtDateTo
'            txtDateTo.Text = DdMmmYy(mDate)
'        End If
'        If mPreviousYearMode Then
'              If Not (mDate >= DateAdd("yyyy", -1, gbStartingDate) And mDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
'                txtDateTo.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
'            End If
'        Else
'            If Not (mDate >= gbStartingDate And mDate <= gbEndingDate) Then
'                txtDateTo.Text = DdMmmYy(gbTransactionDate)
'            End If
'        End If
End Sub


Private Sub txtRpFrom_LostFocus()
    txtRpFrom.Text = Format(txtRpFrom.Text, "dd-mmm-yyyy")
End Sub

Private Sub txtRpTo_LostFocus()
    txtRpTo.Text = Format(txtRpTo.Text, "dd-mmm-yyyy")
End Sub

Private Sub vsGrid_Click()
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        'Dim Rec         As New ADODB.RecordsettxtDatefrom

        If vsGrid.Row > 0 Then 'vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
            gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 1)
            gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 0)
            arInput = Array(gbSearchStr)
            frmNewRpt.rptFileName = App.Path & "\Reports\rptRequisition.rpt"
            frmNewRpt.WindowState = vbMaximized
            frmNewRpt.InputParameters = arInput
            frmSearchRequisitions.Hide
            'Unload Me
            Call frmNewRpt.ShowReport
            frmNewRpt.Show


            'Unload Me
        End If

End Sub
