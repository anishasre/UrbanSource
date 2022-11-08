VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmAFSClosingSourceOfFund 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ACR - Closing Source Of Fund"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "frmAFSClosingSourceOfFund.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5370
      Left            =   60
      TabIndex        =   0
      Top             =   855
      Width           =   10590
      _cx             =   18680
      _cy             =   9472
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
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAFSClosingSourceOfFund.frx":1CCA
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
      Editable        =   2
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
   Begin VB.CommandButton cmdReturnforCorrection 
      Caption         =   "Return for Correction"
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
      Left            =   1380
      TabIndex        =   10
      Top             =   6330
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve"
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
      Left            =   165
      TabIndex        =   9
      Top             =   6330
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   6300
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
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
      Left            =   5955
      TabIndex        =   3
      Top             =   6300
      Width           =   1185
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2925
      TabIndex        =   1
      Text            =   "31-March-2017"
      Top             =   405
      Width           =   1635
   End
   Begin VB.Frame frmSource 
      Caption         =   "Add New Source"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9360
      TabIndex        =   11
      Top             =   6390
      Visible         =   0   'False
      Width           =   2985
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   900
         TabIndex        =   16
         Top             =   1935
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   1440
         Width           =   2355
      End
      Begin VB.CommandButton cmdSource 
         Caption         =   "..."
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
         Left            =   2430
         TabIndex        =   14
         Top             =   675
         Width           =   375
      End
      Begin VB.TextBox txtSourceOfFund 
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Top             =   675
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "Source Of Fund :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   17
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Amount :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Label lblmsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13305
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10710
      TabIndex        =   7
      Top             =   1125
      Width           =   195
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Closing Balance of Source Of Funds as on 31-March-2017"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1005
      Left            =   10980
      TabIndex        =   6
      Top             =   1140
      Width           =   1905
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5370
      Left            =   10650
      TabIndex        =   5
      Top             =   855
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing SourceOfFund Balance As On"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   405
      Width           =   2715
   End
End
Attribute VB_Name = "frmAFSClosingSourceOfFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim RowCount As Integer
    Dim mCheckPrintStatus As Integer
    Dim mFlag As Integer
    Dim mApprovalStatus As Integer
    
    Private Sub cmdAdd_Click()
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mFlag   As Integer
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = " SELECT * FROM faExtractAllotments WHERE intSourceOfFundID=19"
        Rec.Open mSQL, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            mFlag = 1
        Else
            mFlag = 0
        End If
        Rec.Close
        
        If mFlag = 1 Then
            mSQL = "INSERT INTO faExtractAllotments  "
            mSQL = mSQL + "(intAllotmentID, vchAllotmentNo, dtAllotmentDate, intSourceOfFundID, intCategoryID, intSchemeID, tnyInstalmentNo, intCrAccountHeadID,"
            mSQL = mSQL + " intFunctionaryID, intFunctionID, intGrossAccountHeadID, fltAmount, dtOfEntry, intLocalBodyID, intFinancialYearID, tnyCancelledFlag, tnyStatus,"
            mSQL = mSQL + " intTransactionTypeID, tnyOpening)"
            mSQL = mSQL + "VALUES (NULL,NULL,NULL,19,1,NULL,NULL,NULL,NULL,NULL,NULL," & val(txtAmount.Text) & "," & gbTransactionDate & "," & gbLBID & "," & gbFinancialYearID & ",0,2,NULL,1)"
            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        Else
            MsgBox "Already Added Source Of Fund", vbInformation
        End If
    End Sub

    Private Sub cmdApprove_Click()
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mDBVersion As Integer
        
         If objDB.SetConnection(mCnn) Then
            If gbLBPanchayat = 1 Then
                mSQL = " SELECT * FROM faDBSubVersions WHERE intDBSubVersionID>=15"
            Else
                mSQL = " SELECT * FROM faDBSubVersions WHERE intDBSubVersionID>=17"
            End If
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mDBVersion = 1
            Else
                mDBVersion = 0
            End If
            Rec.Close
        End If
        
        If mDBVersion = 1 Then
            Dim mLoop As Integer
            For mLoop = 1 To vsGrid.Rows - 1
                If val(vsGrid.TextMatrix(mLoop, 2)) <> val(vsGrid.TextMatrix(mLoop, 3)) Then
                    If Trim(vsGrid.TextMatrix(mLoop, 3)) <> "" Then
                        mSQL = "UPDATE faExtractAllotments SET fltAmount = " & Format(val(vsGrid.TextMatrix(mLoop, 3)), "#0")
                        mSQL = mSQL + " Where intFinancialYearID = " & gbFinancialYearID
                        mSQL = mSQL + " AND intSourceOfFundID = " & val(vsGrid.TextMatrix(mLoop, 4))
                        mSQL = mSQL + " AND intCategoryID = " & val(vsGrid.TextMatrix(mLoop, 7))
                        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                    End If
                End If
            Next mLoop
            mSQL = "Update faExtractAllotments set tnyStatus=2 where intFinancialYearID=" & gbFinancialYearID & "  "
            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            cmdApprove.Enabled = False
            cmdReturnforCorrection.Enabled = False
        Else
            MsgBox "VERSION UPDATE PENDING!!!!!!REPLACE LATEST EXE", vbCritical
            Exit Sub
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub
     
     Private Sub FillGridFromExtract()
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSQL As String
        Dim mArrIn As Variant
        

        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = " SELECT faExtractAllotments.intID intID,tnyStatus,suSourceofFund.intSourceFundID as SourceID,suSourceofFund.vchSourceFundName as Source,faTransactionCategory.intCategoryID as CategoryID,"
        mSQL = mSQL + " faTransactionCategory.vchTransactionCategory as Category,fltAmount as Balance,intFinancialYearID, fltActual, ISNULL(intVersion,0) intVersion "
        mSQL = mSQL + " From faExtractAllotments"
        mSQL = mSQL + " INNER JOIN suSourceofFund ON faExtractAllotments.intSourceofFundID=suSourceofFund.intSourceFundID"
        mSQL = mSQL + " LEFT JOIN  faTransactionCategory ON faExtractAllotments.intCategoryID=faTransactionCategory.intCategoryID"
        mSQL = mSQL + " Where intFinancialYearID=" & gbFinancialYearID
        Rec.Open mSQL, mCnn
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRow = 1
        If Not (Rec.BOF And Rec.EOF) Then
             While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!Source), "", Rec!Source)
                vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!Category), "", Rec!Category)
                'vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!fltActual), 0, Rec!fltActual)
                vsGrid.TextMatrix(mRow, 2) = 0
                If Rec!fltActual <> Rec!Balance Then
                    vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!Balance), "", Rec!Balance)
                Else
                    vsGrid.TextMatrix(mRow, 3) = ""
                End If
                
                vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!SourceID), "", Rec!SourceID)
                vsGrid.TextMatrix(mRow, 5) = IIf(IsNull(Rec!intFinancialYearID), "", Rec!intFinancialYearID)
                vsGrid.TextMatrix(mRow, 6) = mRow
                vsGrid.TextMatrix(mRow, 7) = IIf(IsNull(Rec!CategoryID), "", Rec!CategoryID)
                
                mRow = mRow + 1
                Rec.MoveNext
             Wend
             RowCount = mRow
        End If
        Rec.Close
        mCnn.Close
       
    End Sub
    
    Private Sub FillGrid()
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSQL As String
        Dim mArrIn As Variant
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mArrIn = Array(gbFinancialYearID)
        Set Rec = objDB.ExecuteSP("spGetExtractAllotmentDetails", mArrIn, , , mCnn, adCmdStoredProc)
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRow = 1
        If Not (Rec.BOF And Rec.EOF) Then
             While Not Rec.EOF
                '                vsGrid.Rows = vsGrid.Rows + 1
                '                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!Source), "", Rec!Source)
                '                vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!Category), "", Rec!Category)
                '                vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!Balance), "", Rec!Balance)
                '                vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!SourceID), "", Rec!SourceID)
                '                vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!YearID), "", Rec!YearID)
                '                vsGrid.TextMatrix(mRow, 5) = mRow
                '                vsGrid.TextMatrix(mRow, 6) = IIf(IsNull(Rec!CategoryID), "", Rec!CategoryID)
                
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!Source), "", Rec!Source)
                vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!Category), "", Rec!Category)
                'vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!Balance), 0, Rec!Balance)
                vsGrid.TextMatrix(mRow, 2) = 0
                vsGrid.TextMatrix(mRow, 3) = ""
                
                vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!SourceID), "", Rec!SourceID)
                vsGrid.TextMatrix(mRow, 5) = IIf(IsNull(Rec!YearID), "", Rec!YearID)
                vsGrid.TextMatrix(mRow, 6) = mRow
                vsGrid.TextMatrix(mRow, 7) = IIf(IsNull(Rec!CategoryID), "", Rec!CategoryID)
                mRow = mRow + 1
                Rec.MoveNext
             Wend
             RowCount = mRow
        End If
        Rec.Close
        mCnn.Close
    End Sub
    
    
''''
''''    Private Function CheckClosingBalnce() As Boolean
''''        Dim mCnn  As New ADODB.Connection
''''        Dim objDb As New clsDB
''''        Dim Rec   As New ADODB.Recordset
''''        Dim mSql  As String
''''        Dim mTrAccHeadId As Integer
''''
''''        If objDb.SetConnection(mCnn) Then
''''            mSql = " SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID=" & gbFinancialYearID - 1
''''            Rec.Open mSql, mCnn
''''            If Not (Rec.EOF And Rec.BOF) Then
''''                While Not (Rec.EOF)
''''                    If IsNull(Rec!tnyStatus) Then
''''                         CheckClosingBalnce = True
''''                    Else
''''                        CheckClosingBalnce = False
''''                    End If
''''                    Rec.MoveNext
''''                Wend
''''            Else
''''                CheckClosingBalnce = True
''''            End If
''''            Rec.Close
''''        End If
''''    End Function
''''    Private Function CheckVerifyStatus() As Boolean
''''        Dim mCnn  As New ADODB.Connection
''''        Dim objDb As New clsDB
''''        Dim Rec   As New ADODB.Recordset
''''        Dim mSql  As String
''''        Dim mTrAccHeadId As Integer
''''
''''        If objDb.SetConnection(mCnn) Then
''''            mSql = " SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID=" & gbFinancialYearID
''''            Rec.Open mSql, mCnn
''''            If Not (Rec.EOF And Rec.BOF) Then
''''                While Not (Rec.EOF)
''''                    If Rec!tnyStatus = 0 Then
''''                         CheckVerifyStatus = True
''''                    Else
''''                        CheckVerifyStatus = False
''''                    End If
''''                    Rec.MoveNext
''''                Wend
''''            Else
''''                CheckVerifyStatus = False
''''            End If
''''            Rec.Close
''''        End If
''''    End Function
''''    Private Function CheckPrintDeclarationStatus() As Boolean
''''        Dim mCnn  As New ADODB.Connection
''''        Dim objDb As New clsDB
''''        Dim Rec   As New ADODB.Recordset
''''        Dim mSql  As String
''''        Dim mTrAccHeadId As Integer
''''
''''        If objDb.SetConnection(mCnn) Then
''''            mSql = " SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID=" & gbFinancialYearID
''''            Rec.Open mSql, mCnn
''''            If Not (Rec.EOF And Rec.BOF) Then
''''                While Not (Rec.EOF)
''''                    If Rec!tnyStatus = 1 Then
''''                         CheckPrintDeclarationStatus = True
''''                    Else
''''                        CheckPrintDeclarationStatus = False
''''                    End If
''''                    Rec.MoveNext
''''                Wend
''''            Else
''''                CheckPrintDeclarationStatus = False
''''            End If
''''            Rec.Close
''''        End If
''''    End Function
''''    Private Function CheckApproveStatusBySec() As Boolean
''''        Dim mCnn  As New ADODB.Connection
''''        Dim objDb As New clsDB
''''        Dim Rec   As New ADODB.Recordset
''''        Dim mSql  As String
''''        Dim mTrAccHeadId As Integer
''''
''''        If objDb.SetConnection(mCnn) Then
''''            mSql = " SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID=" & gbFinancialYearID
''''            Rec.Open mSql, mCnn
''''            If Not (Rec.EOF And Rec.BOF) Then
''''                While Not (Rec.EOF)
''''                    If Rec!tnyStatus = 1 Then
''''                         CheckApproveStatusBySec = True
''''                    Else
''''                        CheckApproveStatusBySec = False
''''                    End If
''''                    Rec.MoveNext
''''                Wend
''''            Else
''''                CheckApproveStatusBySec = False
''''            End If
''''            Rec.Close
''''        End If
''''    End Function
''''     Private Function CheckFinalApproveStatusBySec() As Boolean
''''        Dim mCnn  As New ADODB.Connection
''''        Dim objDb As New clsDB
''''        Dim Rec   As New ADODB.Recordset
''''        Dim mSql  As String
''''        Dim mTrAccHeadId As Integer
''''
''''        If objDb.SetConnection(mCnn) Then
''''            mSql = " SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID=" & gbFinancialYearID
''''            Rec.Open mSql, mCnn
''''            If Not (Rec.EOF And Rec.BOF) Then
''''                While Not (Rec.EOF)
''''                    If Rec!tnyStatus = 2 Then
''''                         CheckFinalApproveStatusBySec = True
''''                    Else
''''                        CheckFinalApproveStatusBySec = False
''''                    End If
''''                    Rec.MoveNext
''''                Wend
''''            Else
''''                CheckFinalApproveStatusBySec = False
''''            End If
''''            Rec.Close
''''        End If
''''    End Function
    Private Sub cmdReturnforCorrection_Click()
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        
        mSQL = "Delete from  faExtractAllotments Where intFinancialYearID=" & gbFinancialYearID & "  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        cmdReturnforCorrection.Enabled = False
        Call FillGridFromExtract
    End Sub



    Private Sub cmdSource_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund Where intSourceFundID=19"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtSourceOfFund.Text = gbSearchStr
            txtSourceOfFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdVerify_Click()
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSQL As String
        Dim mArrIn As Variant
        Dim i As Integer
        Dim mStatusFlag As Integer
        

        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mArrIn = Array(gbFinancialYearID, 0)
        objDB.ExecuteSP "spExtractAllotment", mArrIn, , , mCnn
        
        mStatusFlag = GetStatusFlag
        If mStatusFlag = 0 Then         'If Extracted to the table
            frmViewVoucher.MultipleVouchers = False
            frmViewVoucher.FormName = "frmAFSClosingSourceOfFund"
            frmViewVoucher.ArrayIn = Array(gbFinancialYearID)
            frmViewVoucher.Height = 8000
            frmViewVoucher.Top = 200
            frmViewVoucher.crvReport.Height = 7000
            frmViewVoucher.Left = (frmMenu.Width - Me.Width) / 2
            
            frmViewVoucher.Show vbModal
            
            If mCheckPrintStatus = 1 Then
                mSQL = "Update faExtractAllotments set tnyStatus=1 where intFinancialYearID=" & gbFinancialYearID & "  "
                cmdVerify.Enabled = False
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
        End If
        mCnn.Close
    End Sub
    
    Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mTrAccHeadId As Integer
        
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                
                'NOTE: Checking in Previous Year
                '      IF APPROVED tnyStatus will be 0 ELSE NULL
                Rec.Close
                mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID - 1
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If IsNumeric(Rec!tnyStatus) Then
                        If Rec!tnyStatus = 2 Then
                            GetStatusFlag = 9
                        Else
                            GetStatusFlag = -1
                        End If
                    Else
                        GetStatusFlag = -1
                    End If
                Else
                    GetStatusFlag = -1
                End If
                
            End If
            If Rec.State = 1 Then
                Rec.Close
            End If
        End If
    End Function
    
    Private Sub CheckClosingBalanceWithExtractedData()
        'Note: If difference is in Extracted Data and Calculated Closing Balance
        
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSQL As String
        Dim mArrIn As Variant
        Dim RecExtracted As New ADODB.Recordset
        Dim mMsg As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mArrIn = Array(gbFinancialYearID)
        Set Rec = objDB.ExecuteSP("spGetExtractAllotmentDetails", mArrIn, , , mCnn, adCmdStoredProc)
        
        mSQL = " SELECT faExtractAllotments.intID intID,tnyStatus,suSourceofFund.intSourceFundID as SourceID,suSourceofFund.vchSourceFundName as Source,faTransactionCategory.intCategoryID as CategoryID,"
        mSQL = mSQL + " faTransactionCategory.vchTransactionCategory as Category,fltAmount as Balance,intFinancialYearID"
        mSQL = mSQL + " From faExtractAllotments"
        mSQL = mSQL + " INNER JOIN suSourceofFund ON faExtractAllotments.intSourceofFundID=suSourceofFund.intSourceFundID"
        mSQL = mSQL + " LEFT JOIN  faTransactionCategory ON faExtractAllotments.intCategoryID=faTransactionCategory.intCategoryID"
        mSQL = mSQL + " Where intFinancialYearID=" & gbFinancialYearID

        RecExtracted.CursorLocation = adUseClient
        RecExtracted.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
        
        If Not (RecExtracted.BOF And RecExtracted.EOF) Then
            While Not Rec.EOF
                RecExtracted.MoveFirst
                RecExtracted.Find "SourceID = " & Rec!SourceID, , adSearchForward
                If RecExtracted.EOF Then
                    MsgBox "Not Found"
                Else
                    If RecExtracted!Balance <> Rec!Balance Then
                        'MsgBox "NotMatching : " & RecExtracted!Balance & " , " & Rec!Balance & " | " & RecExtracted!SourceID & " | " & Rec!SourceID
                        'Debug.Print RecExtracted!Source, RecExtracted!Balance, Rec!Balance, RecExtracted!SourceID, Rec!SourceID
                        mFlag = 1
                    End If
                End If
                Rec.MoveNext
            Wend
            If mFlag = 1 Then
                cmdVerify.Enabled = False
                mMsg = ""
                mMsg = mMsg + "Closing Balance Of Source Of Fund Changed after Verification" + vbCrLf
                mMsg = mMsg + "   Therefore Declaration Cannot be Printed                  " + vbCrLf
                mMsg = mMsg + "Do you Want to Brought Down New Closing Balance?            "
                
                If MsgBox(mMsg, vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                    Call FillGrid
                    cmdVerify.Caption = "Verify"
                    cmdVerify.Enabled = True
                End If
            End If
        End If
        RecExtracted.Close
    End Sub

    Private Sub Form_Load()
    
        Dim mStatusFlag As Integer   'NOTE:
                                     '      -1 = Previous year Data not found
                                     '       9 = Previous Year OB is Approved
                                     '       0 = Current Year's OB is Extracted
                                     '       1 = Declaration is Printed after verification
                                     '       2 = Approved by Secretary
        mStatusFlag = GetStatusFlag
        mApprovalStatus = mStatusFlag
        
        '-------------------------------------------------------------------------------'
        ' NOTE: Checks whether Secretary is Approved the Source of Fund Closing Balance '
        '-------------------------------------------------------------------------------'
        If mStatusFlag = 2 Then 'NOTE: 2=Sectretary Approved
            cmdApprove.Enabled = False
            cmdReturnforCorrection.Enabled = False
            cmdVerify.Enabled = False
            Call FillGridFromExtract
            Exit Sub
        End If
         
        ' NOTE: Sectretary is NOT Approved
        If gbSeatGroupID = gbSeatGroupSecretary Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdApprove.Visible = True
            cmdApprove.Left = 4700
            cmdApprove.Top = 6300
            'cmdReturnforCorrection.Enabled = True
            'cmdReturnforCorrection.Enabled = False
            If gbFinancialYearID = 2014 Then
                cmdReturnforCorrection.Visible = True
            Else
                cmdReturnforCorrection.Visible = False
            End If
            cmdReturnforCorrection.Left = 6000
            cmdReturnforCorrection.Top = 6300
            cmdVerify.Visible = False
            cmdClose.Visible = False
            
            If Not (mStatusFlag = 9 Or mStatusFlag = -1) Then
                Call FillGridFromExtract ' NOTE: Fetch and Displays Data from Table
                If mStatusFlag = 1 Then 'NOTE: Declaration Printed
                    cmdApprove.Enabled = True
                    'cmdReturnforCorrection.Enabled = True
                    'cmdReturnforCorrection.Enabled = False
                    If gbFinancialYearID = 2014 Then
                        cmdReturnforCorrection.Visible = True
                    Else
                        cmdReturnforCorrection.Visible = False
                    End If
                Else
                    cmdApprove.Enabled = False
                    cmdReturnforCorrection.Enabled = False
                End If
            End If
        Else 'NOTE: User --> Accountant
            
            'NOTE: Previous Year's Opening Balance Extracted is Approved
            '      then CheckClosingBalance = FALSE
            If mStatusFlag = 9 Then
                cmdVerify.Enabled = True
                lblmsg.Visible = False
                Call FillGrid           'NOTE: Calculated and Fill Grid
            ElseIf mStatusFlag = 0 Then 'NOTE: Data is already Extracted by Verify Command
                cmdVerify.Caption = "Print Declartion"
                cmdVerify.Enabled = True
                Call FillGridFromExtract
                Call CheckClosingBalanceWithExtractedData
            ElseIf mStatusFlag = 1 Then  'NOTE: Declaration Prited and send to Secretary
                cmdVerify.Caption = "Verify"               '      Or Approved (status=2)
                cmdVerify.Enabled = False
                Call FillGridFromExtract
            ElseIf mStatusFlag = -1 Then
                cmdVerify.Enabled = False
                lblmsg.Caption = "Please Verify the SourceOf Fund Opening Balance of 2014-15"
            End If
        End If
    End Sub
    
    Public Property Let CheckPrintStatus(mData As Variant)
        mCheckPrintStatus = mData
    End Property
    
    Public Property Get CheckPrintStatus() As Variant
        CheckPrintStatus = mCheckPrintStatus
    End Property
    
    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If gbSeatGroupID = gbSeatGroupSecretary Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            If Col = 3 Then
                Cancel = False
                If vsGrid.TextMatrix(Row, 3) = "" Then
                    vsGrid.TextMatrix(Row, 3) = vsGrid.TextMatrix(Row, 2)
                End If
            Else
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End Sub

Private Sub vsGrid_LeaveCell()
    If gbSeatGroupID = gbSeatGroupSecretary Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            If vsGrid.Col = 3 Then
                
                If vsGrid.TextMatrix(vsGrid.Row, 3) <> "" Then
                    If val(vsGrid.TextMatrix(vsGrid.Row, 3)) = val(vsGrid.TextMatrix(vsGrid.Row, 2)) Then
                        vsGrid.TextMatrix(vsGrid.Row, 3) = vbNullString
                    End If
                End If
            
            End If
       End If
End Sub
