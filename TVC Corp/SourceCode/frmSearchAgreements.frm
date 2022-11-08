VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearchAgreements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SearchAgreements"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmSearchAgreements.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   5100
      TabIndex        =   13
      Top             =   1680
      Width           =   1020
   End
   Begin VB.TextBox txtWorkTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   10
      Top             =   840
      Width           =   6975
   End
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1170
      Width           =   4935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
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
      Left            =   4050
      TabIndex        =   8
      Top             =   1680
      Width           =   1020
   End
   Begin VB.TextBox txtProjectNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1170
      Width           =   1995
   End
   Begin VB.TextBox txtAgreementDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   5
      Top             =   510
      Width           =   1725
   End
   Begin VB.CommandButton cmdProject 
      Caption         =   "..."
      Height          =   270
      Left            =   8670
      TabIndex        =   2
      Top             =   1170
      Width           =   285
   End
   Begin VB.TextBox txtAgreementNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   150
      Width           =   1725
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   2160
      Width           =   9795
      _cx             =   17277
      _cy             =   5424
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchAgreements.frx":1CCA
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
   Begin MSComCtl2.DTPicker dtpAgreementDate 
      Height          =   315
      Left            =   3360
      TabIndex        =   12
      Top             =   510
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59244545
      CurrentDate     =   40087
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work Title:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   90
      TabIndex        =   11
      Top             =   870
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AgreementNo:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   330
      TabIndex        =   7
      Top             =   180
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agreement Date:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   555
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project No:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   510
      TabIndex        =   3
      Top             =   1200
      Width           =   870
   End
End
Attribute VB_Name = "frmSearchAgreements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
    Private Sub cmdClear_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdProject_Click()
        frmSulekhaIntegration.Show vbModal
        txtProjectNo.Tag = gbProject.decProjectID
        txtProjectNo.Text = gbProject.chvProjectSlNo
        txtProjectName.Text = gbProject.chvProjectName
    End Sub
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        vsGrid.Clear 1, 1
    End Sub
    Private Sub cmdSearch_Click()
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim mRowCount   As Integer
        Dim mDate       As String
        If txtAgreementDate.Text = "" Then
            mDate = Format(gbTransactionDate, "DD/MMM/YYYY")
        Else
            mDate = Format(txtAgreementDate.Text, "DD/MMM/YYYY")
        End If
        mSQL = "Select faAgreements.*,chvProjectName,vchTitle,vchSubLedgerType,chvProjectNameEnglish From faAgreements "
        mSQL = mSQL + " Inner Join suProjectDetails ON faAgreements.numProjectID=suProjectDetails.decProjectID"
        mSQL = mSQL + " Inner join faSubSidiaryAccountHeads ON faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID=faAgreements.numSubledgerID"
        mSQL = mSQL + " Inner Join faSubLedgerTypes On faSubLedgerTypes.intSubLedgerTypeID=faAgreements.intSubLedgerTypeID"
        mSQL = mSQL + " Where vchAgreementNo like '%" & txtAgreementNo & "%'"
        mSQL = mSQL + " And vchWorkTitle Like '%" & txtWorkTitle.Text & "%'"
        mSQL = mSQL + " And numProjectID Like '%" & txtProjectNo.Tag & "%'"
        mSQL = mSQL + " And dtAgreementDate <='" & mDate & "'"
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
        mRowCount = 1
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAgreementNo), "", Rec!vchAgreementNo)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchWorkTitle), "", Rec!vchWorkTitle)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchSubLedgerType), "", Rec!vchSubLedgerType)
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!dtDueDateToStart), "", Rec!dtDueDateToStart)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!dtActualStartedDate), "", Rec!dtActualStartedDate)
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltPAC), "", Rec!fltPAC)
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
            vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intAgreementID), "", Rec!intAgreementID)
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
            AutoWordWrap vsGrid
    End Sub

    Private Sub dtpAgreementDate_CloseUp()
        txtAgreementDate.Text = Format(dtpAgreementDate.value, "DD/MMM/YYYY")
    End Sub
    
    Private Sub Form_Load()
        Call FormInitialize
        Me.Left = (frmMenu.Width - Me.Width - 200) / 2
        Me.Top = (frmMenu.Height - Me.Height - 1700) / 2
        
        Call FillGrid
    End Sub
    Private Sub FillGrid()
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim mRowCount   As Integer
        Dim mDate       As Date
        mSQL = "Select faAgreements.*,chvProjectName,vchTitle,vchSubLedgerType,chvProjectNameEnglish From faAgreements "
        mSQL = mSQL + " Inner Join suProjectDetails ON faAgreements.numProjectID=suProjectDetails.decProjectID"
        mSQL = mSQL + " Inner join faSubSidiaryAccountHeads ON faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID=faAgreements.numSubledgerID"
        mSQL = mSQL + " Inner Join faSubLedgerTypes On faSubLedgerTypes.intSubLedgerTypeID=faAgreements.intSubLedgerTypeID"
        mSQL = mSQL + " where faAgreements.tnyStatus=1"
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
        mRowCount = 1
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAgreementNo), "", Rec!vchAgreementNo)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchWorkTitle), "", Rec!vchWorkTitle)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchSubLedgerType), "", Rec!vchSubLedgerType)
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!dtDueDateToStart), "", Rec!dtDueDateToStart)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!dtActualStartedDate), "", Rec!dtActualStartedDate)
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltPAC), "", Rec!fltPAC)
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
            vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intAgreementID), "", Rec!intAgreementID)
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
            AutoWordWrap vsGrid
    End Sub
    Function AutoWordWrap(vs As VSFlexGrid)
        With vs
            .AutoSizeMode = flexAutoSizeRowHeight
            .WordWrap = True
            .AutoSize 0, .Cols - 1
        End With
    End Function
    Private Sub txtAgreementDate_LostFocus()
         txtAgreementDate.Text = DdMmmYy(gbTransactionDate)
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 0)
            gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 8)
            Unload Me
        End If
    End Sub
