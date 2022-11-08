VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfAgreements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of Agreements"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frm1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   12615
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   12240
      Top             =   7560
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   12615
      _cx             =   22251
      _cy             =   9551
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfAgreements.frx":0000
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12615
      TabIndex        =   0
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmListOfAgreements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmdCancel_Click()
        Me.Hide
    End Sub

    Private Sub cmdNew_Click()
        frmAgreement.cmdSave.Enabled = True
        frmAgreement.txtCompletionDate.Enabled = False
        frmAgreement.Show vbModal
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        txtFromDate.Text = DdMmmYy(gbStartingDate)
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        Call FillGrid
    End Sub
    Public Sub FillGrid()
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mRowCnt As Integer
        

        If objDB.SetConnection(mCnn) Then
            mSQL = " Select faAgreements.* ,suProjectDetails.decProjectID,suProjectDetails.chvProjectSLNo,suProjectDetails.chvProjectNameEnglish"
            mSQL = mSQL + " from faAgreements left join suProjectDetails on faAgreements.numProjectID=suProjectDetails.decProjectID"
            mSQL = mSQL + " where faAgreements.dtAgreementDate BETWEEN '" & txtFromDate.Text & " '  AND '" & txtToDate.Text & " ' "
            Rec.CursorLocation = adUseClient
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            mRowCnt = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCnt, 0) = mRowCnt
                vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAgreementNo), "", Rec!vchAgreementNo)
                vsGrid.TextMatrix(mRowCnt, 2) = DdMmmYy(IIf(IsNull(Rec!dtAgreementDate), "", Rec!dtAgreementDate))
                vsGrid.TextMatrix(mRowCnt, 3) = DdMmmYy(IIf(IsNull(Rec!dtActualStartedDate), "", Rec!dtActualStartedDate))
                vsGrid.TextMatrix(mRowCnt, 4) = DdMmmYy(IIf(IsNull(Rec!dtDueDateToStart), "", Rec!dtDueDateToStart))
                'VSGrid.TextMatrix(mRowCnt, 5) = DdMmmYy(IIf(IsNull(Rec!dtActualCompletedDate), "", Rec!dtActualCompletedDate))
                vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!fltPAC), "", Rec!fltPAC)
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!vchOrderNo), "", Rec!vchOrderNo)
                vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
                vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!vchAssetName), "", Rec!vchAssetName)
                If (Rec!tnyStatus) = 1 Then
                    vsGrid.TextMatrix(mRowCnt, 10) = vbChecked
                Else
                    vsGrid.TextMatrix(mRowCnt, 10) = Unchecked
                End If
                Rec.MoveNext
                mRowCnt = mRowCnt + 1
            Wend
            Rec.Close
        End If
    End Sub
    Private Sub txtFromDate_LostFocus()
        If Not IsDate(txtFromDate.Text) Then
            txtFromDate.Text = DdMmmYy(gbStartingDate)
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub
    Private Sub txtToDate_LostFocus()
        If Not IsDate(txtToDate.Text) Then
            txtToDate.Text = DdMmmYy(gbTransactionDate)
        Else
            txtToDate.Text = CheckDateInMMM(Trim(txtToDate))
        End If
        
        If txtFromDate.Text <> "" Then
            Call FillGrid
        End If
    End Sub
    Private Sub vsGrid_DblClick()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSQL As String
        Dim marr As Variant
        
        If vsGrid.Row > 0 Then
            If vsGrid.TextMatrix(vsGrid.Row, 10) = 1 Then
                MsgBox "Approved!!!", vbInformation, "Saankhya"
                Exit Sub
            End If
            marr = Split(vsGrid.TextMatrix(vsGrid.Row, 1), "/")
            frmAgreement.txtAgreementNoPart1.Text = marr(0)
            frmAgreement.txtAgreementNoPart2.Text = marr(1)
            frmAgreement.txtAgreementDate.Text = DdMmmYy(vsGrid.TextMatrix(vsGrid.Row, 2))
            frmAgreement.txtComencementDate.Text = DdMmmYy(vsGrid.TextMatrix(vsGrid.Row, 3))
            frmAgreement.txtDueDateofCommencement.Text = DdMmmYy(vsGrid.TextMatrix(vsGrid.Row, 4))
            'frmAgreement.txtCompletionDate.Text = DdMmmYy(VSGrid.TextMatrix(VSGrid.Row, 5))
            frmAgreement.txtPAC.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
            frmAgreement.txtOrderNo.Text = vsGrid.TextMatrix(vsGrid.Row, 7)
            frmAgreement.txtProjectName.Text = vsGrid.TextMatrix(vsGrid.Row, 8)
            frmAgreement.txtAsset.Text = vsGrid.TextMatrix(vsGrid.Row, 9)
            
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSQL = "Select faAgreements.* ,suProjectDetails.decProjectID,suProjectDetails.chvProjectSLNo,suProjectDetails.chvProjectNameEnglish, faSubLedgerTypes.vchSubLedgerType "
            mSQL = mSQL + " from faAgreements left join suProjectDetails on faAgreements.numProjectID=suProjectDetails.decProjectID"
            mSQL = mSQL + " left join faSubLedgerTypes on faAgreements.intSubLedgerTypeID=faSubLedgerTypes.vchSubLedgerTypeCode"
            mSQL = mSQL + " where faAgreements.vchAgreementNo = '" & vsGrid.TextMatrix(vsGrid.Row, 1) & " '"
            Rec.Open mSQL, mCnn
            While Not (Rec.EOF Or Rec.BOF)
                frmAgreement.txtAgreementNoPart1.Tag = Rec!intAgreementID
                frmAgreement.txtProjectNo.Text = Rec!chvProjectSlNo
                frmAgreement.txtProjectNo.Tag = Rec!decProjectID
                frmAgreement.txtduedateofCompletion.Text = Rec!dtDueDateOfCompletion
                frmAgreement.txtWorkDate.Text = Rec!dtWorkDate
                frmAgreement.txtWorkTitle.Text = Rec!vchWorkTitle
                frmAgreement.txtContractors.Text = Rec!vchSubLedgerType
                frmAgreement.txtContractors.Tag = Rec!numSubLedgerID
                'frmAgreement.txtAsset = ""
                Rec.MoveNext
            Wend
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                frmAgreement.cmdApprove.Visible = True
                frmAgreement.cmdSave.Enabled = False
            End If
            frmAgreement.Show vbModal
        End If
        
    End Sub
