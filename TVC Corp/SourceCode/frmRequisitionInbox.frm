VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRequisitionInbox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisition Inbox"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   16155
   Icon            =   "frmRequisitionInbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   16155
   Begin VB.Frame Frame2 
      Height          =   1740
      Left            =   11520
      TabIndex        =   22
      Top             =   6975
      Width           =   4575
      Begin VB.CommandButton cmdSync 
         Caption         =   "<<SYNC>>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   18240
      Top             =   10440
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   16020
      TabIndex        =   20
      Top             =   0
      Width           =   16020
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   720
         TabIndex        =   21
         Top             =   120
         Width           =   15000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   0
      TabIndex        =   1
      Top             =   6930
      Width           =   16095
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   19
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearchSourceID 
         Caption         =   "..."
         Height          =   285
         Left            =   6915
         TabIndex        =   17
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchIMPO 
         Caption         =   "..."
         Height          =   285
         Left            =   6915
         TabIndex        =   16
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtSourceOfFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtIMPO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   4815
      End
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtTokenNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtProjectNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   10860
         TabIndex        =   9
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   17760257
         CurrentDate     =   40197
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   10860
         TabIndex        =   13
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   17760257
         CurrentDate     =   40197
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Source Of Fund"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8235
         TabIndex        =   14
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   10
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8205
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Token No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Project No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Implementing Officer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6135
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   15975
      _cx             =   28178
      _cy             =   10821
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRequisitionInbox.frx":1CCA
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
Attribute VB_Name = "frmRequisitionInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mTimer  As Integer
    
    Private Sub FillGrid()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mRowCnt As Integer
        Dim mLoop As Integer
        
        On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQL = " SELECT * FROM faRequisitionInbox"
        mSQL = mSQL + " LEFT JOIN suImplementingOfficer ON suImplementingOfficer.intImplementingOfficerID=faRequisitionInbox.intImplementingOfficersID"
        mSQL = mSQL + " INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faRequisitionInbox.intSourceID"
        mSQL = mSQL + " INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faRequisitionInbox.intFundCategoryID"
        'msql = msql + " WHERE tnyStage=0"
        
        If txtFromDate.Text <> "" And txtToDate.Text <> "" Then
            mSQL = mSQL + " AND dtRequestDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "'"
        End If
        If txtProjectNo.Text <> "" Then
            mSQL = mSQL + " AND vchProjectNo LIKE   '%" & Trim(txtProjectNo.Text) & "%' "
        End If
        If txtSourceOfFund.Text <> "" Then
            mSQL = mSQL + " AND faRequisitionInbox.intSourceID= " & val(txtSourceOfFund.Tag) & " "
        End If
        If txtTokenNo.Text <> "" Then
            mSQL = mSQL + " AND faRequisitionInbox.intTokenID= " & val(txtTokenNo.Tag) & " "
        End If
        
        If cmbMonth.ListIndex <> -1 Then
            mSQL = mSQL + " AND Month(dtRequestDate)= " & cmbMonth.ItemData(cmbMonth.ListIndex) & " "
        End If
        
        Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        While Not (Rec.EOF Or Rec.BOF)
            vsGrid.Rows = vsGrid.Rows + 1
            
            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchImplementingOfficer), "", Rec!vchImplementingOfficer)
            vsGrid.TextMatrix(mRowCnt, 1) = IIf(Rec!intRequisitionNo = 0, "", Rec!intRequisitionNo)
            If Rec!dtRequestDate <> "" Then
                vsGrid.TextMatrix(mRowCnt, 2) = DdMmmYy(IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate))
            Else
                vsGrid.TextMatrix(mRowCnt, 2) = ""
            End If
            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo)
            vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
            vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!fltRequestedAmt), "", Rec!fltRequestedAmt)
            vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!intTockenNo), "", Rec!intTockenNo)
            
            If Rec!tnyStage = 0 Then
                vsGrid.TextMatrix(mRowCnt, 8) = "Request Waiting" 'SYNCED FROM WEB
            ElseIf Rec!tnyStage = 1 Then
                vsGrid.TextMatrix(mRowCnt, 8) = "Verified"        'SAVED TO faAllotments[Req No Generated]
            ElseIf Rec!tnyStage = 2 Then
                vsGrid.TextMatrix(mRowCnt, 8) = "Returned to Secretary" 'WRITE BACK TO WEB
            ElseIf Rec!tnyStage = 3 Then
                vsGrid.TextMatrix(mRowCnt, 8) = "Return" 'APPROVED FROM WEB AND RETURNED BACK TO LOCAL
            End If
            
            If Rec!tnyStatus = 2 Then  'IF REQ IS CANCELLED
                For mLoop = 0 To vsGrid.Cols - 1
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HC0E0FF
                Next mLoop
                  
            End If
  
            vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intImplementingOfficersID), 0, Rec!intImplementingOfficersID)
            vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
            vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID)
            
            vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!tnyStage), 0, Rec!tnyStage)
            vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
            vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!intID), 0, Rec!intID)
            
            Rec.MoveNext
            mRowCnt = mRowCnt + 1
            Wend
            Rec.Close
            mCnn.Close
        Exit Sub
Err:
        MsgBox Err.Description
    
    End Sub
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    
    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub

    Private Sub cmdSearchIMPO_Click()
        gbSearchID = -1                                         ''  Setting the Search ID to -1
        frmSearchSubsidiaryAccountHeads.SubLedgerType = 1       ''  1. Implementing Officer
        frmSearchSubsidiaryAccountHeads.Show vbModal
        txtIMPO.SetFocus
    End Sub

    Private Sub cmdSearchSourceID_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtSourceOfFund.Text = gbSearchStr
            txtSourceOfFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdSync_Click()
        Dim mArrIN As Variant
        Dim mArrOut As Variant
        Dim mUrl   As String
        Dim client1 As New MSSOAPLib.SoapClient
        Dim objSOAP As Variant
        Dim clnt As New SoapClient30
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL  As String
        Dim mReqIDCheck As Boolean
     
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
        '--------------'
        ' Web Service  '
        '--------------'
        mReqIDCheck = False
        mUrl = gbDefaultUrlForRequisition
        mArrIN = Array(gbLBID, gbFinancialYearID)
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
        
        Dim i As Integer
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                    mSQL = "SELECT * FROM faRequisitionInbox WHERE intID=" & Rec!intID & ""
                    
                    RecID.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                    If Not (RecID.BOF And RecID.EOF) Then
                        mReqIDCheck = True
                    Else
                        mReqIDCheck = False
                    End If
                    RecID.Close
            
                    If mReqIDCheck = False Then
''''                        mSql = ""
''''                        mSql = "INSERT INTO faRequisitionInbox "
''''                        mSql = mSql + "   (intID,intTockenNo,intSlNo, intRequisitionNo, dtRequestDate, intImplementingOfficersID, " & vbNewLine
''''                        mSql = mSql + "    intImpoUserID, tnyPlanOrNonPlan, numProjectID,vchProjectNo, fltRequestedAmt, nvchDescription," & vbNewLine
''''                        mSql = mSql + "    intSourceID , intFundCategoryID, intSchemeID, intSubSecID, intMircoSectorID, vchDPCApprovalNo," & vbNewLine
''''                        mSql = mSql + "    dtDPCDate, intTreasuryID, intLBID, intFinancialYearID, tnyStage, tnyStatus, tnySyncFlag)" & vbNewLine
''''                        mSql = mSql + "     VALUES   " & vbNewLine
''''                        mSql = mSql + "   (" & Rec!intID & " ," & Rec!intTockenNo & " ," & Rec!intSlNo & " ," & IIf(IsNull(Rec!intRequisitionNo), 0, Rec!intRequisitionNo) & " , " & vbNewLine
''''                        mSql = mSql + " '" & DdMmmYy(Rec!dtRequestDate) & "', " & Rec!intImplementingOfficersID & " , " & Rec!intImpoUserID & " , " & Rec!tnyPlanOrNonPlan & " , " & Rec!numProjectID & " ," & vbNewLine
''''                        mSql = mSql + "    '" & Rec!vchProjectNo & "' , " & Rec!fltRequestedAmt & " , '" & Rec!nvchDescription & " ', " & Rec!intSourceID & " , " & vbNewLine
''''                        mSql = mSql + "  " & Rec!intFundCategoryID & " , " & Rec!intSchemeID & " , " & Rec!intSubSecID & " , " & Rec!intMircoSectorID & " , '" & IIf(IsNull(Rec!vchDPCApprovalNo), 0, Rec!vchDPCApprovalNo) & " '," & vbNewLine
''''                        mSql = mSql + "    '" & DdMmmYy(Rec!dtDPCDate) & "' , " & Rec!intTreasuryID & " , " & Rec!intLBID & " , " & vbNewLine
''''                        mSql = mSql + " " & Rec!intFinancialYearID & " ," & Rec!tnyStage & " , " & Rec!tnyStatus & " ," & IIf(IsNull(Rec!tnySyncFlag), 0, Rec!tnySyncFlag) & " )" & vbNewLine
                        
                        mArrIN = Array(Rec!intID, _
                                    Rec!intTockenNo, _
                                    Rec!intSlNo, _
                                    IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate), _
                                    IIf(IsNull(Rec!intImplementingOfficersID), 0, Rec!intImplementingOfficersID), _
                                    IIf(IsNull(Rec!intImpoUserID), 0, Rec!intImpoUserID), _
                                    IIf(IsNull(Rec!tnyPlanOrNonPlan), "", Rec!tnyPlanOrNonPlan), _
                                    IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID), _
                                    IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo), _
                                    IIf(IsNull(Rec!fltRequestedAmt), 0, Rec!fltRequestedAmt), _
                                    IIf(IsNull(Rec!nvchDescription), "", Rec!nvchDescription), _
                                    IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID), _
                                    IIf(IsNull(Rec!intFundCategoryID), 0, Rec!intFundCategoryID), _
                                    IIf(IsNull(Rec!intSchemeID), 0, Rec!intSchemeID), _
                                    IIf(IsNull(Rec!intSubSecID), 0, Rec!intSubSecID), _
                                    IIf(IsNull(Rec!intMircoSectorID), 0, Rec!intMircoSectorID), _
                                    IIf(IsNull(Rec!vchDPCApprovalNo), "", Rec!vchDPCApprovalNo), _
                                    IIf(IsNull(Rec!dtDPCDate), "", Rec!dtDPCDate), _
                                    IIf(IsNull(Rec!intTreasuryID), 0, Rec!intTreasuryID), _
                                    IIf(IsNull(Rec!intLBID), 0, Rec!intLBID), _
                                    IIf(IsNull(Rec!intFinancialYearID), 0, Rec!intFinancialYearID), _
                                    Rec!tnyStage, Rec!tnyStatus, _
                                    IIf(IsNull(Rec!tnySyncFlag), 0, Rec!tnySyncFlag) _
                                    )
                            objDB.ExecuteSP "spSaveRequisitionInbox", mArrIN, , , mCnn, adCmdStoredProc
                   End If
                Rec.MoveNext
            Wend
        End If
        Rec.Close
        Call FillGrid
        Call SynTOWEB
    End Sub
    Private Sub SynTOWEB() '******************UPDATE STATUS,STAGE,REQUISITION NO TO DB_FinanceWEB******************
        Dim mArrIN As Variant
        Dim mArrOut As Variant
        Dim mUrl   As String
        Dim client1 As New MSSOAPLib.SoapClient
        Dim objSOAP As Variant
        Dim clnt As New SoapClient30
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL  As String
        Dim mLoop As Integer
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
        '--------------'
        ' Web Service  '
        '--------------'

        mUrl = gbDefaultUrlForRequisition
        'mArrIN = Array(mReqInboxID, gbLBID, txtRequisition.Text)
        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        objSOAP.mssoapinit mUrl + "?WSDL"

        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 1) <> "" Then
              mArrOut = objSOAP.SyncUpdateSyncFlagToRequisitionInbox(val(vsGrid.TextMatrix(mLoop, 14)), gbLBID, vsGrid.TextMatrix(mLoop, 1))
            
              mSQL = "Update faRequisitionInbox set  tnyStage=2 WHERE intID=" & val(vsGrid.TextMatrix(mLoop, 14)) & "  "
              objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
        Next mLoop
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0 '1350
        Call FillGrid
    End Sub

    Private Sub Form_Load()
        Call FillMonth
        Call FillGrid
        Call CheckConnectionStatus
    End Sub
    Private Sub FillMonth()
        'cmbMonth.AddItem "", 0
        'cmbMonth.ItemData(0) = 0
        
        cmbMonth.AddItem "April", 0
        cmbMonth.ItemData(0) = 4
        
        cmbMonth.AddItem "May", 1
        cmbMonth.ItemData(1) = 5
        
        cmbMonth.AddItem "June", 2
        cmbMonth.ItemData(2) = 6
        
        cmbMonth.AddItem "July", 3
        cmbMonth.ItemData(3) = 7
        
        cmbMonth.AddItem "August", 4
        cmbMonth.ItemData(4) = 8
        
        cmbMonth.AddItem "September", 5
        cmbMonth.ItemData(5) = 9
        
        cmbMonth.AddItem "October", 6
        cmbMonth.ItemData(6) = 10
        
        cmbMonth.AddItem "November", 7
        cmbMonth.ItemData(7) = 11
        
        cmbMonth.AddItem "December", 8
        cmbMonth.ItemData(8) = 12
        
        cmbMonth.AddItem "January", 9
        cmbMonth.ItemData(9) = 1
        
        cmbMonth.AddItem "February", 10
        cmbMonth.ItemData(10) = 2
        
        cmbMonth.AddItem "March", 11
        cmbMonth.ItemData(11) = 3
    End Sub
    Private Sub CheckConnectionStatus()
        
        Dim mArrOut As Variant
        Dim mUrl   As String
        Dim client1 As New MSSOAPLib.SoapClient
        Dim objSOAP As Variant
        Dim clnt As New SoapClient30
        Dim mSQL  As String

        On Error GoTo Err
        
        mUrl = gbDefaultUrlForRequisition
        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        objSOAP.mssoapinit mUrl + "?WSDL"

        mArrOut = objSOAP.Test(1, 2)
        If mArrOut = 3 Then
            lblStatus.Caption = "CONNECTION TO WEB SERVICE ESTABLISHED"
        End If
        Exit Sub
Err:
        'MsgBox err.Description
        lblStatus.Caption = "ERROR!!!CONNECTION COULD NOT BE ESTABLISHED"
    End Sub
    Private Sub dtpDateFrom_CloseUp()
        txtFromDate.Text = dtpDateFrom.value
        txtFromDate.SetFocus
    End Sub
    
    Private Sub dtpDateTo_CloseUp()
        txtToDate.Text = dtpDateTo.value
        txtToDate.SetFocus
    End Sub
    
    Private Sub Form_Paint()
        Call FillGrid
    End Sub

    Private Sub Timer1_Timer()
        If mTimer = 0 Then
            mTimer = 1
            lblStatus.Visible = True
        ElseIf mTimer = 1 Then
            mTimer = 0
            lblStatus.Visible = False
        End If
    End Sub

    Private Sub txtFromDate_LostFocus()
        If txtFromDate.Text <> "" Then
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
            If CDate(txtFromDate.Text) < CDate(gbStartingDate) Then
                If CDate(txtFromDate.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                    txtFromDate.Text = DateAdd("yyyy", -1, gbStartingDate)
                    txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
                End If
                txtToDate.Text = DateAdd("yyyy", -1, gbEndingDate)
                txtToDate.Text = CheckDateInMMM(txtToDate.Text)
            End If
        End If
    End Sub

    Private Sub txtIMPO_GotFocus()
        If gbSearchID > 0 Then
            Dim objSubLedger As New clsSubLedger
            objSubLedger.SetSubLedgerDetails (gbSearchID)
            If objSubLedger.SubsidiaryAccountHeadID Then
                txtIMPO.Tag = IIf(IsNull(objSubLedger.SubsidiaryAccountHeadID), 0, objSubLedger.SubsidiaryAccountHeadID)
                txtIMPO.Text = IIf(IsNull(objSubLedger.NameOfSubLedger), "", objSubLedger.NameOfSubLedger)
            Else
                txtIMPO.Tag = ""
                txtIMPO.Text = ""
            End If
        End If
        gbSearchID = -1
    End Sub

    Private Sub txtToDate_LostFocus()
        If txtToDate.Text <> "" Then
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
            If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
                MsgBox "Please Enter a Date less than Or Equal to To Date", vbInformation
                txtFromDate.Text = txtToDate.Text
                txtFromDate.SetFocus
                Exit Sub
            End If
        End If
    End Sub
    Private Sub vsGrid_DblClick()
        If gbSeatGroupID = gbSeatGroupChairPerson Then
            frmRequisitionAuthorisation.txtRequisitionNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 14)
            frmRequisitionAuthorisation.GetProjectAllotmentDetails
            frmRequisitionAuthorisation.Show vbModal
        Else
            If vsGrid.TextMatrix(vsGrid.Row, 12) = 0 Then
                frmRequisition.LoadMode = 20
                frmRequisition.ReqInboxID = (val(vsGrid.TextMatrix(vsGrid.Row, 14)))
                frmRequisition.TokenID = (val(vsGrid.TextMatrix(vsGrid.Row, 7)))
                frmRequisition.Show vbModal
            End If
        End If
    End Sub
