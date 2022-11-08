VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmZonalMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "frmZonalMain"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   Tag             =   "0"
   Begin VB.Frame fraRebuildIndex 
      Height          =   1035
      Left            =   120
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   13785
      Begin VB.CommandButton cmdIndexing 
         Caption         =   "Rebuild Index"
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
         Left            =   12060
         TabIndex        =   23
         Top             =   450
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   390
         TabIndex        =   22
         Top             =   540
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   14265
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZONAL OFFICES"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   255
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Fram 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   945
      Width           =   14145
      Begin VB.CommandButton cmdPre 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11160
         TabIndex        =   8
         Top             =   330
         Width           =   495
      End
      Begin VB.TextBox txtTrnDate 
         Height          =   345
         Left            =   11655
         TabIndex        =   7
         Top             =   330
         Width           =   1515
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13185
         TabIndex        =   6
         Top             =   330
         Width           =   495
      End
      Begin VB.TextBox txtTotalAmt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6375
         TabIndex        =   4
         Top             =   5160
         Width           =   2175
      End
      Begin VSFlex8LCtl.VSFlexGrid VSGridZonal 
         Height          =   4140
         Left            =   150
         TabIndex        =   3
         Top             =   900
         Width           =   13875
         _cx             =   24474
         _cy             =   7302
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalMain.frx":0000
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5580
         TabIndex        =   5
         Top             =   5130
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   0
      TabIndex        =   9
      Top             =   6510
      Width           =   14190
      Begin VB.Frame Frame3 
         Height          =   1410
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   14055
         Begin VB.Frame Frame2 
            Caption         =   "fraMissing Reports"
            Height          =   615
            Left            =   240
            TabIndex        =   24
            Top             =   750
            Width           =   6225
            Begin VB.CommandButton cmdMissing 
               Caption         =   "VIEW"
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
               Left            =   5430
               TabIndex        =   29
               Top             =   180
               Width           =   735
            End
            Begin VB.ComboBox cmbZonal 
               Height          =   315
               Left            =   3120
               TabIndex        =   28
               Text            =   "Zonal"
               Top             =   240
               Width           =   2310
            End
            Begin VB.TextBox txtDate 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Left            =   465
               TabIndex        =   25
               Top             =   210
               Width           =   1515
            End
            Begin MSComCtl2.DTPicker dtpdate 
               Height          =   360
               Left            =   1995
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   210
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   635
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
               Format          =   16515073
               CurrentDate     =   43509
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Zonal"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   2490
               TabIndex        =   30
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Date"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   240
               Left            =   60
               TabIndex        =   27
               Top             =   255
               Width           =   360
            End
         End
         Begin VB.CommandButton cmdRebuildIndex 
            Caption         =   "Rebuild Index"
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
            Left            =   6960
            TabIndex        =   20
            Top             =   390
            Width           =   3105
         End
         Begin VB.TextBox txtFromDate 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   645
            TabIndex        =   14
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtToDate 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   2715
            TabIndex        =   13
            Top             =   360
            Width           =   1515
         End
         Begin VB.CommandButton cndView 
            Caption         =   "VIEW"
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
            Left            =   4680
            TabIndex        =   12
            Top             =   330
            Width           =   735
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "CLOSE"
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
            Left            =   12780
            TabIndex        =   11
            Top             =   345
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   360
            Left            =   2175
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   360
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   635
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
            Format          =   16515073
            CurrentDate     =   39612
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   360
            Left            =   4245
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   360
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   635
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
            Format          =   16515073
            CurrentDate     =   39612
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zonal Collection Report"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1725
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&From"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   240
            Left            =   240
            TabIndex        =   18
            Top             =   405
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&To"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   240
            Left            =   2490
            TabIndex        =   17
            Top             =   480
            Width           =   180
         End
      End
   End
End
Attribute VB_Name = "frmZonalMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mStatus, mLoop As Integer
Dim mFromDate, mToDate As Date
Dim mnumZonID, mZonTotal, mintVerifyStatus As Variant


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdIndexing_Click()
   Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New Recordset
    
    objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    objdb.ExecuteSP "spDBREINDEX", , , , mCnn
    MsgBox " Successfully REBUILED", vbApplicationModal
    
End Sub

Private Sub cmdMissing_Click()

    
     Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim mZone As Double
        Dim frmNewViewer As New frmRptViewer

            If CDate(txtDate.Text) Then
   
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpFrom.Value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub

            Else
                MsgBox "Please enter Date", vbInformation
                Exit Sub
            End If
            If cmbZonal.ItemData(cmbZonal.ListIndex) > 1 Then
                mZone = cmbZonal.ItemData(cmbZonal.ListIndex)
            Else
                MsgBox "Please enter Date", vbInformation
                Exit Sub
            End If
                arInput = Array(CDate(txtFromDate.Text), mZone)
                frmNewRpt.rptFileName = App.Path & "\Reports\rptMissingReceipt.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
End Sub

Private Sub cmdPost_Click()
    Dim mDatepost As Date

    mDatepost = Format(txtTrnDate.Text, "dd-mmm-yyyy")
    mDatepost = Format(DateAdd("d", 1, mDatepost), "dd-mmm-yyyy")
    txtTrnDate.Text = mDatepost
    Call FillGrid
    
End Sub

Private Function chkChangeDate(mdtChange As Date)
'     Dim mdtCurr As Date
'     'mdtCurr = Format(txtDate.Text, "dd-mmm-yyyy")
'     mdtCurr = Format(txtTrnDate.Tag, "dd-mmm-yyyy")
'     If mdtCurr >= mdtChange Then
'        txtTrnDate.Text = mdtChange
'    Else
'        MsgBox "Please Transfer Current period", vbApplicationModal, "Saankhya"
'        txtTrnDate.Text = mdtCurr
'    End If
End Function

Private Sub cmdPre_Click()
    Dim mDatepre As Date
    mDatepre = Format(txtTrnDate.Text, "dd-mmm-yyyy")
    mDatepre = Format(DateAdd("d", -1, mDatepre), "dd-mmm-yyyy")
    txtTrnDate.Text = mDatepre
    'Call chkChangeDate(mDatepre)
    Call FillGrid
End Sub

Private Sub cmdRebuildIndex_Click()
    If fraRebuildIndex.Visible = True Then
        fraRebuildIndex.Visible = False
        cmdRebuildIndex.Caption = "Rebuild Index to Show"
    Else
        fraRebuildIndex.Visible = True
        cmdRebuildIndex.Caption = "Rebuild Index to Hide"
    End If
End Sub

Private Sub cndView_Click()

        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer

            If CDate(txtFromDate.Text) Then
                If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpFrom.Value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter From date", vbInformation
                Exit Sub
            End If

                arInput = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptZonalContaDetails.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Unload Me
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
End Sub





Private Sub dtpdate_CloseUp()
    txtDate.Text = DdMmmYy(dtpdate.Value)
End Sub

Private Sub dtpFrom_CloseUp()
     txtFromDate.Text = DdMmmYy(dtpFrom.Value)
End Sub



Private Sub dtpToDate_CloseUp()
     txtToDate.Text = DdMmmYy(dtpToDate.Value)
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
    Me.Height = frmMenu.Height - 1590
    Call FormIntialize
    Call FillGrid
    Call PopulateList(cmbZonal, "Select chvZoneNameEnglish, numZoneID From GM_Zone WHERE Right(numZoneID,2)<>1 AND intLBID =" & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster)
End Sub



Private Sub txtTrnDate_LostFocus()
    Call chkChangeDate(Format(txtTrnDate.Text, "dd-mmm-yyyy"))
End Sub

Private Sub VSGridZonal_DblClick()
    If VSGridZonal.Row > 0 Then
        frmZonalDetails.ZonalDetailsID = VSGridZonal.TextMatrix(VSGridZonal.Row, 5)
        frmZonalDetails.ZonalName = VSGridZonal.TextMatrix(VSGridZonal.Row, 1)
        frmZonalDetails.ZonalTrnDate = VSGridZonal.TextMatrix(VSGridZonal.Row, 0)
        frmZonalDetails.Show
    End If
End Sub

Private Sub FormIntialize()
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim Rec As New Recordset
    Dim mRowCount As Variant
    Dim aryInput, mArr As Variant
    Dim mLoop, mVouFin, mVouHO As Integer
    Dim mFromDate, mToDate, mMaxDate As Variant
    

    objdb.CreateNewConnection mCnn, SaankhyaHO
    
    mSql = "SELECT  ISNULL(MIN(dtDate),'01-Apr-2014') AS dtDate FROM faVouchers "
    mSql = mSql + " INNER JOIN faSyncLog ON faVouchers.intVoucherID=faSyncLog.intVoucherID"
    mSql = mSql + " AND faVouchers.numLocationID= faSyncLog.intLocationID Where tnySyncStatus = 1"
       
    Rec.Open mSql, mCnn
    
    If IsDate(Rec!dtDate) Then
        txtTrnDate.Text = Rec!dtDate
        txtTrnDate.Tag = Rec!dtDate
    End If
    dtpFrom.Value = gbTransactionDate
    dtpToDate.Value = gbTransactionDate
    txtFromDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
    txtToDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
    Rec.Close
    mCnn.Close
    'Call FillGrid
End Sub

Private Function FillGrid()
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim Rec As New Recordset
    Dim mRowCount As Variant
    Dim aryInput, aryOut As Variant
    Dim mLoop As Integer
    Dim mDateMax As Date
    Dim mNoVou, mNotTranVouNo, mDiff As Integer
    Dim mRecCount As Integer
    
    objdb.CreateNewConnection mCnn, SaankhyaHO
    
    aryInput = Array(DdMmmYy(txtTrnDate.Text))
    Set Rec = objdb.ExecuteSP("spGetTransactionSummaryofZonals", aryInput, , , mCnn)
    
    mRowCount = 1
    VSGridZonal.Rows = 1
    
    While Not Rec.EOF
        If Rec!tnyZoneNo <> 1 Then
            VSGridZonal.Rows = VSGridZonal.Rows + 1
            VSGridZonal.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            VSGridZonal.TextMatrix(mRowCount, 1) = UCase(IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish))
            VSGridZonal.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!fltTotal), "", Format(Rec!fltTotal, "0.00  "))
            VSGridZonal.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!intNoVouchers), "", Rec!intNoVouchers) & "  "
            VSGridZonal.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
            VSGridZonal.TextMatrix(mRowCount, 6) = ""
            mRowCount = mRowCount + 1
        End If
        Rec.MoveNext
    Wend
    Rec.Close
    
    '=========================================================
    ' TO UPDATE THE STATUS
    '=========================================================
    aryInput = Array(DdMmmYy(txtTrnDate.Text))
    Set Rec = objdb.ExecuteSP("spGetTransferStatusofZonals", aryInput, , , mCnn)
    If Not (Rec.EOF And Rec.BOF) Then
        While Not Rec.EOF
            For mLoop = 1 To VSGridZonal.Rows - 1
                If VSGridZonal.TextMatrix(mLoop, 5) = Rec!intLocationID Then
                    If Rec!VoucherCount = Rec!TransferedCount Then
                        VSGridZonal.TextMatrix(mLoop, 6) = "TRANSFERED"
                    ElseIf Rec!TransferedCount > 0 Then
                        VSGridZonal.TextMatrix(mLoop, 6) = "TRANSFER NOT COMPLETED"
                    Else
                        VSGridZonal.TextMatrix(mLoop, 6) = "NOT TRANSFERED"
                    End If
                End If
            Next mLoop
            Rec.MoveNext
        Wend
    End If
    Rec.Close
    mCnn.Close
    Call Calculation
    
    '    For mLoop = 1 To VSGridZonal.Rows - 1
    '        ' READING RECORD COUNTS FROM ZONAL DATA
    '        aryInput = Array(DdMmmYy(txtTrnDate.Text), val(VSGridZonal.TextMatrix(mLoop, 5)))
    '        aryOut = Null
    '        objDB.ExecuteSP "spGetTransferStatusofZonals", aryInput, aryOut, , mCnn
    '        If Not IsNull(aryOut) Then
    '            VSGridZonal.TextMatrix(mLoop, 6) = aryOut(1, 0)
    '        End If
    '    Next mLoop
    'mCnn.Close
    'Call Calculation(2, txtTotalAmt)

End Function

Private Sub Calculation()
    Dim mLoop As Integer
    mZonTotal = 0
    For mLoop = 1 To VSGridZonal.Rows - 1
        mZonTotal = mZonTotal + val(VSGridZonal.TextMatrix(mLoop, 2))
    Next
    txtTotalAmt.Text = Format(mZonTotal, "0.00 ")
End Sub

Public Property Let ZonalID(mData As Variant)
    mnumZonID = mData
End Property

Public Property Get ZonalID() As Variant
    ZonalID = mnumZonID
End Property
