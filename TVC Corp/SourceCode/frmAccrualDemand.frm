VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccrualDemand 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accrual Demand Register"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   135
      Left            =   1590
      TabIndex        =   3
      Top             =   5595
      Visible         =   0   'False
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEE6E6&
      Caption         =   "Cance&L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5955
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5970
      Width           =   1410
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00CEE6E6&
      Caption         =   "&Posting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5970
      Width           =   1380
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4620
      Left            =   195
      TabIndex        =   0
      Top             =   720
      Width           =   11460
      _cx             =   20214
      _cy             =   8149
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      BackColorFixed  =   13559526
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14349042
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   3
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAccrualDemand.frx":0000
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
End
Attribute VB_Name = "frmAccrualDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillGrid()
    '---------------------------------------------------------------------------'
    '                                                                           '
    '---------------------------------------------------------------------------'
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim mLoop As Long
    Dim mDemandID As Variant
    Dim arrInput As Variant
    vsGrid.Clear 1, 1
    objDB.SetConnection mCnn
    
    arrInput = Array(gbTransactionDate)
    Set Rec = objDB.ExecuteSP("spGetAccruedDemands")
    If Not (Rec.BOF And Rec.EOF) Then
        While Not Rec.EOF
            If mDemandID <> Rec!numDemandID Then
                mDemandID = Rec!numDemandID
                mLoop = mLoop + 1
                vsGrid.TextMatrix(mLoop, 0) = mLoop
                vsGrid.TextMatrix(mLoop, 1) = Rec!dtDueDate
                vsGrid.TextMatrix(mLoop, 2) = Rec!vchTransactionType
                vsGrid.TextMatrix(mLoop, 3) = Rec!fltAmount
                vsGrid.TextMatrix(mLoop, 4) = 0
                vsGrid.TextMatrix(mLoop, 5) = Rec!numDemandID
            End If
            Rec.MoveNext
        Wend
    End If
    
End Sub
Private Sub cmdCancel_Click()
    FillGrid
End Sub

Private Sub cmdSave_Click()
    Dim mLoop As Long
    For mLoop = 1 To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpChecked, mLoop, 4) = vbChecked Then
            'Debug.Print vsGrid.Cell(flexcpText, mLoop, 5)
            Call AccrualJournalByDemandID(vsGrid.Cell(flexcpText, mLoop, 5))
            
        End If
    Next mLoop
    ProgressBar.Min = 0
    ProgressBar.Max = 10000
    ProgressBar.Visible = True
    For mLoop = 1 To 10000
        ProgressBar.Value = ProgressBar.Value + 1
    Next mLoop
    ProgressBar.Visible = False
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
End Sub
Private Sub Form_Load()
    Call FillGrid
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 4 Then
        Cancel = True
    ElseIf Col = 4 Then
        If Val(vsGrid.Cell(flexcpText, Row, 3)) <= 0 Then
            Cancel = True
        End If
    End If
    
End Sub

