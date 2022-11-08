VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmImplementingOfficerList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Implementing Officer"
   ClientHeight    =   7155
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   11025
      Top             =   7770
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   6270
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   13710
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "  List of Implementing Officer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   420
         Left            =   0
         TabIndex        =   3
         Top             =   135
         Width           =   13815
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5715
      Left            =   45
      TabIndex        =   0
      Top             =   795
      Width           =   13680
      _cx             =   24130
      _cy             =   10081
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
      HighLight       =   0
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImplementingOfficerList.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
Attribute VB_Name = "frmImplementingOfficerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
    frmImpOfficer.Form_Load
    frmImpOfficer.Show vbModal
End Sub
Private Sub FillGrid()
        Dim objdb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSql As String
        vsGrid.Rows = 1
        vsGrid.TextMatrix(0, 1) = "Sl.No"
        vsGrid.TextMatrix(0, 2) = "Implementing Officer"
        vsGrid.TextMatrix(0, 3) = "Title"
        vsGrid.TextMatrix(0, 4) = "Designation"
     
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select intSubsidiaryAccountHeadID,intSubLedgerTypeID,vchSubLedgerCode,vchTitle,vchName,vchDesignation,tnySyncFlag from faSubsidiaryAccountHeads Where intSubLedgerTypeID=1"
        mSql = mSql + " Order By  intSubsidiaryAccountHeadID desc"
        Rec.Open mSql, mCnn
        mRow = 1
        If Not (Rec.BOF And Rec.EOF) Then
             While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
                vsGrid.TextMatrix(mRow, 1) = mRow
                vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
                vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
                If Rec!tnySyncFlag = 0 Or Rec!tnySyncFlag = 1 Then
                     vsGrid.Cell(flexcpBackColor, mRow, 0, , 4) = &HC0FFC0
                End If
                If Rec!tnySyncFlag = 2 Then
                     vsGrid.Cell(flexcpBackColor, mRow, 0, , 4) = &H80C0FF
                End If


                Rec.MoveNext
                mRow = mRow + 1
             Wend
        End If
End Sub

Private Sub Form_Activate()
    Me.Top = 1300
    Me.Left = 0
 Call FillGrid
End Sub

Private Sub Form_Load()
  '  Call fillGrid
    WindowsXPC1.InitIDESubClassing
End Sub

Private Sub vsGrid_DblClick()
    If val(vsGrid.TextMatrix(vsGrid.Row, 0)) > 0 Then
            Call frmImpOfficer.fillimpOfficer(val(vsGrid.TextMatrix(vsGrid.Row, 0)))
            frmImpOfficer.Form_Load
            frmImpOfficer.Show vbModal
    End If
End Sub
