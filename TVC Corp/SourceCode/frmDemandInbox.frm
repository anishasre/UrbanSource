VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmDemandInbox 
   BackColor       =   &H00DDEDED&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand From Other Collection Points"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "frmDemandInbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   11415
      TabIndex        =   11
      Top             =   0
      Width           =   11415
   End
   Begin VB.ListBox lstZonals 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   1230
      TabIndex        =   9
      Top             =   990
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton cmdSearchLocation 
      Caption         =   "..."
      Height          =   285
      Left            =   3555
      TabIndex        =   8
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox txtToDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2010
      TabIndex        =   7
      Text            =   "99-WWW-0000"
      Top             =   5820
      Width           =   1275
   End
   Begin VB.TextBox txtFromDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   690
      TabIndex        =   6
      Text            =   "99-WWW-0000"
      Top             =   5820
      Width           =   1275
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9135
      TabIndex        =   4
      Top             =   720
      Width           =   1665
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00DDEDED&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   6165
      Width           =   11415
      Begin VB.CommandButton cmdSearchDemand 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10230
         TabIndex        =   10
         Top             =   60
         Width           =   915
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4380
      Left            =   0
      TabIndex        =   12
      Top             =   1245
      Width           =   11310
      _cx             =   19950
      _cy             =   7726
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
      BackColorBkg    =   16777215
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDemandInbox.frx":1CCA
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   150
      TabIndex        =   5
      Top             =   5835
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   8535
      TabIndex        =   3
      Top             =   750
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
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
      Left            =   210
      TabIndex        =   2
      Top             =   735
      Width           =   945
   End
   Begin VB.Label lblZonal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zonal"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1230
      TabIndex        =   1
      Top             =   705
      Width           =   2265
   End
End
Attribute VB_Name = "frmDemandInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSearchDemand_Click()
    ''----------------------------------------------'
    '               Validations                     '
    ''----------------------------------------------'
    If Trim(txtFromDate.Text) = "" Then
        MsgBox "From date is Mandatory"
        Exit Sub
    End If
    If Trim(txtToDate.Text) = "" Then
        MsgBox "To date is Mandatory"
        Exit Sub
    End If
    Call FillGrid(txtFromDate.Text, txtToDate.Text, IIf(val(lblZonal.Tag) = 0, "%", lblZonal.Tag), "%")
End Sub

Private Sub cmdSearchLocation_Click()
    lstZonals.Visible = True
    lstZonals.SetFocus
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
    If Trim(txtFromDate.Text) <> "" And Trim(txtToDate.Text) <> "" Then
        Call FillGrid(txtFromDate.Text, txtToDate.Text, IIf(val(lblZonal.Tag) = 0, "%", lblZonal.Tag), "%")
    End If
End Sub

Private Sub Form_Load()
    txtDate.Text = DdMmmYy(gbTransactionDate)
    txtFromDate.Text = DdMmmYy(DateAdd("m", -1, gbTransactionDate))
    txtToDate.Text = DdMmmYy(gbTransactionDate)
    PopulateList lstZonals, "Select chvZoneNameEnglish,numZoneID From GM_Zone Where intLBID = " & gbLocalBodyID & " Order By chvZoneNameEnglish", , True, , True, enuSourceString.DBMaster
End Sub

Private Sub FillGrid(ByVal dtFrom As Date, ByVal dtTo As Date, mLocationID As String, mTransactionTypeID As String)
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mRow As Integer
    Dim arrIn As Variant
    Dim RecFin As New ADODB.Recordset
    Dim mVoucherNo As Variant
    Dim mCnnFin As New ADODB.Connection
    
    arrIn = Array(dtFrom, dtTo, CStr(mLocationID), CStr(mTransactionTypeID))
    objdb.CreateNewConnection mCnn, enuSourceString.SaankhyaHO
    objdb.CreateNewConnection mCnnFin, enuSourceString.Saankhya
    Set Rec = objdb.ExecuteSP("spGetZonalDemand", arrIn, , , mCnn)
    vsGrid.Rows = 1
    vsGrid.Rows = 20
    mRow = 1
    If Not (Rec.EOF And Rec.BOF) Then
        While Not Rec.EOF
            With vsGrid
                .Rows = .Rows + 1
                .TextMatrix(mRow, 0) = Format(Rec!dtDemandDate, "Dd-MMM-Yyyy")
                .TextMatrix(mRow, 1) = GetZonalName(Rec!numLocationID)
              '  .TextMatrix(mRow, 2) = Rec!vchTransactionType
                .TextMatrix(mRow, 2) = Format(Rec!fltAmount, "0.00")
                If Rec!tnyStatus = 1 Then
                       vsGrid.Cell(flexcpBackColor, mRow, 0, , 4) = &HC0FFC0
                End If
                .TextMatrix(mRow, 3) = Rec!vchDemandNo
                .TextMatrix(mRow, 4) = Rec!numDemandID
                .TextMatrix(mRow, 5) = Rec!numLocationID
                .TextMatrix(mRow, 6) = Rec!intTransactionTypeID
            End With
            Rec.MoveNext
            vsGrid.Rows = vsGrid.Rows + 1
            mRow = mRow + 1
        Wend
    End If
    Rec.Close
End Sub

Private Sub lstZonals_DblClick()
        lblZonal.Caption = lstZonals.Text
        lblZonal.Tag = lstZonals.ItemData(lstZonals.ListIndex)
        lstZonals.Visible = False
End Sub

Private Sub lstZonals_LostFocus()
    lstZonals.Visible = False
End Sub

Private Sub txtFromDate_LostFocus()
    txtFromDate.Text = DdMmmYy(txtFromDate.Text)
End Sub

Private Sub txtToDate_LostFocus()
    txtToDate.Text = DdMmmYy(txtToDate.Text)
End Sub

Private Sub vsGrid_DblClick()
If vsGrid.Row > 0 And vsGrid.TextMatrix(vsGrid.Row, 3) <> "" And val(vsGrid.TextMatrix(vsGrid.Row, 5)) > 0 Then
     frmTransactionTypeWiseDemandInbox.ZonalID = val(vsGrid.TextMatrix(vsGrid.Row, 5))
     frmTransactionTypeWiseDemandInbox.Show
End If
End Sub

Private Function GetZonalName(mZonalID As Double) As String
    Dim mCount As Integer
    Dim strZonal As String
    For mCount = 0 To lstZonals.ListCount - 1
        If lstZonals.ItemData(mCount) = mZonalID Then
            strZonal = lstZonals.List(mCount)
        End If
    Next
    GetZonalName = strZonal
End Function

