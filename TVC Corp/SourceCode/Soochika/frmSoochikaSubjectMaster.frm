VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSoochikaSubjectMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject Master"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6960
   Begin VB.CommandButton btnClose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6990
      Width           =   915
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   90
      Width           =   6855
      Begin VB.CheckBox chkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5790
         TabIndex        =   4
         Top             =   240
         Width           =   585
      End
      Begin VB.ComboBox cboCategory 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1230
         TabIndex        =   3
         Text            =   "cboCategory"
         Top             =   240
         Width           =   4425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   6855
      Begin VSFlex8LCtl.VSFlexGrid vsSubjectMaster 
         Height          =   5685
         Left            =   75
         TabIndex        =   6
         Top             =   255
         Width           =   6705
         _cx             =   11827
         _cy             =   10028
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSoochikaSubjectMaster.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
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
End
Attribute VB_Name = "frmSoochikaSubjectMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub cboCategory_Click()
Call FillGrid
End Sub

Private Sub chkAll_Click()
Call FillGrid
End Sub

Private Sub Form_Load()
    Call PopulateList(cboCategory, "Select chvLevel1, intLevel1 from TblLevel1", , , , True, enuSourceString.SOOCHIKA)
    Call FillGrid
End Sub
Private Sub FillGrid()
        Dim objdb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mCount As Integer
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False Then
            MsgBox "Cannot Continue.., Connection not present", vbInformation, "Soochika"
            Exit Sub
        End If
        mSQL = "SELECT TblSubjectCoding.intSubID,TblSubjectCoding.chvSubject,TblLevel1.chvLevel1"
        mSQL = mSQL + " FROM TblSubjectCoding INNER JOIN"
        mSQL = mSQL + " TblKWLevel ON TblSubjectCoding.intDistrID = TblKWLevel.intKWID INNER JOIN TblLevel1 ON TblKWLevel.intLevel1 = TblLevel1.intLevel1 "
        If (chkAll.Value <> 1) Then
            If (cboCategory.ListIndex <> -1) Then
                mSQL = mSQL + "Where TblLevel1.intLevel1=" & cboCategory.ItemData(cboCategory.ListIndex) & ""
            End If
        End If
        mSQL = mSQL + " Order by TblSubjectCoding.intSubID "
        Rec.Open mSQL, mCnn
        vsSubjectMaster.Rows = 1
        vsSubjectMaster.Rows = 7
        mCount = 1
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                If mCount > vsSubjectMaster.Rows - 1 Then
                    vsSubjectMaster.Rows = vsSubjectMaster.Rows + 1
                End If
                vsSubjectMaster.TextMatrix(mCount, 0) = IIf(IsNull(Rec!intSubID), "", Rec!intSubID)
                vsSubjectMaster.TextMatrix(mCount, 1) = IIf(IsNull(Rec!chvSubject), "", Rec!chvSubject)
                Rec.MoveNext
                mCount = mCount + 1
            Wend
        End If
    End Sub



Private Sub vsSubjectMaster_DblClick()
    If vsSubjectMaster.Row > 0 And vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 0) <> "" Then
         If gbSubID = 1 Then
            frmSoochikaInward.txtSubID.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 0)
            frmSoochikaInward.txtSubject.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 1)
            gbSubID = 0
            frmSoochikaInward.txtSubID.SetFocus
        ElseIf gbSubID = 2 Then
            frmSoochikaManualInward.txtSubID.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 0)
            frmSoochikaManualInward.txtSubject.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 1)
            gbSubID = 0
            frmSoochikaManualInward.txtSubID.SetFocus
        End If
            Unload Me
'        frmSoochikaInward.txtSubID.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 0)
'        frmSoochikaInward.txtSubject.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 1)
'        frmSoochikaInward.txtSubID.SetFocus
'        Unload Me
  End If
End Sub
