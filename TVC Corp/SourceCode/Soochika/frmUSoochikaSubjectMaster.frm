VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmUSoochikaSubjectMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject Master"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   7050
   Begin VB.CommandButton btnClose 
      BackColor       =   &H80000000&
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
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7140
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   6945
      Begin VB.CheckBox chkListAll 
         BackColor       =   &H80000005&
         Caption         =   "List All"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5820
         TabIndex        =   5
         Top             =   285
         Width           =   945
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   615
         TabIndex        =   0
         Top             =   6615
         Width           =   6135
      End
      Begin VSFlex8LCtl.VSFlexGrid vsSubjectMaster 
         Height          =   5925
         Left            =   30
         TabIndex        =   1
         Top             =   615
         Width           =   6735
         _cx             =   11880
         _cy             =   10451
         Appearance      =   0
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
         FormatString    =   $"frmUSoochikaSubjectMaster.frx":0000
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
      Begin VB.Label lblheading 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   1605
         TabIndex        =   6
         Top             =   255
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   6735
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmUSoochikaSubjectMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub chkListAll_Click()
'CHNAGED by soumya vs
GetAllSubjectList
End Sub

Private Sub Form_Load()
    'CHANGED
    'Call FillGrid
    PreferedSubjectList
End Sub
Private Sub PreferedSubjectList()
Dim mCnn As New ADODB.Connection
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
End If
Set Rec = objDB.ExecuteSP("Sp_SelectSubjectPrefList", , , False, mCnn, adCmdStoredProc)
 If Not (Rec.EOF Or Rec.BOF) Then
 lblheading.Caption = "Prefered Subject List"
          vsSubjectMaster.Clear 1
            vsSubjectMaster.Rows = 1
            While Not Rec.EOF
                If mCount > vsSubjectMaster.Rows - 1 Then
                    vsSubjectMaster.Rows = vsSubjectMaster.Rows + 1
                End If
                vsSubjectMaster.TextMatrix(mCount, 0) = IIf(IsNull(Rec!numSubjectID), "", Rec!numSubjectID)
                vsSubjectMaster.TextMatrix(mCount, 1) = IIf(IsNull(Rec!chvSubject), "", Rec!chvSubject)
                Rec.MoveNext
                mCount = mCount + 1
            Wend
 Else
 lblheading.Caption = "All subject List"
 chkListAll.value = 1
 GetAllSubjectList
 End If

End Sub
Private Sub GetAllSubjectList()
Dim mCnn As New ADODB.Connection
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
End If
If (chkListAll.value = 1) Then
Set Rec = objDB.ExecuteSP("Sp_SelectSubjectList", , , False, mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblheading.Caption = "All subject List"
vsSubjectMaster.Clear 1
            vsSubjectMaster.Rows = 1
            While Not Rec.EOF
                If mCount > vsSubjectMaster.Rows - 1 Then
                    vsSubjectMaster.Rows = vsSubjectMaster.Rows + 1
                End If
                vsSubjectMaster.TextMatrix(mCount, 0) = IIf(IsNull(Rec!numSubjectID), "", Rec!numSubjectID)
                vsSubjectMaster.TextMatrix(mCount, 1) = IIf(IsNull(Rec!chvSubject), "", Rec!chvSubject)
                Rec.MoveNext
                mCount = mCount + 1
            Wend
End If
Else
lblheading.Caption = "Prefered Subject List"
PreferedSubjectList
End If

End Sub
Private Sub FillGrid()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mCount As Integer
        
        If objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False Then
            MsgBox "Cannot Continue.., Connection not present", vbInformation, "Soochika"
            Exit Sub
        End If
        mSQL = "SELECT mSubject.numSubjectID,chvSubject FROM mSubject"
        mSQL = mSQL + " where chvsubject like '%" & txtSubject.Text & "%'"
        mSQL = mSQL + " Order by mSubject.numSubjectID "
        Rec.Open mSQL, mCnn
        mCount = 1
        If Not (Rec.EOF And Rec.BOF) Then
            vsSubjectMaster.Clear 1
            vsSubjectMaster.Rows = 1
            While Not Rec.EOF
                If mCount > vsSubjectMaster.Rows - 1 Then
                    vsSubjectMaster.Rows = vsSubjectMaster.Rows + 1
                End If
                vsSubjectMaster.TextMatrix(mCount, 0) = IIf(IsNull(Rec!numSubjectID), "", Rec!numSubjectID)
                vsSubjectMaster.TextMatrix(mCount, 1) = IIf(IsNull(Rec!chvSubject), "", Rec!chvSubject)
                Rec.MoveNext
                mCount = mCount + 1
            Wend
        End If
    End Sub



Private Sub txtSubject_Change()
    FillGrid
End Sub

Private Sub vsSubjectMaster_DblClick()
    If InwardMode = 0 Then
        frmUSoochikaInward.txtSubID.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 0)
        frmUSoochikaInward.txtSubject.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 1)
    ElseIf InwardMode = 1 Then
        frmSoochikaManualInward.txtSubID.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 0)
        frmSoochikaManualInward.txtSubject.Text = vsSubjectMaster.TextMatrix(vsSubjectMaster.Row, 1)
    End If
    Unload Me
End Sub
