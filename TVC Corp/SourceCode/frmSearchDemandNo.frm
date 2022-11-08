VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchDemandNo 
   BackColor       =   &H00D3F7EA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Demand Numbers"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstMasters 
      Height          =   255
      Left            =   4620
      TabIndex        =   8
      Top             =   165
      Width           =   2580
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3F7EA&
      Height          =   3300
      Left            =   -15
      TabIndex        =   6
      Top             =   780
      Width           =   8205
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   3150
         Left            =   15
         TabIndex        =   7
         Top             =   120
         Width           =   8175
         _cx             =   14420
         _cy             =   5556
         Appearance      =   2
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
         BackColorFixed  =   13891562
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSearchDemandNo.frx":0000
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
   Begin VB.CommandButton cmdTransactionType 
      Caption         =   "..."
      Height          =   285
      Left            =   4155
      TabIndex        =   5
      Top             =   480
      Width           =   300
   End
   Begin VB.TextBox txtTransactionType 
      Height          =   285
      Left            =   1410
      TabIndex        =   4
      Top             =   495
      Width           =   2745
   End
   Begin VB.CommandButton cmdSection 
      Caption         =   "..."
      Height          =   285
      Left            =   4170
      TabIndex        =   2
      Top             =   150
      Width           =   300
   End
   Begin VB.TextBox txtSection 
      Height          =   285
      Left            =   1425
      TabIndex        =   1
      Top             =   165
      Width           =   2745
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   105
      Shape           =   3  'Circle
      Top             =   4275
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   510
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sections:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   735
      TabIndex        =   0
      Top             =   180
      Width           =   690
   End
End
Attribute VB_Name = "frmSearchDemandNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mType As Variant
Private Sub FillGrid()
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mCount As Long
    Dim mSql As String
    vsGrid.Clear 1, 0
    If val(txtTransactionType.Tag) > 0 Then
        mSql = mSql + " Select A.numDemandID, vchDemandNo, vchName + IsNull(' .'+vchInit1,'') + IsNull(' .'+vchInit2,'') As Name,"
        mSql = mSql + " A.intWardNo, A.intDoorNo , A.vchDoorNo2, chvSeatTitle,"
        mSql = mSql + " (Select Sum(C.fltAmount) From faIDemandChild C Where C.numDemandID = A.numDemandID) As fltAmount"
        mSql = mSql + " From faIDemandTbl A Inner Join faIDemandAddress On faIDemandAddress.numDemandID = A.numDemandID"
        mSql = mSql + " Left join DB_Masters.dbo.GL_Seats As GL_Seats ON GL_Seats.numSeatID = A.numSeatID AND GL_Seats.intLocalBodyID = A.intLBID"
        mSql = mSql + " Where A.intTransactionTypeID = " & val(txtTransactionType.Tag)
        mSql = mSql + " AND A.tnyStatus = 0 "
        mSql = mSql + " AND isNull(A.tnyExtModuleID,99) <> 55 "
        mSql = mSql + " Order By A.numDemandID Desc"
        objdb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        If Not (Rec.BOF And Rec.EOF) Then
            On Error Resume Next
            While Not Rec.EOF
                mCount = mCount + 1
                If vsGrid.Rows - 1 = mCount Then
                    vsGrid.Rows = vsGrid.Rows + 25
                End If
                vsGrid.TextMatrix(mCount, 0) = Rec!numDemandID
                vsGrid.TextMatrix(mCount, 1) = Rec!vchDemandNo
                vsGrid.TextMatrix(mCount, 2) = Rec!Name
                vsGrid.TextMatrix(mCount, 3) = Rec!intWardNo
                vsGrid.TextMatrix(mCount, 4) = Rec!intDoorNo & IIf(IsNull(Rec!vchDoorNo2), "", "\" & Rec!vchDoorNo2)
                vsGrid.TextMatrix(mCount, 5) = Rec!fltAmount
                vsGrid.TextMatrix(mCount, 6) = Rec!chvSeatTitle
                Rec.MoveNext
            Wend
            On Error GoTo 0
        End If
        Rec.Close
    End If
End Sub
Private Sub ListMaster()
    Dim mSql As String
    If mType = 1 Then
        mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID = 10 Order By vchTransactionType"
    Else
        mSql = "Select vchSectionName, intSectionID From faSection WHERE intSectionID  > 99 Order by vchSectionName"
    End If
    PopulateList lstMasters, mSql, , , True, True
    lstMasters.Height = 3500
    lstMasters.Width = 3500
    lstMasters.Visible = True
    lstMasters.SetFocus
End Sub
Private Sub cmdSection_Click()
    mType = 2
    Call ListMaster
End Sub
Private Sub cmdTransactionType_Click()
    mType = 1
    Call ListMaster
End Sub
Private Sub FormInitialize()
    txtTransactionType.Text = ""
    txtTransactionType.Tag = ""
    txtSection.Text = ""
    txtSection.Tag = ""
    lstMasters.Visible = True
    vsGrid.Clear 0, 1
    mType = Null
End Sub
Private Sub Form_Activate()
    If val(txtTransactionType.Tag) > 0 Then
        Dim objTrn As New clsTransactionType
        objTrn.SetTransactionType val(txtTransactionType.Tag)
        If objTrn.TransactionTypeID > 0 Then
            txtTransactionType.Text = objTrn.TransactionType
            txtTransactionType.Tag = objTrn.TransactionTypeID
            Call FillGrid
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub lstMasters_DblClick()
    Call lstMasters_LostFocus
End Sub
Private Sub lstMasters_LostFocus()
    lstMasters.Visible = False
    If mType = 1 Then
        If lstMasters.ListIndex > -1 Then
            txtTransactionType.Text = lstMasters.Text
            txtTransactionType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
            txtTransactionType.SetFocus
            Call FillGrid
        End If
    Else
        If lstMasters.ListIndex > -1 Then
            txtSection.Text = lstMasters.Text
            txtSection.Tag = lstMasters.ItemData(lstMasters.ListIndex)
            txtSection.SetFocus
        End If
    End If
End Sub
Private Sub vsGrid_DblClick()
    If val(vsGrid.TextMatrix(vsGrid.Row, 0)) > 0 Then
        gbSearchID = val(vsGrid.TextMatrix(vsGrid.MouseRow, 0))
        gbSearchStr = vsGrid.TextMatrix(vsGrid.MouseRow, 1)
        Unload Me
    End If
End Sub
