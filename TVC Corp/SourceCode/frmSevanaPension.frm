VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSevanaPension 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S e v a n a   P e n s  i o n"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   5100
      TabIndex        =   8
      Top             =   6015
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CanceL"
      Height          =   405
      Left            =   6150
      TabIndex        =   7
      Top             =   6015
      Width           =   1005
   End
   Begin VB.ListBox lstPensionerID 
      Height          =   1635
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   4035
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Width           =   1605
   End
   Begin VB.TextBox txtGrandTotal 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   9030
      TabIndex        =   2
      Top             =   6060
      Width           =   1605
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5055
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   11715
      _cx             =   20664
      _cy             =   8916
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
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
      Rows            =   10
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSevanaPension.frx":0000
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   225
      Left            =   8520
      TabIndex        =   4
      Top             =   300
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Money Order Return Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   480
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   225
      Left            =   8010
      TabIndex        =   1
      Top             =   6150
      Width           =   960
   End
End
Attribute VB_Name = "frmSevanaPension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim numPensionIDPart1 As String

Private Sub Form_Load()
    Call FillBill
    vsGrid.MergeCells = flexMergeFree
    vsGrid.MergeRow(0) = True
    Call GetPensinorID
    vsGrid.TextMatrix(1, 1) = numPensionIDPart1
    vsGrid.TextMatrix(1, 0) = 1
    
    vsGrid.ColComboList(3) = "|..."
       
End Sub

Private Function FillBill()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim intPensionTypeID As Integer
    Dim Rec As New ADODB.Recordset
    Dim aryIn As Variant
    Dim numPensionerID As Double
        objDB.CreateNewConnection mCnn, enuSourceString.SevanaPension
    gbLocalBodyID = 187
    intPensionTypeID = 1
    numPensionerID = 101870100003#
    aryIn = Array(gbLocalBodyID, intPensionTypeID, numPensionerID)
    Set Rec = objDB.ExecuteSP("Sp_TR_PensionBill_S5", aryIn, , , mCnn, adCmdStoredProc)
    
    While Not (Rec.EOF Or Rec.BOF)
            
''        lstPensionerID.AddItem IIf(IsNull(Rec!chvBillNo), "", Rec!chvBillNo)
''        lstPensionerID.ItemData(lstPensionerID.NewIndex) = IIf(IsNull(Rec!intAllotReqID), "", Rec!intAllotReqID)
''        Rec.MoveNext
        'vsGrid.ComboItem(Rec!intAllotReqID) = Rec!chvBillNo
        vsGrid.AddItem (Rec!chvBillNo)
        'vsGrid.ComboData (Rec!intAllotReqID)
        Rec.MoveNext
    Wend
    
    'objDB.FillGridCombo
    
    
End Function


Private Sub GetPensinorID()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New Recordset
    Dim mSQL As String
        objDB.CreateNewConnection mCnn, enuSourceString.SevanaPension
    mSQL = "Select tnyDBVolumeNo, intLBID From GM_LBSettings"
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        numPensionIDPart1 = CStr(Rec!tnyDBVolumeNo) + "0" + CStr(Rec!intLBID)
    End If
    If Rec.State = 1 Then Rec.Close
    If mCnn.State = 1 Then mCnn.Close
End Sub

Private Sub vsGrid_EnterCell()
    If vsGrid.Col = 2 Then
        vsGrid.TextMatrix(vsGrid.Row, 1) = numPensionIDPart1
        If vsGrid.Row <> 1 Then
            vsGrid.TextMatrix(vsGrid.Row, 0) = Val(vsGrid.TextMatrix(vsGrid.Row - 1, 0) + 1)
        End If
    End If
End Sub

Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        vsGrid.RemoveItem (vsGrid.Row)
    End If
End Sub

Private Sub vsGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vsGrid.Col = 6 Then
            vsGrid.Row = vsGrid.Row + 1
            vsGrid.TextMatrix(vsGrid.Row, 1) = numPensionIDPart1
            vsGrid.TextMatrix(vsGrid.Row, 0) = Val(vsGrid.TextMatrix(vsGrid.Row - 1, 0) + 1)
            vsGrid.Col = 2
            vsGrid.Rows = vsGrid.Rows + 1
        Else
            vsGrid.Col = vsGrid.Col + 1
        End If
    End If
End Sub

Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If Col = 2 Then
            Dim mIndex As Long
            Dim mStr
            With lstPensionerID
                mIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal vsGrid.TextMatrix(vsGrid.Row, 2))
                If mIndex >= 0 Then
                    .ListIndex = mIndex
                    
                    vsGrid.TextMatrix(vsGrid.Row, 2) = lstPensionerID.List(mIndex)
                End If
            End With
            
            'If mIndex >= 0 Then
            '    mStr = vsGrid.TextMatrix(vsGrid.Row, 2)
            '    vsGrid.TextMatrix(vsGrid.Row, 2) = lstPensionerID.List(mIndex)
            'End If
            
        End If
End Sub

Private Function GetHeadWiseTotal(ByVal intAccountHeadID As Long) As Long
    
End Function
