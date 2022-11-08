VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchMODetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search MO Details"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "frmSearchMODetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbBillNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1110
      Width           =   1680
   End
   Begin VB.ComboBox cmbPensionType 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   435
      Width           =   3570
   End
   Begin VB.TextBox txtPensionerID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2910
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   795
      Width           =   2205
   End
   Begin VB.TextBox txtPrefix 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   795
      Width           =   1320
   End
   Begin VB.CommandButton cmdSearchBill 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5550
      TabIndex        =   4
      Top             =   960
      Width           =   1050
   End
   Begin VB.TextBox txtRowCount 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   1110
      Width           =   915
   End
   Begin VSFlex8LCtl.VSFlexGrid vsBill 
      Height          =   2325
      Left            =   75
      TabIndex        =   5
      Top             =   1650
      Width           =   6750
      _cx             =   11906
      _cy             =   4101
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchMODetails.frx":1CCA
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
   Begin VB.Label lblPensionType 
      AutoSize        =   -1  'True
      Caption         =   "Pensioner Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   10
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label lblPensionerID 
      AutoSize        =   -1  'True
      Caption         =   "Pensioner ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   405
      TabIndex        =   9
      Top             =   795
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "   Search Bills"
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
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label lblBillID 
      AutoSize        =   -1  'True
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   7
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lblRowCount 
      AutoSize        =   -1  'True
      Caption         =   "RowCount"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   6
      Top             =   1140
      Width           =   870
   End
End
Attribute VB_Name = "frmSearchMODetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim gPensionerID As Variant
    
    Private Sub FillvsBill()
        Dim mcnn            As New ADODB.Connection
        Dim objDB           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSQL            As String
        Dim mArrIn          As Variant
        Dim mPensionerID    As String
        Dim mRowCount       As Double
        Dim mBillID         As Integer
        
        On Error GoTo err
        If (objDB.CreateNewConnection(mcnn, enuSourceString.SevanaPension)) Then
            mPensionerID = gPensionerID
            If cmbBillNo.ListIndex > 0 Then
                If cmbBillNo.ItemData(cmbBillNo.ListIndex) > 0 Then
                    mBillID = cmbBillNo.ItemData(cmbBillNo.ListIndex)
                End If
            End If
            gbLocalBodyID = 1250
            mArrIn = Array(gbLocalBodyID, _
                            mPensionerID, _
                            mBillID, _
                            IIf(Trim(txtRowCount.Text) = "", 0, txtRowCount.Text) _
                            )
            Rec.CursorLocation = adUseClient
            Set Rec = objDB.ExecuteSP("KMAM_BillDetails", mArrIn, , , mcnn)
                
            vsBill.Rows = 1
            vsBill.Clear 1, 1
            mRowCount = 1
            If Rec.State = 1 Then
                While Not Rec.EOF
                    vsBill.Rows = vsBill.Rows + 1
                    vsBill.TextMatrix(mRowCount, 0) = mRowCount
                    vsBill.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!chvPensionerName), "", Rec!chvPensionerName)
                    vsBill.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!chvBillNo), "", Rec!chvBillNo)
                    vsBill.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsBill.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!intAllotReqID), "", Rec!intAllotReqID)
                    vsBill.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!tnyPensionTypeID), "", Rec!tnyPensionTypeID)
                    mRowCount = mRowCount + 1
                    Rec.MoveNext
                Wend
            End If
        Else
            MsgBox "Connection To Pension Database does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdSearchBill_Click()
        FillvsBill
    End Sub

    Private Sub Form_Load()
        Dim mSQL            As String
        
        mSQL = "Select chvPensionNameEnglish,tnyPensionTypeID From GM_PensionType"
        PopulateList cmbPensionType, mSQL, , True, True, True, SevanaPension
                
        If gPensionerID <> "" Then
            mSQL = "Select Isnull(TR_PensionBill.chvBillNo,''),TR_PensionerBill.intAllotReqID As intAllotReqID"
            mSQL = mSQL + " From TR_PensionerBill"
            mSQL = mSQL + " Inner Join TR_Pension On TR_Pension.numPensionerID = TR_PensionerBill.numPensionerID"
            mSQL = mSQL + " Inner Join TR_PensionBill On TR_PensionerBill.intAllotReqID = TR_PensionBill.intAllotReqID"
            mSQL = mSQL + " And TR_PensionerBill.tnypensiontypeID =TR_PensionBill.tnypensiontypeID"
            mSQL = mSQL + " Where TR_PensionerBill.numPensionerId = " & gPensionerID
            PopulateList cmbBillNo, mSQL, , True, True, True, enuSourceString.SevanaPension
        End If
        
        FillvsBill
    End Sub
    
    Private Sub txtBillID_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtPensionerID_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtPrefix_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtRowCount_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsBill_DblClick()
        Dim mRowCount As Double
        
        If vsBill.Row > 0 Then
            If vsBill.TextMatrix(vsBill.Row, 4) <> "" And val(vsBill.TextMatrix(vsBill.Row, 3)) > 0 Then
                'frmMOReturned.cmbPensionType.ListIndex = cmbPensionType.ListIndex
                'frmMOReturned.txtPrefix.Text = txtPrefix.Text
                'frmMOReturned.txtPensionerID.Text = txtPensionerID.Text
                For mRowCount = 0 To frmMOReturned.vsBill.Rows - 1
                    If frmMOReturned.vsBill.TextMatrix(mRowCount, 6) = gPensionerID Then
                        MsgBox "Another bill is already added for this Pensioner", vbInformation
                        Exit Sub
                    End If
                Next
                frmMOReturned.txtPrefix.Tag = vsBill.TextMatrix(vsBill.Row, 4)
                frmMOReturned.txtPensionerID.Tag = vsBill.TextMatrix(vsBill.Row, 2)
                frmMOReturned.vsBill.Rows = frmMOReturned.vsBill.Rows + 1
                mRowCount = frmMOReturned.vsBill.Rows
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 0) = mRowCount - 1
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 1) = vsBill.TextMatrix(vsBill.Row, 1)
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 2) = vsBill.TextMatrix(vsBill.Row, 2)
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 3) = vsBill.TextMatrix(vsBill.Row, 3)
                frmMOReturned.vsBill.Cell(flexcpChecked, mRowCount - 1, 4) = 1
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 5) = vsBill.TextMatrix(vsBill.Row, 4)
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 6) = gPensionerID
                frmMOReturned.vsBill.TextMatrix(mRowCount - 1, 7) = vsBill.TextMatrix(vsBill.Row, 5)
                Unload Me
            End If
        End If
    End Sub
    
    Public Property Let PensionerID(mdata As Variant)
        gPensionerID = mdata
    End Property
