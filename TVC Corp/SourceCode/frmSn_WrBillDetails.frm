VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSn_WrBillDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Details"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSn_WrBillDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4245
      TabIndex        =   27
      Top             =   3345
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3090
      TabIndex        =   26
      Top             =   3345
      Width           =   1140
   End
   Begin VSFlex8LCtl.VSFlexGrid fgBillDetails 
      Height          =   1260
      Left            =   1650
      TabIndex        =   22
      Top             =   1515
      Width           =   5685
      _cx             =   10028
      _cy             =   2222
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSn_WrBillDetails.frx":1CCA
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
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5745
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   2820
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   15
      TabIndex        =   25
      Top             =   -75
      Width           =   8520
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   6780
         TabIndex        =   28
         Top             =   870
         Width           =   1695
      End
      Begin VB.TextBox txtReading2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4470
         TabIndex        =   21
         Top             =   1200
         Width           =   1500
      End
      Begin VB.TextBox txtReading1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   4470
         TabIndex        =   19
         Top             =   885
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearchOffice 
         Caption         =   "..."
         Height          =   285
         Left            =   2910
         TabIndex        =   7
         Top             =   540
         Width           =   315
      End
      Begin VB.CommandButton cmdSearchCaretaker 
         Caption         =   "..."
         Height          =   300
         Left            =   2910
         TabIndex        =   2
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtOfficeInst 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         TabIndex        =   6
         Top             =   525
         Width           =   1380
      End
      Begin VB.TextBox txtCaretaker 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1530
         TabIndex        =   1
         Top             =   210
         Width           =   1380
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1530
         TabIndex        =   9
         Top             =   1155
         Width           =   1380
      End
      Begin VB.TextBox txtConsumerNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtBillDate 
         Height          =   300
         Left            =   4470
         TabIndex        =   11
         Top             =   210
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   15728641
         CurrentDate     =   40106
      End
      Begin MSComCtl2.DTPicker dtBillDueDate 
         Height          =   300
         Left            =   4470
         TabIndex        =   13
         Top             =   525
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   15728641
         CurrentDate     =   40106
      End
      Begin MSComCtl2.DTPicker dtpBillFrom 
         Height          =   300
         Left            =   6795
         TabIndex        =   15
         Top             =   210
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   15728641
         CurrentDate     =   40106
      End
      Begin MSComCtl2.DTPicker dtpBillTo 
         Height          =   300
         Left            =   6795
         TabIndex        =   17
         Top             =   525
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   15728641
         CurrentDate     =   40106
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6105
         TabIndex        =   29
         Top             =   870
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crnt. Reading"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3330
         TabIndex        =   20
         Top             =   1230
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prvs. Reading"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   18
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumer No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   405
         TabIndex        =   3
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No"
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3750
         TabIndex        =   10
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3675
         TabIndex        =   12
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Care Taker"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   600
         TabIndex        =   0
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office/Institution"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   45
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6345
         TabIndex        =   14
         Top             =   225
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6540
         TabIndex        =   16
         Top             =   555
         Width           =   210
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5265
      TabIndex        =   23
      Top             =   2850
      Width           =   420
   End
End
Attribute VB_Name = "frmSn_WrBillDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mCnn As New ADODB.Connection
    Public intBillId As Integer
    '*********************************************************************************************'
    '                                   Form to generate the Water Bill                           '
    '*********************************************************************************************'
    Private Sub FillvsGrid(BillID As Integer)
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim mRowCount   As Integer
                
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        
        fgBillDetails.Clear 1, 1
        fgBillDetails.Rows = 1
        mRowCount = 1
        
        mSql = "Select * From snWrBillDetailsChild"
        mSql = mSql + " Left Join snWrBillDetailsHead On snWrBillDetailsChild.intBillHeadID = snWrBillDetailsHead.intID"
        mSql = mSql + " Where intBillID = " & BillID
        Rec.Open mSql, mCnn
        While Not Rec.EOF
            fgBillDetails.Rows = fgBillDetails.Rows + 1
            fgBillDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intBillHeadID), "", Rec!intBillHeadID)
            fgBillDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!chvSaankhyaCode), "", Rec!chvSaankhyaCode)
            fgBillDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!chvName), "", Rec!chvName)
            fgBillDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
        Rec.Close
        Call CalcTotalAmount
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub
    
    Private Sub cmdSave_Click()
        Dim vParamIn    As Variant
        Dim vPAryOut    As Variant, i As Integer
        Dim objDb       As New clsDB
        ReDim vParamIn(14)
        
        '*********************************************************************************************'
        '                                   Procedure to generate the Water Bill                      '
        '*********************************************************************************************'
        On Error GoTo err
        If txtCaretaker.Tag = "" Then
            MsgBox "Please select the Caretaker", vbInformation
            cmdSearchCaretaker.SetFocus
            Exit Sub
        End If
        If txtOfficeInst.Tag = "" Then
            MsgBox "Please select the Office/Institution", vbInformation
            cmdSearchOffice.SetFocus
            Exit Sub
        End If
        If txtConsumerNo.Text = "" Then
            MsgBox "Please enter the Consumer No", vbInformation
            txtConsumerNo.SetFocus
            Exit Sub
        End If
        If txtBillNo.Text = "" Then
            MsgBox "Please enter the BillNo", vbInformation
            txtBillNo.SetFocus
            Exit Sub
        End If
        If txtReading2.Text = "" Then
            MsgBox "Please enter the Reading", vbInformation
            txtReading2.SetFocus
            Exit Sub
        End If
        vParamIn(0) = txtConsumerNo.Tag 'Connection ID
        vParamIn(1) = Trim(txtConsumerNo)
        If Trim(txtBillNo.Text) <> "" Then
            vParamIn(2) = Trim(txtBillNo)
        Else
            MsgBox "Please enter the Bill No", vbInformation
            txtBillNo.SetFocus
            Exit Sub
        End If
        If dtBillDate.Enabled = True Then
            vParamIn(3) = Format(dtBillDate.value, "DD/MM/YYYY")
        Else
            vParamIn(3) = Null
        End If
        If dtBillDueDate.Enabled = True Then
            vParamIn(4) = Format(dtBillDueDate.value, "DD/MM/YYYY")
        Else
            vParamIn(4) = Null
        End If
        vParamIn(5) = val(txtTotal)
        vParamIn(6) = Trim(txtRemarks)
        vParamIn(7) = Format(dtpBillFrom.value, "DD/MM/YYYY")
        vParamIn(8) = Format(dtpBillTo.value, "DD/MM/YYYY")
        vParamIn(9) = val(Trim(txtReading1))
        vParamIn(10) = val(Trim(txtReading2))
        vParamIn(11) = gbUserID
        vParamIn(12) = gbSeatID
        vParamIn(13) = 0
        If intBillId = 0 Then
            vParamIn(14) = 0 'intBillId
        Else
            vParamIn(14) = intBillId 'intBillId
        End If
        
'        ExecuteSP "snWrBillDetails_I", rinsert, adCmdStoredProc, vParamIn, vPAryOut, conSanchaya
        objDb.ExecuteSP "snWrBillDetails_I", vParamIn, vPAryOut, , mCnn, adCmdStoredProc
        If IsArray(vPAryOut) Then
            intBillId = vPAryOut(0, 0)
        End If
        'ExecuteSP "delete from snWrBillDetailsChild where intBillId=" & intBillId, RDelete, adCmdText, , , conSanchaya
        mCnn.Execute "delete from snWrBillDetailsChild where intBillId=" & intBillId
        ReDim vParamIn(2)
        For i = 1 To fgBillDetails.Rows - 1
            vParamIn(0) = intBillId 'intBillId
            vParamIn(1) = val(fgBillDetails.Cell(flexcpText, i, 0)) 'intBillHeadId
'            If fgBillDetails.Cell(flexcpText, i, 3) = "" Then
                vParamIn(2) = val(fgBillDetails.Cell(flexcpText, i, 3)) 'fltAmount
 '           End If
            'ExecuteSP "snWrBillDetailsChild_I", rinsert, adCmdStoredProc, vParamIn, , conSanchaya
            objDb.ExecuteSP "snWrBillDetailsChild_I", vParamIn, , , mCnn, adCmdStoredProc
        Next i
        MsgBox "Saved Successfully", vbInformation, "Sanchaya"
        cmdSave.Enabled = False
        'frmSn_WrBillList.PopulateBills
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdSearchOffice_Click()
        On Error GoTo err
        If txtCaretaker.Tag <> "" Then
            intWrBillSearchID = 7
            intWrBillCaretakerID = txtCaretaker.Tag
            frmSn_WrBillSearchName.Show 1
        End If
        txtConsumerNo.SetFocus
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdSearchCaretaker_Click()
        intWrBillSearchID = 6
        txtCaretaker.Text = ""
        txtCaretaker.Tag = ""
        txtOfficeInst.Text = ""
        txtOfficeInst.Tag = ""
        frmSn_WrBillSearchName.Show 1
    End Sub

    Private Sub fgBillDetails_Click()
        If fgBillDetails.col = 3 Then
            fgBillDetails.Editable = flexEDKbdMouse
        Else
            fgBillDetails.Editable = flexEDNone
        End If
    End Sub
    
    Private Sub fgBillDetails_KeyPressEdit(ByVal row As Long, ByVal col As Long, KeyAscii As Integer)
        If col = 3 Then
            If Chr(KeyAscii) = "." And (InStr(1, Txt, ".", vbTextCompare) > 0) Then
                 KeyAscii = 0
                 Exit Sub
             Else
                 If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 46)) Then
                     KeyAscii = 0
                 End If
            End If
        End If
    End Sub
    
    Private Sub fgBillDetails_LostFocus()
        CalcTotalAmount
    End Sub
    
    Private Sub Form_Activate()
        Me.Height = 4245
        Me.Width = 8625
    End Sub

    Private Sub Form_Load()
        Dim varyOut     As Variant
        Dim i           As Integer
        Dim objDb       As New clsDB
        Dim mSql        As String
        
        On Error GoTo err
        'CenterForm Me
        dtBillDate.value = Date
        dtBillDueDate.value = Date
        fgBillDetails.Clear 1
        fgBillDetails.Rows = 2
        If frmSn_WrBillListOfTransactionDetails.txtCaretaker.Tag <> "" Then
            txtCaretaker.Text = frmSn_WrBillListOfTransactionDetails.txtCaretaker.Text
            txtCaretaker.Tag = frmSn_WrBillListOfTransactionDetails.txtCaretaker.Tag
        End If
        'Set conSanchaya = gFunSetConnection(Dsn.Sanchaya)
        objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        
        'ExecuteSP "spsnWrBillDetailsHead_S", rselect, adCmdStoredProc, , vAryOut, conSanchaya
        objDb.ExecuteSP "spsnWrBillDetailsHead_S", , varyOut, , mCnn, adCmdStoredProc
        If IsArray(varyOut) Then
            For i = 0 To UBound(varyOut, 2)
                fgBillDetails.Cell(flexcpText, i + 1, 0) = varyOut(0, i) 'BillDetailsHeadID
                fgBillDetails.Cell(flexcpText, i + 1, 2) = varyOut(1, i) 'BillDetailsHead
                fgBillDetails.Cell(flexcpText, i + 1, 1) = varyOut(2, i) 'chvSaankhyaCode
                fgBillDetails.Rows = fgBillDetails.Rows + 1
            Next i
            fgBillDetails.Rows = fgBillDetails.Rows - 1
        End If
        If frmSn_WrBillDetails.intBillId <> 0 Then
            Call FillvsGrid(frmSn_WrBillDetails.intBillId)
        End If
        Exit Sub
err:
        MsgBox err.Description
        End Sub
    
    Public Function CalcTotalAmount()
    Dim TotalAmount As Double, fgRowCnt As Integer
        
        On Error GoTo err
        dblTotAmount = 0
        For fgRowCnt = 0 To fgBillDetails.Rows - 1
            If fgBillDetails.Cell(flexcpText, fgRowCnt, 3) <> "" Then
                dblTotAmount = dblTotAmount + val(fgBillDetails.Cell(flexcpText, fgRowCnt, 3))
            End If
        Next fgRowCnt
        txtTotal = dblTotAmount
        Exit Function
err:
        MsgBox err.Description
    End Function
 
    Private Sub txtConsumerNo_GotFocus()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        On Error GoTo err
        If txtOfficeInst.Tag <> "" Then
            objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
            
            mSql = "Select chvConsumerNo,intID From snWrBillConnections"
            mSql = mSql + " Where intOfficeID = " & txtOfficeInst.Tag
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtConsumerNo.Text = IIf(IsNull(Rec!chvConsumerNo), "", Rec!chvConsumerNo)
                txtConsumerNo.Tag = IIf(IsNull(Rec!intID), "", Rec!intID)
            End If
            Rec.Close
            If txtConsumerNo.Text <> "" Then
                mSql = "Select numCurrentReading From snWrBillDetails"
                mSql = mSql + " Where chvConsumerNo ='" & txtConsumerNo.Text & "'"
                mSql = mSql + " And intBillID = (Select MAX(intBillID) From snWrBillDetails Where chvConsumerNo ='" & txtConsumerNo.Text & "')"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    txtReading1.Text = IIf(IsNull(Rec!numCurrentReading), 0, Rec!numCurrentReading)
                End If
                Rec.Close
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub txtReading1_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtReading2_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
