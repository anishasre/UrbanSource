VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmTest3 
   Caption         =   "."
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPTAX 
      Caption         =   "PTAX CALCULATOR"
      Height          =   510
      Left            =   8280
      TabIndex        =   40
      Top             =   495
      Width           =   1500
   End
   Begin VB.ComboBox cmbWard 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2220
      Width           =   2580
   End
   Begin VB.TextBox txtAmt 
      Height          =   285
      Left            =   1860
      TabIndex        =   38
      Top             =   1140
      Width           =   1455
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   3630
      TabIndex        =   36
      Top             =   420
      Width           =   1455
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1860
      TabIndex        =   34
      Top             =   420
      Width           =   1455
   End
   Begin VB.CommandButton ccmdTransType 
      Caption         =   "..."
      Height          =   255
      Left            =   4860
      TabIndex        =   33
      Top             =   1890
      Width           =   255
   End
   Begin VB.CommandButton cmdAcHead 
      Caption         =   "..."
      Height          =   255
      Left            =   4860
      TabIndex        =   32
      Top             =   3300
      Width           =   255
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   330
      Left            =   1860
      TabIndex        =   31
      Top             =   3660
      Width           =   1020
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1875
      Left            =   270
      TabIndex        =   30
      Top             =   4110
      Width           =   9405
      _cx             =   16589
      _cy             =   3307
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTest3.frx":0000
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
   Begin VB.CommandButton cmdBankCharge 
      Caption         =   "..."
      Height          =   255
      Left            =   9300
      TabIndex        =   29
      Top             =   3180
      Width           =   255
   End
   Begin VB.CommandButton cmdRetured 
      Caption         =   "..."
      Height          =   255
      Left            =   9300
      TabIndex        =   28
      Top             =   2820
      Width           =   255
   End
   Begin VB.CommandButton cmdRemittance 
      Caption         =   "..."
      Height          =   255
      Left            =   9300
      TabIndex        =   27
      Top             =   2460
      Width           =   255
   End
   Begin VB.TextBox txttBankCharge 
      Height          =   285
      Left            =   7800
      TabIndex        =   25
      Top             =   3180
      Width           =   1455
   End
   Begin VB.TextBox dttBankCharge 
      Height          =   285
      Left            =   6300
      TabIndex        =   24
      Top             =   3180
      Width           =   1455
   End
   Begin VB.TextBox txtReturned 
      Height          =   285
      Left            =   7800
      TabIndex        =   22
      Top             =   2820
      Width           =   1455
   End
   Begin VB.TextBox dttReturned 
      Height          =   285
      Left            =   6300
      TabIndex        =   21
      Top             =   2820
      Width           =   1455
   End
   Begin VB.TextBox txtRemittance 
      Height          =   285
      Left            =   7800
      TabIndex        =   19
      Top             =   2460
      Width           =   1455
   End
   Begin VB.TextBox dtRemittance 
      Height          =   285
      Left            =   6300
      TabIndex        =   18
      Top             =   2460
      Width           =   1455
   End
   Begin VB.CheckBox chkCheque 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque Enrolled in Bank A/c "
      Height          =   195
      Left            =   6330
      TabIndex        =   16
      Top             =   1770
      Width           =   2415
   End
   Begin VB.TextBox txtBankAccounthead 
      Height          =   285
      Left            =   1860
      TabIndex        =   15
      Top             =   3300
      Width           =   2955
   End
   Begin VB.TextBox txtRefNo 
      Height          =   285
      Left            =   1860
      TabIndex        =   13
      Top             =   2940
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2820
      TabIndex        =   11
      Top             =   2580
      Width           =   495
   End
   Begin VB.TextBox txtDoorNo 
      Height          =   285
      Left            =   1860
      TabIndex        =   10
      Top             =   2580
      Width           =   915
   End
   Begin VB.TextBox txtWardNo 
      Height          =   285
      Left            =   1860
      TabIndex        =   8
      Top             =   2220
      Width           =   435
   End
   Begin VB.TextBox txtTransactiontype 
      Height          =   285
      Left            =   1860
      TabIndex        =   6
      Top             =   1860
      Width           =   2955
   End
   Begin VB.TextBox txtPartyName 
      Height          =   285
      Left            =   1860
      TabIndex        =   4
      Top             =   1500
      Width           =   2955
   End
   Begin VB.TextBox txtInstrumentNo 
      Height          =   285
      Left            =   1860
      TabIndex        =   1
      Top             =   780
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   3360
      TabIndex        =   37
      Top             =   450
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Date Period"
      Height          =   195
      Left            =   930
      TabIndex        =   35
      Top             =   420
      Width           =   840
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Bank Charge"
      Height          =   195
      Left            =   5370
      TabIndex        =   26
      Top             =   3240
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Returned"
      Height          =   195
      Left            =   5550
      TabIndex        =   23
      Top             =   2880
      Width           =   660
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Remittance"
      Height          =   195
      Left            =   5400
      TabIndex        =   20
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   " Date                 Scroll Entry        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6330
      TabIndex        =   17
      Top             =   2160
      Width           =   2955
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Bank Account Head"
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Ref. No."
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Door No"
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ward"
      Height          =   195
      Left            =   1380
      TabIndex        =   7
      Top             =   2280
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Type"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name Of Party"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label txtAmount 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Left            =   1260
      TabIndex        =   2
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Insturment No"
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "frmTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''Private Sub cmd_Save_Click()
''''''''Dim mCnn                  As New ADODB.Connection
''''''''Dim objDB                 As New clsDb
''''''''Dim Rec                   As New Recordset
''''''''Dim mSql                  As Variant
''''''''Dim mArrIn                As Variant
''''''''Dim mArrOut               As Variant
''''''''
''''''''objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
'''''''''*************************** validations*****************
''''''''If txtAllotmentNo = "" Then
''''''''  MsgBox "Please Enter the Allotment No", vbInformation
''''''''  txtAllotmentNo.SetFocus
''''''''  Exit Sub
''''''''End If
''''''''If txtAllotDate = "" Then
''''''''  MsgBox "please Select the Allotment Date", vbInformation
''''''''  txtAllotDate.SetFocus
''''''''  Exit Sub
''''''''End If
''''''''If cmbCategory.ListIndex = -1 Then
''''''''  MsgBox "Plase select the Category", vbInformation
''''''''  cmbCategory.SetFocus
''''''''  Exit Sub
''''''''End If
''''''''
''''''''If cmbCategory.ItemData(cmbCategory.ListIndex) = 0 Then
''''''''  MsgBox "Please Select the Category", vbInformation
''''''''  cmbCategory.SetFocus
''''''''  Exit Sub
'''''''' End If
''''''''
'''''''''objDB.ExecuteSP "spAllotmentLetterSave", mArrIn, mArrOut, mCnn, adCmdStoredProc
''''''''
''''''''
''''''''End Sub
''''''''
''''''''Private Sub Command1_Click()
''''''''
''''''''    Debug.Print Chr(64)
''''''''
''''''''End Sub
''''''''
''''''''Private Sub txtTreasuryName_KeyDown(KeyCode As Integer, Shift As Integer)
''''''''    txtAccHead.Text = Chr$(KeyCode)
''''''''End Sub
''''''''
''''''''Private Sub txtTreasuryName_KeyPress(KeyAscii As Integer)
''''''''
''''''''End Sub
''''''''
''''''''Private Sub txtTreasuryName_LostFocus()
''''''''    txtAccHead.Text = Chr$(Val(txtTreasuryName))
''''''''End Sub
''''''''
''''''''Private Sub Text1_Change()
''''''''
''''''''End Sub
''''''''
''''''''Private Sub Text2_Change()
''''''''
''''''''End Sub
Option Explicit

    Private Sub cmdAcHead_Click()
       On Error GoTo err:
            Dim mSql As String
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & 2
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.Show vbModal
            txtBankAccounthead.Text = gbSearchStr
            txtBankAccounthead.Tag = gbSearchID
            txtBankAccounthead.SetFocus
            gbSearchID = -1
            gbSearchStr = ""
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdFind_Click()
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim mSql        As String
        If txtBankAccounthead.Text = "" Then
            
        End If
        
        mSql = "Select intVoucherID,intVoucherNo,dtDate,vchInstrumentNo,dtInstrumentDate,vchBank,fltAmount From Vouchers Where intKeyID=" & val(txtBankAccounthead.Tag)
        Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
        If Not (Rec.EOF Or Rec.BOF) Then
            vsGrid.Rows = Rec.RecordCount + 1
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColHidden(0) = True
            'vsGrid.ColHidden(3) = True
            vsGrid.ColSel = 3
            vsGrid.RowSel = vsGrid.Rows - 1
            mSql = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = mSql
        End If
        Rec.Close
        
    End Sub
    
    Private Sub FillWard()
        Dim mSql As String
        On Error Resume Next
        mSql = "SELECT chvWardNameEnglish, intWardNo FROM GM_Ward"
        mSql = mSql + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID
        mSql = mSql + " AND numZoneID = " & gbLocationID
        mSql = mSql + " Order By chvWardNameEnglish"
        PopulateList cmbWard, mSql, , , , True, enuSourceString.DBMaster
    End Sub

    Private Sub cmdPTAX_Click()
        frmPTaxCalculator.Show
    End Sub

    Private Sub txtWardNo_Change()
        Dim mCount As Integer
        cmbWard.ListIndex = -1
        For mCount = 0 To cmbWard.ListCount - 1
            If val(txtWardNo.Text) = cmbWard.ItemData(mCount) Then
                cmbWard.ListIndex = mCount
                Exit For
            End If
        Next
    End Sub
    Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtFrom_GotFocus()
        txtFrom.SelStart = 0
        txtFrom.SelLength = Len(txtFrom)
    End Sub

    Private Sub txtFrom_LostFocus()
        If txtFrom.Text <> "" Then
            txtFrom.Text = CheckDateInMMM(txtFrom.Text)
        End If
    End Sub
    Private Sub cmbWard_Click()
        If cmbWard.ListIndex > -1 Then
            txtWardNo.Text = cmbWard.ItemData(cmbWard.ListIndex)
        End If
    End Sub
    Private Sub txtInstrumentNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
        End If
    End Sub

    Private Sub txtTo_GotFocus()
        txtTo.SelStart = 0
        txtTo.SelLength = Len(txtTo)
    End Sub

    Private Sub txtTo_LostFocus()
        '    If mRegularPensionID = "" Or mContigentPensionID = "" Then
        '        If val(cmbAccountHead.Tag) <> mRegularPensionID Or val(cmbAccountHead.Tag) <> mContigentPensionID Then
        '            If cmbPenstionType.ListIndex = 1 Then
        '                'Update Regular Pension Fund Field
        '            ElseIf cmbPenstionType.ListIndex = 2 Then
        '                'Update Contingent Pension Fund Field
        '            End If
        '        Else
        '
        '            If val(cmbAccountHead.Tag) = mRegularPensionID Then
        '                If cmbPenstionType.ListIndex = 0 Then
        '                    'Set Regular Pension Fund Field Null
        '                ElseIf cmbPenstionType.ListIndex = 2 Then
        '                    'Set Regular Pension Fund Field Null
        '                    'Update Contingent Pension Fund Field
        '                End If
        '            End If
        '
        '            If val(cmbAccountHead.Tag) = mContigentPensionID Then
        '                If cmbPenstionType.ListIndex = 0 Then
        '                    'Set Contingent Pension Fund Field
        '                ElseIf cmbPenstionType.ListIndex = 2 Then
        '                    'Set Contingent Pension Fund Field
        '                    'Update Regular Pension Fund Field
        '                End If
        '            End If
        '        End If
    End Sub
