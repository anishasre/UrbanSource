VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRemitBackofUnUtilizeddrawnAmounts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RemitBack of UnUtilized drawn Amounts"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   Icon            =   "frmRemitBackofUnUtilizeddrawnAmounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8730
      TabIndex        =   3
      Top             =   7020
      Width           =   1095
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3375
      Left            =   45
      TabIndex        =   0
      Top             =   3600
      Width           =   10695
      _cx             =   18865
      _cy             =   5953
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   13
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRemitBackofUnUtilizeddrawnAmounts.frx":1CCA
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
      TextStyle       =   1
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
   Begin VB.PictureBox PicCaption 
      BackColor       =   &H80000009&
      Height          =   510
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   10725
      TabIndex        =   1
      Top             =   0
      Width           =   10785
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   8550
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Of RemitBack of Unutilized drawn Amounts"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   45
         Width           =   3780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3225
      Left            =   45
      TabIndex        =   4
      Top             =   330
      Width           =   10680
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4635
         TabIndex        =   27
         Top             =   2700
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   2445
         Left            =   5310
         TabIndex        =   17
         Top             =   135
         Width           =   5325
         Begin VB.TextBox txtAllotCategory 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1215
            Width           =   1785
         End
         Begin VB.TextBox txtAllotmentNo 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   180
            Width           =   1785
         End
         Begin VB.CommandButton cmdAllotmentNo 
            Caption         =   ".."
            Height          =   285
            Left            =   3105
            TabIndex        =   21
            Top             =   180
            Width           =   330
         End
         Begin VB.TextBox txtAllotAmount 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   510
            Width           =   1785
         End
         Begin VB.TextBox txtAllotSourceFund 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   855
            Width           =   3945
         End
         Begin VB.TextBox txtAllotExpdHead 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1575
            Width           =   1785
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   495
            TabIndex        =   29
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   630
            TabIndex        =   26
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Source of Fund"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   25
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expd Head"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   450
            TabIndex        =   24
            Top             =   1620
            Width           =   780
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allotment No"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   300
            TabIndex        =   23
            Top             =   225
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2445
         Left            =   45
         TabIndex        =   5
         Top             =   135
         Width           =   5235
         Begin VB.TextBox txtVrDate 
            Height          =   285
            Left            =   3645
            TabIndex        =   30
            Top             =   270
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.TextBox txtVrExpdHead 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1575
            Width           =   1785
         End
         Begin VB.TextBox txtVrSourceFund 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1215
            Width           =   1785
         End
         Begin VB.TextBox txtVrAmount 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   465
            Width           =   1785
         End
         Begin VB.TextBox txtVrTrType 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   825
            Width           =   3945
         End
         Begin VB.CommandButton cmdSearchVr 
            Caption         =   ".."
            Height          =   285
            Left            =   3015
            TabIndex        =   7
            Top             =   135
            Width           =   330
         End
         Begin VB.TextBox txtVoucherNo 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   135
            Width           =   1785
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expd Head"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   330
            TabIndex        =   16
            Top             =   1620
            Width           =   780
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Source of Fund"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   45
            TabIndex        =   15
            Top             =   1215
            Width           =   1110
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   555
            TabIndex        =   14
            Top             =   495
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tr Type"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   510
            TabIndex        =   13
            Top             =   855
            Width           =   570
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voucher No."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   12
            Top             =   135
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmRemitBackofUnUtilizeddrawnAmounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFunctionId    As Integer
Dim mFunctionaryID As Integer
Dim mAccHeadId     As Integer
Dim mAccHeadCode   As Variant
Dim mProjectID As Variant
Dim mProjectNo As Variant
Dim mProjectCost As Variant
Dim mDPCNo As Variant
Dim mDPCDate As Variant
Dim mSubsectorID As Integer
Dim mMicroSectorID As Integer
Dim mSourceOfFundID As Integer
Dim mCategoryID As Integer
Dim mExistingFlag As Boolean
Dim mYearID         As Integer
    Private Function SaveValidate() As Boolean
        SaveValidate = True
        If mYearID < 2012 Then
            MsgBox "Please Select a Valid Year", vbApplicationModal
            SaveValidate = False
            Exit Function
        End If
        
        If val(txtVoucherNo.Tag) < 1 Then
            MsgBox "Please Select Voucher No", vbApplicationModal
            SaveValidate = False
            Exit Function
        End If
        If val(txtAllotmentNo.Tag) < 1 Then
            MsgBox "Please Select Allotment No", vbApplicationModal
            SaveValidate = False
            Exit Function
        End If
        If val(txtVoucherNo.Tag) > 0 And val(txtAllotmentNo.Tag) > 0 Then
            If txtVrSourceFund.Tag <> txtAllotSourceFund.Tag Then
                MsgBox "Source of Fund is different!", vbInformation
                'MsgBox "Source of Fund not matching", vbApplicationModal
                'SaveValidate = False
                'Exit Function
            End If
            If txtVrExpdHead.Tag <> txtAllotExpdHead.Tag Then
                MsgBox "Expenditure Head not matching", vbApplicationModal
                SaveValidate = False
                Exit Function
            End If
            
            'If val(txtAllotAmount.Text) > val(txtVrAmount.Text) Then
            If val(txtAllotAmount.Text) < val(txtVrAmount.Text) Then
                MsgBox "Requisition Amount Should Be Greater than Voucher Amount", vbApplicationModal
                SaveValidate = False
                Exit Function
            End If
        End If

    End Function

    Private Sub cmbYear_Click()
        If cmbYear.ListIndex > -1 Then
            mYearID = cmbYear.ItemData(cmbYear.ListIndex)
        End If
        Call FormInitialize
        Call FillGrid
    End Sub

    Private Sub cmdAdd_Click()
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim arrInput        As Variant
        Dim mSQL            As String
        Dim objVrSub        As uVoucherSub
        
            If SaveValidate Then
                If val(txtVoucherNo.Tag) > 0 And val(txtAllotmentNo.Tag) > 0 Then
                    If objDB.SetConnection(mCnn) Then
                        mSQL = "SELECT intVoucherID FROM faVoucherSub WHERE intVoucherID = " & val(txtVoucherNo.Tag) & "  "
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mSQL = " UPDATE faVoucherSub SET "
                            mSQL = mSQL + " intTypeID =9,intAllotmentID=" & val(txtAllotmentNo.Tag)
                            mSQL = mSQL + " WHERE intVoucherID=" & txtVoucherNo.Tag
                            objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                            MsgBox "Successfully Added", vbApplicationModal
                        Else
                            With objVrSub
                                .intVoucherID = val(txtVoucherNo.Tag)
                                .decProjectID = Null
                                .intSourceOfFundID = val(txtAllotSourceFund.Tag)
                                .intCategoryID = Null
                                .intSectorID = Null
                                .intAllotmentID = val(txtAllotmentNo.Tag)  ''Letter of Authority ID
                                .intAgreementID = Null
                                .intCashBookID = Null
                                .intImplementingOfficerID = Null
                                .intCreditorTypeID = Null
                                .intCreditorsID = Null
                                .intTypeID = 9                  'To Identify Receipts with type Remittback of UnUtilized Amount
                                .intLocalBodyID = gbLocalBodyID
                                
                                arrInput = Array(.intVoucherID, _
                                                .intLocalBodyID, _
                                                .decProjectID, _
                                                .intSourceOfFundID, _
                                                .intCategoryID, _
                                                .intSectorID, _
                                                .intAllotmentID, _
                                                .intAgreementID, _
                                                .intCashBookID, _
                                                .intImplementingOfficerID, _
                                                .intCreditorTypeID, _
                                                .intCreditorsID, _
                                                .intTypeID)
                                objDB.ExecuteSP "spSaveVoucherSub", arrInput, , , mCnn
                                MsgBox "Successfully Added", vbApplicationModal
                            End With
                        End If
                        Call FillGrid
                        Rec.Close
                        
                    End If
                End If
            End If
    End Sub

    Private Sub cmdAllotmentNo_Click()
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSQL            As String
        Dim objAllotment    As New clsAllotmentLetter
        Dim objAcc          As New clsAccounts
        
        If mYearID < 2012 Then
            Exit Sub
        Else
            If mYearID = gbFinancialYearID - 1 Then
                frmListOfAllotmentLetters.PreviousYearMode = 1
            Else
                frmListOfAllotmentLetters.PreviousYearMode = 0
            End If
        End If
        frmListOfAllotmentLetters.RemitBackMode = 1
        'frmListOfAllotmentLetters.PreviousYearMode = 1
        
        frmListOfAllotmentLetters.Show vbModal
        If gbSearchID <> -1 Then
            txtAllotmentNo.Text = gbSearchCode
            txtAllotmentNo.Tag = gbSearchID
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
            If val(txtAllotmentNo.Tag) > 0 Then
                objAllotment.SetAllotment (txtAllotmentNo.Tag)
                txtAllotSourceFund.Text = IIf(IsNull(objAllotment.SourceOfFund), "", objAllotment.SourceOfFund)
                txtAllotSourceFund.Tag = IIf(IsNull(objAllotment.SourceOfFundID), -1, objAllotment.SourceOfFundID)
                txtAllotAmount.Text = IIf(IsNull(objAllotment.Amount), "", objAllotment.Amount)
                txtAllotExpdHead.Tag = IIf(IsNull(objAllotment.GrossAccountHeadID), -1, objAllotment.GrossAccountHeadID)
                objAcc.SetAccountID (val(txtAllotExpdHead.Tag))
                txtAllotExpdHead.Text = IIf(IsNull(objAcc.AccountHead), "", objAcc.AccountHead)
                
                 If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                    mSQL = " SELECT * FROM faAllotments "
                    mSQL = mSQL + " INNER JOIN faTransactionCategory On faTransactionCategory.intCategoryID=faAllotments.intFundCategoryID "
                    mSQL = mSQL + " WHERE faAllotments.intID = " & val(txtAllotmentNo.Tag) & " "
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        txtAllotCategory.Tag = IIf(IsNull(Rec!intCategoryID), -1, Rec!intCategoryID)
                        txtAllotCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    End If
                End If
            End If
        End If

            
    End Sub

    Private Sub cmdSearchVr_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mRec        As New ADODB.Recordset
        Dim mSQL        As String
        
        
        If mYearID < 2012 Then
            Exit Sub
        End If
        
        frmSearchVouchers.PreviousYearMode = 0
        frmSearchVouchers.CheckMode = 10
        'frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbStartingDate))
        'frmSearchVouchers.txtToDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
        
        
        frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateSerial(mYearID, 4, 1))
        frmSearchVouchers.txtToDate.Text = DdMmmYy(DateSerial(mYearID + 1, 3, 31))
        
        
        frmSearchVouchers.chkContra.Visible = False
        frmSearchVouchers.chkReceipt.Visible = True
        frmSearchVouchers.chkReceipt.value = 1
        frmSearchVouchers.chkInterrupted.Visible = True
        frmSearchVouchers.chkInterrupted.value = 1
        
        frmSearchVouchers.chkJournal.Visible = False
        frmSearchVouchers.chkPayment.Visible = False
        frmSearchVouchers.txtFromDate.Enabled = False
        frmSearchVouchers.txtToDate.Enabled = False
''''        frmSearchVouchers.txtTransactionType.Text = "Remit Back of Unutilized drawn Amounts"
''''        frmSearchVouchers.txtTransactionType.Tag = 172
        frmSearchVouchers.Show vbModal
        If gbSearchID <> -1 Then
            txtVoucherNo.Text = gbSearchCode
            txtVoucherNo.Tag = gbSearchID
            gbSearchCode = ""
            gbSearchID = -1
        
            If val(txtVoucherNo.Tag) > 0 Then
                If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                    mSQL = " SELECT faVouchers.intFinancialYearID intFinYearID, * FROM faVouchers "
                    mSQL = mSQL + " INNER JOIN faAccountHeads On faAccountHeads.intAccountHeadID=faVouchers.intKeyID1"
                    mSQL = mSQL + " INNER JOIN faTransactionType On faTransactionType.intTransactionTypeID=faVouchers.intTransactionTypeID "
                    mSQL = mSQL + " INNER JOIN faVoucherSub On faVoucherSub.intVoucherID=faVouchers.intVoucherID"
                    mSQL = mSQL + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faVoucherSub.intSourceOFFundID"
                    mSQL = mSQL + " WHERE faVouchers.intVoucherID = " & txtVoucherNo.Tag & " "
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                    
                         If mYearID <> Rec("intFinYearID") Then
                            MsgBox "Selected Voucher Not belongs the Financial Year Selected!", vbInformation
                            Exit Sub
                         End If
                    
                         mSQL = " Select * From faVoucherchild "
                         mSQL = mSQL + " INNER JOIN faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID "
                         mSQL = mSQL + "Where intVoucherID=" & txtVoucherNo.Tag & " "
                        
                         mRec.CursorLocation = adUseClient
                         mRec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
                         If Not (mRec.EOF And mRec.BOF) Then
                            If mRec.RecordCount = 1 Then
                                txtVrExpdHead.Text = IIf(IsNull(mRec!vchAccountHead), "", mRec!vchAccountHead)
                                txtVrExpdHead.Tag = IIf(IsNull(mRec!intAccountHeadID), -1, mRec!intAccountHeadID)
                            Else
                                MsgBox "More than one expenditure Head, You can't Select this Voucher", vbApplicationModal
                                txtVoucherNo.Text = ""
                                txtVoucherNo.Tag = ""
                                Exit Sub
                            End If
                         End If
                         mRec.Close
                         txtVrDate.Text = DdMmmYy(IIf(IsNull(Rec!dtDate), 0, Rec!dtDate))
                         txtVrAmount.Text = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                         txtVrTrType.Tag = IIf(IsNull(Rec!intTransactionTypeID), -1, Rec!intTransactionTypeID)
                         txtVrTrType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                         txtVrSourceFund.Tag = IIf(IsNull(Rec!intSourceOfFundID), -1, Rec!intSourceOfFundID)
                         txtVrSourceFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    End If
                    Rec.Close
                 End If
                 
            End If
        End If
    End Sub

    Private Sub cmdVerify_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSQL        As String
        Dim mCnt        As Integer
        Dim mCheck      As Boolean
        Dim Reqn        As uRequisition
        Dim arrInput    As Variant
        Dim arrOutPut   As Variant
        Dim mCnnSulekha As New ADODB.Connection
        
        
        If objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha) = False Then
            MsgBox "Plz Check Connection to Sulekha DB, plz", vbInformation
            Exit Sub
        End If
        
        ''''  Insert int to allotments
        mCheck = False
        For mCnt = 0 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCnt, 0) <> "" Then
                If vsGrid.Cell(flexcpChecked, mCnt, 4) = vbChecked Then
                    mCheck = True
                    Exit For
                End If
            End If
        Next
        If mCheck Then
            For mCnt = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mCnt, 0) <> "" Then
                    If vsGrid.TextMatrix(mCnt, 7) = 0 And vsGrid.Cell(flexcpChecked, mCnt, 4) = vbChecked Then  ' Modified By Aiby on 04-July,2014
                        'If vsGrid.Cell(flexcpChecked, mCnt, 4) = vbChecked Then
                        mSQL = " Update faVouchers Set intTransactionTypeID=172 Where intVoucherID=" & val(vsGrid.TextMatrix(mCnt, 5))
                        mSQL = " Update faTransactions set intTransactionTypeID=172 Where intVoucherID=" & val(vsGrid.TextMatrix(mCnt, 5))
                       '--------------Save faAllotments-------------------------------
                        Call GetRequistionDetails(val(vsGrid.TextMatrix(mCnt, 6)), val(vsGrid.TextMatrix(mCnt, 5)))
                        'If Not mExistingFlag Then
                            
                                
                            
                                With Reqn
                                    .tnyStage = 2
                                    .vchRequisition = Null
                                    .dtRequisitionDate = CDate(DdMmmYy(vsGrid.TextMatrix(mCnt, 1)))
                                    .intFinancialYearID = vsGrid.TextMatrix(mCnt, 8) 'mYearID ' gbFinancialYearID - 1
                                    .intImplementingOfficersID = Null
                                    .vchDesignation = Null
                                    .vchNameofIMPO = Null
                                    .vchPlace = Null
                                    .vchDepartment = Null
                                    .vchDDOCode = Null
                                    .fltRequestedAmt = -1 * val(vsGrid.TextMatrix(mCnt, 2))
                                    .tnyPlanOrNonPlan = 1
                                    .numProjectID = mProjectID
                                    .numProjectNo = mProjectNo
                                    .fltProjectCost = mProjectCost
                                    .vchDPCApprovalNo = mDPCNo
                                    .dtDPCDate = mDPCDate
                                    .intSourceID = mSourceOfFundID
                                    .intCategoryID = mCategoryID
                                    .intTreasuryID = Null
                                    .vchTreasuryCode = Null
                                    .vchTreasuryName = Null
                                    .vchGHeadofAccount = Null
                                    .vchGBudgetHead = Null
                                    .vchGDemandNo = Null
                                    .intFunctionaryID = mFunctionaryID
                                    .intFunctionID = mFunctionId
                                    .intAccountHeadID = mAccHeadId
                                    .vchAccountHeadCode = mAccHeadCode
                                    .intLBID = gbLocalBodyID
                                    .tnyStatus = 1
                                    .tnyInstallmentNo = Null
                                    .intSchemeID = Null
                                    .intSubSecID = mSubsectorID
                                    .intMircoSectorID = mMicroSectorID
                            
                                arrInput = Array("", .tnyStage, .vchRequisition, _
                                            .dtRequisitionDate, _
                                            .intImplementingOfficersID, _
                                            .vchDesignation, _
                                            .vchNameofIMPO, _
                                            .vchPlace, _
                                            .vchDepartment, _
                                            .vchDDOCode, _
                                            .fltRequestedAmt, _
                                            .tnyPlanOrNonPlan, _
                                            .numProjectID, _
                                            .numProjectNo, _
                                            .fltProjectCost, _
                                            .vchDPCApprovalNo, _
                                            .dtDPCDate, _
                                            .intSourceID, _
                                            .intCategoryID, _
                                            .intTreasuryID, _
                                            .vchTreasuryCode, _
                                            .vchTreasuryName, _
                                            .vchGHeadofAccount, _
                                            .vchGBudgetHead, _
                                            .vchGDemandNo, _
                                            .intFunctionaryID, .intFunctionID, .intAccountHeadID, .vchAccountHeadCode, .intLBID, .intFinancialYearID, .tnyStatus, Null, Null, Null, Null, Null, Null, Null, Null, .tnyInstallmentNo, _
                                             Null, Null, Null, Null, Null, Null, Null, Null, Null, .intSchemeID, .intSubSecID, .intMircoSectorID, 2, val(vsGrid.TextMatrix(mCnt, 5)))
                             If Not mExistingFlag Then
                                    objDB.ExecuteSP "spSaveAllotmentRequisition", arrInput, arrOutPut, True, mCnn, adCmdStoredProc
                             End If
                                    
                                    arrInput = Array(gbLBID, _
                                    .intFinancialYearID, _
                                    .numProjectID, _
                                    -1, .intSourceID, _
                                    .fltRequestedAmt, _
                                    val(vsGrid.TextMatrix(mCnt, 5)), _
                                    CDate(vsGrid.TextMatrix(mCnt, 1)))
                                    
                                    objDB.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
                                    
                                End With
                                
                                'Call UpdateDetailsToSulekha
                                
                                
                                
                        'End If ' mExistsFlag :: Its already exists
                    End If
                End If
            Next
        Else
            MsgBox "Please tick Verify Column ", vbApplicationModal
        End If
        
        
        '''' Update trtype in vouchers and transactions
    End Sub

    Private Sub Form_Load()
        Call FillGrid
        Call FillYear
    End Sub
    Private Sub FillGrid()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSQL        As String
        Dim mCnt        As Integer
        Dim objAllot    As New clsAllotmentLetter
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSQL = "Select * From faVouchers "
            mSQL = mSQL + " Inner Join faVoucherSub On faVouchers.intVoucherID=faVoucherSub.intVoucherID"
            mSQL = mSQL + " Where faVoucherSub.intTypeID=9"
            Rec.CursorLocation = adUseClient
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            mCnt = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mCnt, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                vsGrid.TextMatrix(mCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtDate), "", Rec!dtDate))
                vsGrid.TextMatrix(mCnt, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                vsGrid.TextMatrix(mCnt, 5) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                vsGrid.TextMatrix(mCnt, 6) = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                objAllot.SetAllotment (val(vsGrid.TextMatrix(mCnt, 6)))
                vsGrid.TextMatrix(mCnt, 3) = IIf(IsNull(objAllot.AllotmentNo), "", objAllot.AllotmentNo)
                If CheckTransferToSulekha(IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)) = True Then
                    vsGrid.TextMatrix(mCnt, 4) = vbChecked
                    vsGrid.TextMatrix(mCnt, 7) = 1 ' ADDED by Aiby on 04-JULY-2014
                Else
                    vsGrid.TextMatrix(mCnt, 4) = vbUnchecked
                    vsGrid.TextMatrix(mCnt, 7) = 0 ' ADDED by Aiby on 04-JULY-2014
                End If
                vsGrid.TextMatrix(mCnt, 8) = IIf(IsNull(Rec!intFinancialYearID), "", Rec!intFinancialYearID)
                Rec.MoveNext
                mCnt = mCnt + 1
            Wend
            Rec.Close
        End If
    End Sub
    Private Sub GetRequistionDetails(mAllotmentID As Integer, mVoucherID As Double)
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSQL            As String
       
     
        If objDB.SetConnection(mCnn) Then
            If mAllotmentID > 0 Then
                'Check whether this record is already updated :: ADDED BY AIBY on 13th JUNE, 2013
                mExistingFlag = False
                mSQL = "Select * From faAllotments WHERE tnyTypeID = 2 And intVoucherID = " & mVoucherID
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mExistingFlag = True
                    'Exit Sub ' MODIFIED BY AIBY ON 04-July,2014
                End If
                Rec.Close
                
                mSQL = " SELECT * FROM faAllotments "
                mSQL = mSQL + " INNER JOIN faFunctionaries ON faAllotments.intFunctionaryID = faFunctionaries.intFunctionaryID"
                mSQL = mSQL + " INNER JOIN faFunctions ON faAllotments.intFunctionID = faFunctions.intFunctionID"
                mSQL = mSQL + " WHERE intID= " & mAllotmentID & " "
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mFunctionId = IIf(IsNull(Rec!intFunctionID), 0, Rec!intFunctionID)
                    mFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), 0, Rec!intFunctionaryID)
                    mAccHeadId = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
                    mAccHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), 0, Rec!vchAccountHeadCode)
                    mProjectID = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
                    mProjectNo = IIf(IsNull(Rec!vchProjectNo), 0, Rec!vchProjectNo)
                    mProjectCost = IIf(IsNull(Rec!fltProjectCost), 0, Rec!fltProjectCost)
                    mDPCNo = IIf(IsNull(Rec!vchDPCApprovalNo), 0, Rec!vchDPCApprovalNo)
                    'mDPCDate = DdMmmYy(IIf(IsNull(Rec!dtDPCDate), 0, Rec!dtDPCDate))
                    mDPCDate = Rec!dtDPCDate                                             ' CHANGED BY AIBY
                    mSubsectorID = IIf(IsNull(Rec!intSubSecID), 0, Rec!intSubSecID)
                    mMicroSectorID = IIf(IsNull(Rec!intMircoSectorID), 0, Rec!intMircoSectorID)
                    mSourceOfFundID = IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID)
                    mCategoryID = IIf(IsNull(Rec!intFundCategoryID), 0, Rec!intFundCategoryID)
                 End If
                Rec.Close
            End If
       End If
    End Sub
    Private Sub UpdateDetailsToSulekha()
        Dim mCnnSulekha   As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL As String
        Dim arrInput As Variant
        
        '[ExpenseDetails_I]
        '
        '@intLBID int,
        '@intYearID int,
        '@decProjectID numeric,
        '@intSlNo int,
        '@intFundSrcID int,
        '@fltAmt float,
        '@intVoucherID bigint
        
        If val(txtVoucherNo.Tag) > 0 Then
            If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                arrInput = Array(gbLBID, _
                               mYearID, _
                               mProjectID, _
                               -1, mSourceOfFundID, _
                               val(txtVrAmount) * -1, _
                               val(txtVoucherNo.Tag), CDate(txtVrDate.Text))
                               
                'gbFinancialYearID - 1, _ >> Changed with mYearID in the INPUT
                
               objDB.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
            Else
               MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
               Exit Sub
            End If
            mCnnSulekha.Close
        End If
    End Sub
    
    Private Function CheckTransferToSulekha(mVoucherID As Variant) As Boolean
        Dim mCnnSulekha   As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL As String
        Dim arrInput As Variant
        Dim Rec As New ADODB.Recordset
        
        If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
            mSQL = "Select * from ExpenseDetails where intVoucherID = " & mVoucherID & "  "
            Rec.Open mSQL, mCnnSulekha
            If Not (Rec.EOF And Rec.BOF) Then
                CheckTransferToSulekha = True
            Else
                CheckTransferToSulekha = False
            End If
            Rec.Close
        Else
            MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
            Exit Function
        End If
        mCnnSulekha.Close
    End Function
        
    Private Sub FillYear()
        PopulateList cmbYear, "Select Cast(intFinancialYearID as varchar(4)) + '-' + Right(Cast(intFinancialYearID+1 as varchar(4)),2),intFinancialYearID  From faFinancialYear WHERE intFinancialYearID > 2011", , , , True
        cmbYear.ListIndex = cmbYear.ListCount - 1
        vsGrid.SelectionMode = flexSelectionByRow
    End Sub
    Private Sub FormInitialize()
        Dim ctrl    As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            ElseIf TypeOf ctrl Is OptionButton Then
                ctrl.value = False
            ElseIf TypeOf ctrl Is ComboBox Then
                'If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                'ctrl.Tag = ""
            End If
        Next
    End Sub
