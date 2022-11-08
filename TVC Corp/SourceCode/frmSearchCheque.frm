VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchCheque 
   Caption         =   "Search Cheque For Reverse Entry"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "frmSearchCheque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtToAmt 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   1170
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5340
      Picture         =   "frmSearchCheque.frx":1CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   53
      Top             =   450
      Width           =   480
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   3525
         Left            =   150
         TabIndex        =   54
         Top             =   735
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdSeat 
      Caption         =   "..."
      Height          =   315
      Left            =   5070
      TabIndex        =   49
      Top             =   6690
      Width           =   300
   End
   Begin VB.TextBox txtSeat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   6705
      Width           =   3210
   End
   Begin VB.TextBox txtRemarks 
      Height          =   555
      Left            =   1860
      TabIndex        =   46
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton cmdRequest 
      Caption         =   "Request for Reverse Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7410
      TabIndex        =   45
      Top             =   6600
      Width           =   2445
   End
   Begin VB.CheckBox chkCheque 
      Caption         =   "Cheque Enrolled in Bank A/c Scroll"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5550
      TabIndex        =   44
      Top             =   1740
      Width           =   2865
   End
   Begin VB.Frame fmeBankScroll 
      Enabled         =   0   'False
      Height          =   1755
      Left            =   5550
      TabIndex        =   30
      Top             =   1920
      Width           =   4425
      Begin VB.TextBox txtRemittanceDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox txtRemittance 
         Height          =   285
         Left            =   2580
         TabIndex        =   38
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox txtReturnedDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Top             =   930
         Width           =   1455
      End
      Begin VB.TextBox txtReturned 
         Height          =   285
         Left            =   2580
         TabIndex        =   36
         Top             =   930
         Width           =   1455
      End
      Begin VB.TextBox txtBankChargeDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox txtBankCharge 
         Height          =   285
         Left            =   2580
         TabIndex        =   34
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemittance 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   570
         Width           =   255
      End
      Begin VB.CommandButton cmdRetured 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   930
         Width           =   255
      End
      Begin VB.CommandButton cmdBankCharge 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   1290
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Date                          Scroll Entry        "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1110
         TabIndex        =   43
         Top             =   270
         Width           =   2880
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remittance"
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
         Left            =   135
         TabIndex        =   42
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Returned"
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
         Left            =   285
         TabIndex        =   41
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Charge"
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
         Left            =   30
         TabIndex        =   40
         Top             =   1350
         Width           =   1020
      End
   End
   Begin VB.TextBox txtTotAmount 
      Height          =   285
      Left            =   7950
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6150
      Width           =   1815
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
      TabIndex        =   15
      Top             =   2220
      Width           =   2580
   End
   Begin VB.TextBox txtAmt 
      Height          =   285
      Left            =   1830
      TabIndex        =   7
      Top             =   1140
      Width           =   1455
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   3630
      TabIndex        =   3
      Top             =   420
      Width           =   1455
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1830
      TabIndex        =   1
      Top             =   420
      Width           =   1455
   End
   Begin VB.CommandButton ccmdTransType 
      Caption         =   "..."
      Height          =   255
      Left            =   4830
      TabIndex        =   12
      Top             =   1890
      Width           =   285
   End
   Begin VB.CommandButton cmdAcHead 
      Caption         =   "..."
      Height          =   255
      Left            =   4860
      TabIndex        =   23
      Top             =   3300
      Width           =   255
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1890
      TabIndex        =   24
      Top             =   3675
      Width           =   1320
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1935
      Left            =   450
      TabIndex        =   27
      Top             =   4140
      Width           =   9345
      _cx             =   16484
      _cy             =   3413
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
      FormatString    =   $"frmSearchCheque.frx":1FD4
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Check1"
         Height          =   225
         Left            =   8760
         TabIndex        =   51
         Top             =   30
         Width           =   225
      End
   End
   Begin VB.TextBox txtBankAccounthead 
      Height          =   285
      Left            =   1830
      TabIndex        =   22
      Top             =   3300
      Width           =   2955
   End
   Begin VB.TextBox txtRefNo 
      Height          =   285
      Left            =   1830
      TabIndex        =   20
      Top             =   2940
      Width           =   1455
   End
   Begin VB.TextBox txtDoorNo1 
      Height          =   285
      Left            =   2820
      TabIndex        =   18
      Top             =   2580
      Width           =   495
   End
   Begin VB.TextBox txtDoorNo 
      Height          =   285
      Left            =   1830
      TabIndex        =   17
      Top             =   2580
      Width           =   915
   End
   Begin VB.TextBox txtWardNo 
      Height          =   285
      Left            =   1830
      TabIndex        =   14
      Top             =   2220
      Width           =   435
   End
   Begin VB.TextBox txtTransactiontype 
      Height          =   285
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1860
      Width           =   2955
   End
   Begin VB.TextBox txtPartyName 
      Height          =   285
      Left            =   1830
      TabIndex        =   25
      Top             =   1500
      Width           =   2955
   End
   Begin VB.TextBox txtInstrumentNo 
      Height          =   285
      Left            =   1830
      TabIndex        =   5
      Top             =   810
      Width           =   1455
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3330
      TabIndex        =   8
      Top             =   1200
      Width           =   180
   End
   Begin VB.Label Label19 
      Caption         =   "After Tick Enter Amount/ Date to find The Entry in Bank Scroll"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   5880
      TabIndex        =   56
      Top             =   1080
      Width           =   4275
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "If Check Box is tick then u can verify the Cheque Amount using  Bank Statement."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   5880
      TabIndex        =   55
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Find Transactions for Reverse Entry Using Cheque details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   60
      TabIndex        =   52
      Top             =   30
      Width           =   10275
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarded Seat "
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
      Left            =   645
      TabIndex        =   50
      Top             =   6735
      Width           =   1230
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   1050
      TabIndex        =   47
      Top             =   6240
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total  Amount"
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
      Left            =   6150
      TabIndex        =   29
      Top             =   6210
      Width           =   1020
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   450
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Period"
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
      Left            =   915
      TabIndex        =   0
      Top             =   420
      Width           =   885
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account Head"
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
      TabIndex        =   21
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. No."
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
      Left            =   1155
      TabIndex        =   19
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Door No"
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
      Left            =   1155
      TabIndex        =   16
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ward"
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
      Left            =   1365
      TabIndex        =   13
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type"
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
      TabIndex        =   26
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of Party"
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
      Left            =   720
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label txtAmount 
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
      Left            =   1215
      TabIndex        =   6
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insturment No"
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
      Left            =   720
      TabIndex        =   4
      Top             =   870
      Width           =   1080
   End
End
Attribute VB_Name = "frmSearchCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Private Sub ccmdTransType_Click()
        frmSearchTransactionType.ModeOfTransaction = 1
        frmSearchTransactionType.Show vbModal
        If gbSearchID <> -1 Then
            txtTransactiontype.Text = gbSearchStr
            txtTransactiontype.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub chkSelectAll_Click()
        If chkSelectAll.value = vbChecked Then
            vsGrid.Cell(flexcpChecked, 1, 7, vsGrid.Rows - 1, 7) = True
        Else
            vsGrid.Cell(flexcpChecked, 1, 7, vsGrid.Rows - 1, 7) = False
            txtTotAmount.Text = ""
        End If
    End Sub

    Private Sub chkCheque_Click()
        If chkCheque.value = vbChecked Then
            fmeBankScroll.Enabled = True
        End If
    End Sub

    Private Sub cmdAcHead_Click()
       On Error GoTo err:
            Dim mSQL As String
            mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & 2
            frmSearchAccountHeads.SQLString = mSQL
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

    Private Sub cmdBankCharge_Click()
        Dim mQry As String
        If txtBankAccounthead.Text = "" Then
            MsgBox "Please Select Bank", vbInformation
            txtBankAccounthead.SetFocus
            Exit Sub
        Else
            mQry = "Select intReconciliationID, intBankAccountHeadID, dtBankEntryDate, dtChequeDate,vchChequeNo, vchParticulars , fltCrAmount,fltDrAmount"
            mQry = mQry + " From faBankReconciliationEntries Where intBankAccountHeadID=" & val(txtBankAccounthead.Tag)
            If txtBankCharge.Text <> "" Then
                mQry = mQry + " And fltCrAmount=" & txtBankCharge.Text
            End If
            If txtBankChargeDate.Text <> "" Then
                mQry = mQry + " And dtBankEntryDate='" & Format(txtBankChargeDate.Text, "dd/mmm/yy") & "'"
            End If
        End If
        frmSearchDishonoredCheque.FillGrid mQry, 2
        frmSearchDishonoredCheque.Show vbModal
        If gbSearchCode <> "" Then
            txtBankCharge.Text = gbSearchCode
            txtBankChargeDate.Text = gbSearchStr
            gbSearchCode = ""
            gbSearchStr = ""
        Else
            txtBankCharge.Text = ""
            txtBankChargeDate.Text = ""
        End If
    End Sub

    Private Sub cmdFind_Click()
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim mSQL        As String
        If txtBankAccounthead.Text = "" Then
            MsgBox "Please select Bank account Head"
            txtBankAccounthead.SetFocus
            Exit Sub
        End If
        If txtAmt.Text = "" Then
            MsgBox "Please Enter Cheque Amount"
            txtAmt.SetFocus
            Exit Sub
        End If
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "Select faVouchers.intVoucherID,intVoucherNo,dtDate,vchInstrumentNo,dtInstrumentDate,vchBank,fltAmount From faVouchers "
            mSQL = mSQL + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
            mSQL = mSQL + " Where tnyCancelFlag=0 And tnyVoucherTypeID = 10 And intInstrumentTypeID<>1 And intKeyID1 = " & val(txtBankAccounthead.Tag)
            If txtFrom.Text <> "" And txtTo.Text <> "" Then
                mSQL = mSQL + " And dtDate between '" & txtFrom.Text & "' And '" & txtTo.Text & "'"
            End If
            If txtInstrumentNo <> "" Then
                mSQL = mSQL + " And vchInstrumentNo like'%" & txtInstrumentNo.Text & "%'"
            End If
            If txtAmt.Text <> "" Then
                mSQL = mSQL + " And fltAmount>=" & val(txtAmt.Text)
            End If
            If txtAmt.Text <> "" Then
                mSQL = mSQL + " And fltAmount<=" & val(txtToAmt.Text)
            End If
            If txtTransactiontype.Text <> "" Then
                mSQL = mSQL + " And intTransactionTypeID=" & txtTransactiontype.Tag
            End If
            If txtWardNo.Text <> "" Then
                mSQL = mSQL + " And numWardID=" & txtWardNo.Text
            End If
            If txtDoorNo.Text <> "" Then
                mSQL = mSQL + " And intDoorNoP1=" & txtDoorNo.Text
            End If
            If txtDoorNo1.Text <> "" Then
                mSQL = mSQL + " And intDoorNoP2=" & txtDoorNo1.Text
            End If
            If txtPartyName.Text <> "" Then
                mSQL = mSQL + " And vchName like '%" & txtPartyName.Text & "%'"
            End If
            vsGrid.Rows = 1
            Rec.Open mSQL, mCnn, adOpenStatic, adLockOptimistic
                If Not (Rec.EOF Or Rec.BOF) Then
                    vsGrid.Rows = Rec.RecordCount + 1
                    vsGrid.Col = 0
                    vsGrid.Row = 1
                    vsGrid.ColSel = 7
                    vsGrid.RowSel = vsGrid.Rows - 1
                    mSQL = Rec.GetString(, , vbTab, Chr(13))
                    vsGrid.Clip = mSQL
                End If
        Else
            MsgBox "Connection to Saankhya Does not Exists"
        End If
    End Sub
    
    Private Sub FillWard()
        Dim mSQL As String
        On Error Resume Next
        mSQL = "SELECT chvWardNameEnglish, intWardNo FROM GM_Ward"
        mSQL = mSQL + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID
        mSQL = mSQL + " AND numZoneID = " & gbLocationID
        mSQL = mSQL + " Order By chvWardNameEnglish"
        PopulateList cmbWard, mSQL, , , , True, enuSourceString.DBMaster
    End Sub

    Private Sub cmdRemittance_Click()
        Dim mQry As String
        If txtBankAccounthead.Text = "" Then
            MsgBox "Please Select Bank", vbInformation
            txtBankAccounthead.SetFocus
            Exit Sub
        Else
            mQry = "Select intReconciliationID, intBankAccountHeadID, dtBankEntryDate, dtChequeDate,vchChequeNo, vchParticulars , fltCrAmount"
            mQry = mQry + " From faBankReconciliationEntries Where isnull(fltCrAmount,0)<>0 And intBankAccountHeadID=" & val(txtBankAccounthead.Tag)
            If txtRemittance.Text <> "" Then
                mQry = mQry + " And fltCrAmount=" & txtRemittance.Text & " "
            End If
            If txtRemittanceDate.Text <> "" Then
                mQry = mQry + " And dtBankEntryDate='" & Format(txtRemittanceDate.Text, "dd/mmm/yy") & "'"
            End If
        End If
        frmSearchDishonoredCheque.FillGrid mQry, 1
        frmSearchDishonoredCheque.Show vbModal
        If gbSearchCode <> "" Then
            txtRemittance.Text = gbSearchCode
            txtRemittanceDate.Text = gbSearchStr
            gbSearchCode = ""
            gbSearchStr = ""
        Else
            txtRemittance.Text = ""
            txtRemittanceDate.Text = ""
        End If
    End Sub
'    Private Sub ValidateAmount()
'
'        If chkCheque.Value = vbChecked Then
'            If txtTotAmount.Text > txtRemittance.Text Then
'                MsgBox "Cheue Amount Exceeded", vbApplicationModal
'                Exit Sub
'
'            End If
'        Else
'            If txtTotAmount.Text > txtAmt.Text Then
'                MsgBox "Cheue Amount Exceeded", vbApplicationModal
'                Exit Sub
'            End If
'        End If
'
'    End Sub

    Private Sub cmdRequest_Click()
        On Error GoTo err:
            Dim objDB       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim arrIn       As Variant
            Dim arrOut      As Variant
            Dim mRequestID  As Integer
            Dim mSQL        As String
            Dim mRowCnt     As Integer
            
            If chkCheque.value = vbChecked Then
                    If txtRemittance.Text <> "" Then
                        If txtTotAmount.Text > txtRemittance.Text Then
                            MsgBox "Cheque Amount Exceeded", vbApplicationModal
                            Exit Sub
                        End If
                    ElseIf txtReturned.Text <> "" Then
                        If txtTotAmount.Text > txtReturned.Text Then
                            MsgBox "Cheque Amount Exceeded", vbApplicationModal
                        Exit Sub
                        End If
                    End If
            Else
                If txtTotAmount.Text > txtAmt.Text Then
                    MsgBox "Cheque Amount Exceeded", vbApplicationModal
                    Exit Sub
                End If
            End If

            If txtSeat.Text = "" Then
                MsgBox "Please Select Seat to Forward"
                txtSeat.SetFocus
                Exit Sub
            End If
            If txtRemarks.Text = "" Then
                If (MsgBox("Remark column is left Blank if u want to Enter Press Yes", vbYesNo) = vbYes) Then
                    txtRemarks.SetFocus
                    Exit Sub
                End If
            End If
            If txtInstrumentNo.Text = "" Then
                MsgBox "Please Enter the Instrument No", vbInformation
                txtInstrumentNo.SetFocus
                Exit Sub
            End If
            If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
                If objDB.SetConnection(mCnn) Then
                    arrIn = Array(-1, _
                                gbTransactionDate, _
                                Null, _
                                10, _
                                500, _
                                Trim(txtRemarks.Text), _
                                gbUserID, _
                                gbSeatID, _
                                Null, _
                                Null, _
                                txtSeat.Tag, _
                                gbFinancialYearID, _
                                0)
            
                    
                    
    '                 arrIn = Array(-1, _
    '                            gbTransactionDate, _
    '                            Null, _
    '                            0, _
    '                            500, _
    '                            Trim(txtRemarks.Text), _
    '                            gbUserID, _
    '                            gbSeatID, _
    '                            Null, _
    '                            Null, _
    '                            txtSeat.Tag, _
    '                            gbFinancialYearID, _
    '                            0, _
    '                            Null, _
    '                            Null, _
    '                            Null)
                    
                    objDB.ExecuteSP "spSaveReverseEntry", arrIn, arrOut, , mCnn, adCmdStoredProc
                    
                    If Not IsNumeric(arrOut) Then
                        mRequestID = arrOut(0, 0)
                    End If
                    
                    For mRowCnt = 1 To vsGrid.Rows - 1
                        arrIn = ""
                        If vsGrid.Cell(flexcpChecked, mRowCnt, 7) = vbChecked Then
                            arrIn = Array(mRequestID, val(vsGrid.TextMatrix(mRowCnt, 0)))
                            objDB.ExecuteSP "spSaveReverseEntryChild", arrIn, , , mCnn, adCmdStoredProc
                        End If
                    Next
                    
                    MsgBox "Reverse Entry Request Send to Higher Authority", vbInformation
                    cmdRequest.Enabled = False
                Else
                    MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
                End If
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub



    Private Sub cmdRetured_Click()
        Dim mQry As String
        If txtBankAccounthead.Text = "" Then
            MsgBox "Please Select Bank", vbInformation
            txtBankAccounthead.SetFocus
            Exit Sub
        Else
            mQry = "Select intReconciliationID, intBankAccountHeadID, dtBankEntryDate, dtChequeDate,vchChequeNo, vchParticulars , fltDrAmount"
            mQry = mQry + " From faBankReconciliationEntries Where isnull(fltDrAmount,0) <> 0 And intBankAccountHeadID=" & val(txtBankAccounthead.Tag)
            If txtReturned.Text <> "" Then
               mQry = mQry + " And fltDrAmount=" & txtReturned.Text
            End If
            If txtReturnedDate.Text <> "" Then
                mQry = mQry + " And dtBankEntryDate='" & Format(txtReturnedDate.Text, "dd/mmm/yy") & "'"
            End If
        End If
        frmSearchDishonoredCheque.FillGrid mQry, 2
        frmSearchDishonoredCheque.Show vbModal
        If gbSearchCode <> "" Then
            txtRemittance.Text = gbSearchCode
            txtReturnedDate.Text = gbSearchStr
            gbSearchCode = ""
            gbSearchStr = ""
        Else
            txtRemittance.Text = ""
            txtReturnedDate.Text = ""
        End If
    End Sub

    Private Sub cmdSeat_Click()
        Dim mSQL As String
        mSQL = "Select chvSeatTitle, numSeatID From GL_Seats Where intGroupID in (5,6) And intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        frmSearchSeat.SQLString = mSQL
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtSeat.Text = gbSearchStr
            txtSeat.Tag = gbSearchID
        End If
    End Sub
    Private Sub Form_Load()
        Call FillWard
        txtFrom.Text = CheckDateInMMM(DateAdd("m", -1, Date))
        txtTo.Text = CheckDateInMMM(Date)
    End Sub

    Private Sub txtAmt_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

'    Private Sub txtTotAmount_Change()
'        If txtAmt.Text > txtTotAmount.Text Then
'            MsgBox "Amount Exceeded", vbApplicationModal
'        End If
'    End Sub

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
    Private Sub Calculate()
        Dim mLoop As Long
        Dim mAmtCr As Double
        Dim mAmtDr As Double
        txtTotAmount.Text = ""
        mAmtCr = 0
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 7) = vbChecked Then
                mAmtCr = mAmtCr + Format(val(vsGrid.TextMatrix(mLoop, 6)), "0.00")
                txtTotAmount.Text = mAmtCr
            End If
        Next
        txtTotAmount.Text = mAmtCr
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
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtTo_GotFocus()
        txtTo.SelStart = 0
        txtTo.SelLength = Len(txtTo)
    End Sub

    Private Sub txtTo_LostFocus()
        If txtTo.Text <> "" Then
            txtTo.Text = CheckDateInMMM(txtTo.Text)
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''        If txtAmt.Text > txtTotAmount.Text Then
''            MsgBox "Amount Exceeded"
'''            Cancel = True
''        End If
    End Sub

    Private Sub vsGrid_Click()
        Dim mCount As Integer
        Dim mInstrumentNo As String
        Dim mBank As String
         If vsGrid.Col = 7 Then
            vsGrid.Editable = flexEDKbdMouse
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = vbChecked Then
                
                If CheckReverseRequestExist(vsGrid.TextMatrix(vsGrid.Row, 0)) = 1 Then
                    MsgBox "Already sent Request for this Voucher", vbInformation
                    vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = vbUnchecked
                    Exit Sub
                ElseIf CheckReverseRequestExist(vsGrid.TextMatrix(vsGrid.Row, 0)) = 2 Then
                    MsgBox "This Voucher Already Reversed", vbInformation
                    vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = vbUnchecked
                    Exit Sub
                End If
                
                
            End If
            
            If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = vbChecked Then
                    mInstrumentNo = vsGrid.TextMatrix(vsGrid.Row, 3)
                    mBank = vsGrid.TextMatrix(vsGrid.Row, 5)
                    Call Calculate
                    For mCount = 1 To vsGrid.Rows - 1
                        If vsGrid.TextMatrix(mCount, 5) = mBank And vsGrid.TextMatrix(mCount, 3) = mInstrumentNo Then
                            vsGrid.Cell(flexcpChecked, mCount, 7) = 1
                        Else
                            vsGrid.Cell(flexcpChecked, mCount, 7) = 2
                        End If
                    Next
                End If
                    Call Calculate
''                    For mCount = 1 To vsGrid.Rows - 1
''                        If vsGrid.Cell(flexcpChecked, mCount, 7) = 1 Then
''                            txtTotAmount.Text = Val(txtTotAmount.Text) + Val(vsGrid.TextMatrix(mCount, 6))
''                        End If
''                    Next
                End If
        Else
            vsGrid.Editable = flexEDNone
        End If
    End Sub
    
   Private Function CheckReverseRequestExist(ByVal VchID As Double) As Integer
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            If objDB.SetConnection(mCnn) Then
                mSQL = " Select tnyStatus from faReverseEntry "
                mSQL = mSQL + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
                mSQL = mSQL + " Where intVoucherID =  " & VchID
                mSQL = mSQL + " And tnyStatus<>3"
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyStatus = 0 Then
                        CheckReverseRequestExist = 1
                    ElseIf Rec!tnyStatus = 2 Then
                        CheckReverseRequestExist = 2
                    Else
                        CheckReverseRequestExist = 3
                    End If
                    Exit Function
                Else
                    CheckReverseRequestExist = 3
                End If
            Else
                MsgBox "Connection to Finance does not Exist, Please contact your Sustem Administrator"
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function

