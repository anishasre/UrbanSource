VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmProfessionalTax 
   BackColor       =   &H80000018&
   Caption         =   "Profession Tax Demand"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10665
   Begin VB.Frame framParty 
      BackColor       =   &H00DDF9F9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1725
      Left            =   30
      TabIndex        =   13
      Top             =   0
      Width           =   10620
      Begin VB.TextBox txtInstitution 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   33
         Top             =   270
         Width           =   2550
      End
      Begin VB.TextBox txtLocation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   930
         Width           =   2550
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   5550
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   255
         Width           =   4695
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   2550
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4140
         TabIndex        =   14
         Top             =   915
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   555
         TabIndex        =   20
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Name  && Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   405
         Left            =   4665
         TabIndex        =   18
         Top             =   270
         Width           =   810
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CheckBox chkFineWaiver 
      BackColor       =   &H00ACECEC&
      Caption         =   "Fine Waiver"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3150
      TabIndex        =   12
      Top             =   4905
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CanceL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9075
      TabIndex        =   11
      Top             =   5520
      Width           =   1395
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to Receipt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7170
      TabIndex        =   10
      Top             =   5520
      Width           =   1875
   End
   Begin VB.TextBox txtGrandTotal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8715
      TabIndex        =   9
      Top             =   4800
      Width           =   1305
   End
   Begin VB.TextBox txtAdvance 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8715
      TabIndex        =   8
      Top             =   5100
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtHalfYearTaxRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   4905
      Width           =   1530
   End
   Begin VB.ComboBox cmbFromYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5220
      Width           =   1530
   End
   Begin VB.ComboBox cmbFromPeriod 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3105
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5220
      Width           =   1455
   End
   Begin VB.ComboBox cmbToYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5565
      Width           =   1530
   End
   Begin VB.ComboBox cmbToPeriod 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3105
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5565
      Width           =   1455
   End
   Begin VB.TextBox txtNoOfHalfYears 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   5895
      Width           =   1530
   End
   Begin VB.CommandButton cmdListDemand 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3105
      TabIndex        =   1
      Top             =   5895
      Width           =   420
   End
   Begin VB.TextBox txtFine 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4995
      TabIndex        =   0
      Top             =   4905
      Width           =   1305
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3660
      Top             =   6585
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2745
      Left            =   30
      TabIndex        =   21
      Top             =   1770
      Width           =   10620
      _cx             =   18732
      _cy             =   4842
      Appearance      =   2
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
      BackColor       =   15400959
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   15400959
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProfessionalTax.frx":0000
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   2
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   405
      TabIndex        =   31
      Top             =   4575
      Width           =   1095
   End
   Begin VB.Label lblTotalArrear 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7395
      TabIndex        =   30
      Top             =   4530
      Width           =   1305
   End
   Begin VB.Label lblTotalCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8715
      TabIndex        =   29
      Top             =   4530
      Width           =   1305
   End
   Begin VB.Label Lebel23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   7545
      TabIndex        =   28
      Top             =   4860
      Width           =   1140
   End
   Begin VB.Label lblAdvance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance to be Adjusted"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   6645
      TabIndex        =   27
      Top             =   5145
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Half Year Tax"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   375
      TabIndex        =   26
      Top             =   4950
      Width           =   1155
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Period"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   495
      TabIndex        =   25
      Top             =   5280
      Width           =   1035
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Period"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   720
      TabIndex        =   24
      Top             =   5625
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No.of Half Years"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   150
      TabIndex        =   23
      Top             =   5910
      Width           =   1380
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fine"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   4620
      TabIndex        =   22
      Top             =   4920
      Width           =   345
   End
End
Attribute VB_Name = "frmProfessionalTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public numSubLedgerID As Long

Private Sub cmdCopy_Click()
    Call copyToReceipt
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmProfessionTaxSearch.Visible = True
    frmProfessionTaxSearch.ZOrder (0)
End Sub

Private Sub Form_Load()
    WindowsXPC.InitIDESubClassing
    Me.Height = 6735
    Me.Width = 10785
    Call FillYear
    frmProfessionTaxSearch.Visible = True
End Sub

Public Sub FillGrid(Optional mSubLedgerID As Long)
    Dim str As String
    Dim mCon As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mRow As Integer
    Dim mArrearFlag As Boolean
    Dim mAmtArrear As Double
    Dim mAmtCurrent As Double
    Dim mTotalAmt As Double
    str = "SELECT faAccountHeads.vchAccountHead, faAccountHeads.vchAccountHeadCode, faIDemandTBL.*, faIDemandChild.* " & _
            "FROM         faIDemandTBL INNER JOIN " & _
            "faIDemandChild ON faIDemandTBL.numDemandID = faIDemandChild.numDemandID" & _
            " Inner Join faAccountHeads on faAccountHeads.intAccountHeadID= faIDemandChild.intAccountHeadID" & _
            " Where (faIDemandTBL.intTransactionTypeID = 2) And (faIDemandTBL.numSubLedgerID =" & numSubLedgerID & ")"
    If objdb.SetConnection(mCon) Then
        vsGrid.Rows = 1
        Set Rec = objdb.ExecuteSP(str, , , , mCon, adCmdText)
        While Not Rec.EOF
                mRow = vsGrid.Rows
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!vchAccountHeadCode), " ", Rec!vchAccountHeadCode)
                vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!vchAccountHead), " ", Rec!vchAccountHead)
                vsGrid.Cell(flexcpText, mRow, 2) = CStr(IIf(IsNull(Rec!intYearID), " ", Rec!intYearID)) & " - " & CStr(IIf(IsNull(Rec!intYearID), " ", Rec!intYearID + 1))
                Select Case IIf(IsNull(Rec!tnyPeriodID), 0, Rec!tnyPeriodID)
                    Case Is = 1: vsGrid.Cell(flexcpText, mRow, 3) = "Ist Half"
                    Case Is = 2: vsGrid.Cell(flexcpText, mRow, 3) = "IInd Half"
                    Case Is = 3: vsGrid.Cell(flexcpText, mRow, 3) = "Full Year"
                End Select
                vsGrid.Cell(flexcpText, mRow, 6) = IIf(IsNull(Rec!intAccountHeadID), " ", Rec!intAccountHeadID)
                vsGrid.Cell(flexcpText, mRow, 7) = IIf(IsNull(Rec!intYearID), " ", Rec!intYearID)
                vsGrid.Cell(flexcpText, mRow, 8) = IIf(IsNull(Rec!tnyPeriodID), " ", Rec!tnyPeriodID)
                vsGrid.Cell(flexcpText, mRow, 9) = IIf(IsNull(Rec!tnyArrearFlag), " ", Rec!tnyArrearFlag)
                vsGrid.Cell(flexcpText, mRow, 10) = IIf(IsNull(Rec!numDemandID), " ", Rec!numDemandID)
                vsGrid.Cell(flexcpText, mRow, 11) = IIf(IsNull(Rec!fltAmount), " ", Rec!fltAmount)
                vsGrid.MergeCol(12) = True
                vsGrid.Cell(flexcpText, mRow, 12) = IIf(IsNull(Rec!numDemandID), " ", Rec!numDemandID)
                vsGrid.Cell(flexcpChecked, mRow, 12) = 1
                mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), 0, Rec!tnyArrearFlag)
                If mArrearFlag Then
                    mAmtArrear = mAmtArrear + IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                    vsGrid.Cell(flexcpText, mRow, 4) = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                Else
                    mAmtCurrent = mAmtCurrent + IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                    vsGrid.Cell(flexcpText, mRow, 5) = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                End If
                mTotalAmt = mTotalAmt + IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)

            Rec.MoveNext
        Wend
        lblTotalArrear.Caption = mAmtArrear
        lblTotalCurrent.Caption = mAmtCurrent
        txtGrandTotal.Text = mAmtArrear + mAmtCurrent
        txtFine.Text = CalculateProfessionTaxFine
    End If
End Sub

Private Sub FillYear()
        Dim mLoop As Long
        For mLoop = 1991 To Year(Date)
            cmbFromYear.AddItem CStr(mLoop) & "-" & CStr(mLoop + 1)
            cmbFromYear.ItemData(cmbFromYear.NewIndex) = mLoop
        Next mLoop
        
        For mLoop = 1991 To Year(Date)
            cmbToYear.AddItem CStr(mLoop) & "-" & CStr(mLoop + 1)
            cmbToYear.ItemData(cmbToYear.NewIndex) = mLoop
        Next mLoop
        
        cmbFromPeriod.AddItem "First Half"
        cmbFromPeriod.ItemData(cmbFromPeriod.NewIndex) = 1
        cmbFromPeriod.AddItem "Second Half"
        cmbFromPeriod.ItemData(cmbFromPeriod.NewIndex) = 2
        cmbFromPeriod.AddItem "Full Year"
        cmbFromPeriod.ItemData(cmbFromPeriod.NewIndex) = 3
        
        cmbToPeriod.AddItem "First Half"
        cmbToPeriod.ItemData(cmbToPeriod.NewIndex) = 1
        cmbToPeriod.AddItem "Second Half"
        cmbToPeriod.ItemData(cmbToPeriod.NewIndex) = 2
        
    End Sub
    
    Private Sub copyToReceipt()
        Dim mRow As Integer
        frmReceiptsCounter.txtAddress.Text = txtAddress.Text
        frmReceiptsCounter.txtHouseName.Text = txtInstitution.Text
        mRow = vsGrid.Rows
        frmReceiptsCounter.vsGrid.Rows = 1
        For mRow = 1 To vsGrid.Rows - 1
            frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 1
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = vsGrid.TextMatrix(mRow, 0)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = vsGrid.TextMatrix(mRow, 1)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = vsGrid.TextMatrix(mRow, 2)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = vsGrid.TextMatrix(mRow, 3)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = vsGrid.TextMatrix(mRow, 4)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = vsGrid.TextMatrix(mRow, 5)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = vsGrid.TextMatrix(mRow, 6)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = vsGrid.TextMatrix(mRow, 7)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = vsGrid.TextMatrix(mRow, 8)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = vsGrid.TextMatrix(mRow, 9)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = vsGrid.TextMatrix(mRow, 10)
            frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = vsGrid.TextMatrix(mRow, 11)
        Next
    End Sub
    Private Function CalculateProfessionTaxFine()
        Dim mCount As Integer
        Dim mFine As Double
        mFine = 0
        For mCount = 1 To vsGrid.Rows - 1
            If val(vsGrid.TextMatrix(mCount, 4)) <> 0 Then
               mFine = mFine + (val(vsGrid.TextMatrix(mCount, 4)) * 0.01)
            End If
        Next mCount
        CalculateProfessionTaxFine = mFine
        
    End Function

