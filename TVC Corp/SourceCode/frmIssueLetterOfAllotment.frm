VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIssueLetterOfAllotment 
   Caption         =   "Issue Letter of Allotment"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   12420
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraAllotments 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1FDFD&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6585
      Left            =   -15
      TabIndex        =   8
      Top             =   840
      Width           =   12495
      Begin VB.TextBox txtNewMode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox MaskAccHead 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4185
         TabIndex        =   39
         Top             =   3195
         Width           =   3075
      End
      Begin VB.TextBox MaskDetailAccHead 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4185
         TabIndex        =   38
         Top             =   3555
         Width           =   3075
      End
      Begin VB.TextBox txtAuthorisedDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9675
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3825
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtScheme 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2850
         Width           =   6525
      End
      Begin VB.TextBox txtCrTreasuryAccountID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9675
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   3480
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   1650
         Left            =   2055
         TabIndex        =   33
         Top             =   4215
         Width           =   8685
         _cx             =   15319
         _cy             =   2910
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
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
         Rows            =   5
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmIssueLetterOfAllotment.frx":0000
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
      Begin VB.TextBox txtAmountInWords 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2505
         Width           =   2010
      End
      Begin VB.TextBox txtLSGIName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8490
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1125
         Width           =   2235
      End
      Begin VB.TextBox txtLSGICode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1125
         Width           =   1620
      End
      Begin VB.TextBox txtIMPOName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1470
         Width           =   6540
      End
      Begin VB.TextBox txtRequisition 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   435
         Width           =   1620
      End
      Begin VB.TextBox txtInstalmentNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   780
         Width           =   1620
      End
      Begin VB.TextBox txtRequisitiontDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9105
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   450
         Width           =   1620
      End
      Begin VB.TextBox txtAmountAuthorized 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Width           =   2010
      End
      Begin VB.TextBox txtTreasury 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1815
         Width           =   4305
      End
      Begin VB.TextBox txtTreasuryCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1815
         Visible         =   0   'False
         Width           =   1425
      End
      Begin MSMask.MaskEdBox MaskDetailAccHead1 
         Height          =   225
         Left            =   9990
         TabIndex        =   36
         Top             =   3195
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   60
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department/Scheme/Programe"
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
         Left            =   1365
         TabIndex        =   29
         Top             =   2835
         Width           =   2760
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount in Words"
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
         Left            =   2610
         TabIndex        =   31
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of LSGI"
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
         Left            =   7245
         TabIndex        =   26
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issuing authority LSGI Code"
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
         Left            =   1710
         TabIndex        =   24
         Top             =   1110
         Width           =   2430
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Implementing Officer"
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
         Left            =   1455
         TabIndex        =   23
         Top             =   1485
         Width           =   2670
      End
      Begin VB.Label lblInstalmentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instalment No"
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
         Left            =   2940
         TabIndex        =   22
         Top             =   765
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dfdsfdsf"
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
         Left            =   -585
         TabIndex        =   21
         Top             =   570
         Width           =   60
      End
      Begin VB.Label lblAllotmentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2880
         TabIndex        =   20
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label lblAllotmentDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Date"
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
         Left            =   8580
         TabIndex        =   19
         Top             =   450
         Width           =   465
      End
      Begin VB.Label lblAmountInFigures 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount in Figures"
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
         Left            =   2520
         TabIndex        =   18
         Top             =   2175
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treasury"
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
         Left            =   3345
         TabIndex        =   17
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head of Account(State Budget)"
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
         Left            =   1440
         TabIndex        =   32
         Top             =   3240
         Width           =   2715
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detailed Head of Account (Apx. IV)"
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
         Left            =   1110
         TabIndex        =   35
         Top             =   3570
         Width           =   3060
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   8835
         TabIndex        =   16
         Top             =   1815
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12360
      TabIndex        =   5
      Top             =   7365
      Width           =   12420
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4988
         TabIndex        =   7
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6218
         TabIndex        =   6
         Top             =   90
         Width           =   1215
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   -3195
         Top             =   375
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve"
      Height          =   420
      Left            =   600
      TabIndex        =   4
      Top             =   7890
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAuthorizedAmt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7740
      TabIndex        =   3
      Top             =   2625
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12420
      TabIndex        =   0
      Top             =   0
      Width           =   12420
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is to Issue the Letter of Allotment for the release of fund."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1275
         TabIndex        =   2
         Top             =   480
         Width           =   5265
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   10800
         Picture         =   "frmIssueLetterOfAllotment.frx":0168
         Top             =   -15
         Width           =   1200
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letter of Allotment :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   1
         Top             =   105
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmIssueLetterOfAllotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim ReqID   As Variant
    Dim numTotalAllotmentIssuedToIMPO As Variant
    Dim numTotalExpenditureExludingThisBill As Variant
    Dim mExpExcluding As Double
    Dim mPreviousYearMode As Integer
    Dim mSourceID As Integer
    
    Dim mLoadModeUnAuth As Integer '10-For UNAUTHORIZED DRAWAL
    
    '*********************************************************************************************'
    '                           Form to issue Letter of Allotment                                 '
    '*********************************************************************************************'

    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        vsGrid.TextMatrix(1, 1) = ""
        vsGrid.TextMatrix(2, 1) = ""
        vsGrid.TextMatrix(3, 1) = ""
        vsGrid.TextMatrix(4, 1) = ""
    End Sub
    Private Sub CalculateBAlanceAmount()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        Dim mYearID As Integer
        Dim mArr    As Variant
        
       
        
        If mPreviousYearMode = 0 Then
            mYearID = gbFinancialYearID
        Else
            mYearID = gbFinancialYearID - 1
        End If
        
        
        numTotalAllotmentIssuedToIMPO = Null
        numTotalExpenditureExludingThisBill = Null
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        ' 2 : TOTAL ALLOTMENT ISSUED TO ALL IMPLEMENTING OFFICERS
     
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21, 27, 28, 19, 21, 10, 11, 12, 13, 14        '21 -Best Panchayat
        mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 "
        mSQL = mSQL + " AND intSourceID IN (1,21) AND intFinancialYearID=" & mYearID & " "
        mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
        mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)   AND ISNULL(intTreasuryID,0)=1"
  
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        Rec.Close
        End Select
         
        ' 1: TOTAL ALLOTMENT RECEIVED
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21, 27, 28, 19, 21, 10, 11, 12, 13, 14
        mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
        mSQL = mSQL + " LEFT JOIN faIDemandTBL ON   faIDemandTBL.numsubLedgerID =  faAllotmentLetters.intAllotmentID"
        mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
        mSQL = mSQL + " AND intSourceofFundID IN (1,21)"
        mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
        mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
        mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
        mSQL = mSQL + " AND ISNULL(faIDemandTBL.intVoucherID,0)=0"
        mSQL = mSQL + "  )A"
        

        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!AmountReceived), "0", Rec!AmountReceived)
        End If
        Rec.Close
        End Select
        
        'CURRENT ALLOTMENT
        vsGrid.TextMatrix(3, 1) = val(txtAmountAuthorized.Text)
       
        'BALANCE AVAILABLE
        vsGrid.TextMatrix(4, 1) = val(vsGrid.TextMatrix(1, 1)) - val(vsGrid.TextMatrix(2, 1)) - val(vsGrid.TextMatrix(3, 1))
    
        ' TOTAL ALLOCATION FOR THE IMPLEMENTING OFFICER IN THE CURRENT YEAR
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21, 27, 28, 19, 21, 10, 11, 12, 13, 14
        mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
        mSQL = mSQL + " AND intSourceID IN (1,21) AND intFinancialYearID=" & mYearID & ""
        mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
        mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)  AND ISNULL(intTreasuryID,0)=1"

        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             numTotalAllotmentIssuedToIMPO = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        mExpExcluding = Abs(numTotalAllotmentIssuedToIMPO - vsGrid.TextMatrix(1, 1))
        Rec.Close
        End Select
        
        
        
        
    End Sub
    
    Private Sub CalculateAmountNewMode()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        Dim mYearID As Integer
        Dim mArr    As Variant
        
        
        
        If mPreviousYearMode = 0 Then
            mYearID = gbFinancialYearID
        Else
            mYearID = gbFinancialYearID - 1
        End If
        
        numTotalAllotmentIssuedToIMPO = Null
        numTotalExpenditureExludingThisBill = Null
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        
        '--------------------------CHECK SOURCEWISE BALANCE--------------------
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21, 27, 28, 19, 21, 10, 11, 12, 13, 14
            mArr = Array(val(txtInstalmentNo.Tag), val(txtRequisitiontDate.Tag), Null, mYearID)
            Set Rec = objDB.ExecuteSP("spCheckSourceWiseACRBalance", mArr, , True, mCnn, adCmdStoredProc)
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!fltBalance > 0 Then
                    Call CalculateBAlanceAmount
                    Exit Sub
                End If
            End If
            Rec.Close
        End Select
        '-----------------------------------------------------------------------
       
        ' 2 : TOTAL ALLOTMENT ISSUED TO ALL IMPLEMENTING OFFICERS
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (1,21) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)   AND ISNULL(intTreasuryID,0)=1"
        Case Else
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)   AND ISNULL(intTreasuryID,0)=1"
        End Select
        
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        Rec.Close
         
        ' 1: TOTAL ALLOTMENT RECEIVED
         Dim mSourceOfFundID As Integer
         mSourceOfFundID = val(txtInstalmentNo.Tag)
        
        If mSourceOfFundID = 10 Or mSourceOfFundID = 11 Or mSourceOfFundID = 12 Or mSourceOfFundID = 13 Or mSourceOfFundID = 14 Then
             
             If val(txtRequisitiontDate.Tag) = 1 Then
                mSourceOfFundID = 1
             ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                mSourceOfFundID = 29
             ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                mSourceOfFundID = 30
             End If
        End If
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21
             mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " LEFT JOIN faIDemandTBL ON   faIDemandTBL.numsubLedgerID =  faAllotmentLetters.intAllotmentID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID IN (1,21)"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
            mSQL = mSQL + " AND ISNULL(faIDemandTBL.intVoucherID,0)=0"
            mSQL = mSQL + "  )A"
        Case Else
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " LEFT JOIN faIDemandTBL ON   faIDemandTBL.numsubLedgerID =  faAllotmentLetters.intAllotmentID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID=" & mSourceOfFundID & "  "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
            mSQL = mSQL + " AND ISNULL(faIDemandTBL.intVoucherID,0)=0"
            mSQL = mSQL + "  )A"
        End Select

        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!AmountReceived), "0", Rec!AmountReceived)
        End If
        Rec.Close
                
        'CURRENT ALLOTMENT
        vsGrid.TextMatrix(3, 1) = val(txtAmountAuthorized.Text)
       
        'BALANCE AVAILABLE
        vsGrid.TextMatrix(4, 1) = val(vsGrid.TextMatrix(1, 1)) - val(vsGrid.TextMatrix(2, 1)) - val(vsGrid.TextMatrix(3, 1))
    
        ' TOTAL ALLOCATION FOR THE IMPLEMENTING OFFICER IN THE CURRENT YEAR
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 21
             mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
            mSQL = mSQL + " AND intSourceID IN (1,21) AND intFinancialYearID=" & mYearID & ""
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)  AND ISNULL(intTreasuryID,0)=1"
        Case Else
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)  AND ISNULL(intTreasuryID,0)=1"
        End Select
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             numTotalAllotmentIssuedToIMPO = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        mExpExcluding = Abs(numTotalAllotmentIssuedToIMPO - vsGrid.TextMatrix(1, 1))
        Rec.Close

    End Sub
       
    
    Private Sub CalculateAmount()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        
        numTotalAllotmentIssuedToIMPO = Null
        numTotalExpenditureExludingThisBill = Null
        
        Dim mYearID As Integer
        
        If mPreviousYearMode = 0 Then
            mYearID = gbFinancialYearID
        Else
            mYearID = gbFinancialYearID - 1
        End If
        
        '*********************************************************************************************'
        'Procedure to Calculate the Total Allotment Received, Total Amount Issued & Balance Available '
        '*********************************************************************************************'
        'On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        
        ' 2 : TOTAL ALLOTMENT ISSUED TO ALL IMPLEMENTING OFFICERS
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' 4         Own Fund                                                                       '
        ' 1,27,28   Development Fund- Special Grant, Road renovation                               '
        ' 16,17-    Maintenance                                                                    '
        ' 25        CFC Grant                                                                      '
        ' 26        KLGSDP Grant                                                                   '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 19, 21              '21 -Best Panchayat
            If val(txtRequisitiontDate.Tag) = 1 Or val(txtRequisitiontDate.Tag) = 2 Or val(txtRequisitiontDate.Tag) = 3 Then

                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1"
                mSQL = mSQL + " AND intSourceID IN (21,27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID= " & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "'  And  tnyStage = 2"
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND intFundCategoryID = 1"
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " Union All"
                
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1"
                mSQL = mSQL + " AND intSourceID IN (1) AND intFinancialYearID= " & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2"
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " )A"
            End If
       
       Case 10, 11, 12, 13, 14
            If val(txtRequisitiontDate.Tag) = 1 Then
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
                mSQL = mSQL + " AND intSourceID IN (21,27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND intFundCategoryID = 1"
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
                mSQL = mSQL + " AND intSourceID IN (1) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " )A"
            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                GoTo SCP:
            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                GoTo TSP:
            End If
         
        Case 16, 17
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & "  AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
        Case 3
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  = 1  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & "  AND intSchemeID = " & val(txtScheme.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"

        Case 10, 11, 12, 13, 14, 29

SCP:
            mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
            mSQL = mSQL + " ("
            mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND intFundCategoryID IN (2)"
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
            mSQL = mSQL + " Union ALL"
            mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (29) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
            mSQL = mSQL + " )A"
        
        Case 10, 11, 12, 13, 14, 30

TSP:
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14,30) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND intFundCategoryID = 3"
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
            
        Case Else
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
        End Select
    
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        Rec.Close
        
        
        ' 1: TOTAL ALLOTMENT RECEIVED
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 19, 21 ' Development Fund (Gen/SPC/TSP) + Special Grant + Road Renovation              10, 11, 12, 13, 14,
            If val(txtRequisitiontDate.Tag) = 1 Or val(txtRequisitiontDate.Tag) = 2 Or val(txtRequisitiontDate.Tag) = 3 Then
                mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
                mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
                mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19) AND intCategoryID=1"
                mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
                mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
                mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
                'If val(txtTreasury.Tag) <> 1 Then
                    mSQL = mSQL + " Union All"
                    mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                    mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19)  AND intCategoryID=1"
                    mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & ""
                'End If
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " SELECT fltAmount fltAmtReceived  FROM faAllotmentLetters"
                mSQL = mSQL + " WHERE ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND intSourceOfFundID = 1"
                mSQL = mSQL + " and faAllotmentLetters.intFinancialYearID = " & mYearID & "  "
                mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) =30"
                mSQL = mSQL + ")A"

            End If
        Case 10, 11, 12, 13, 14
        If val(txtRequisitiontDate.Tag) = 1 Then
                mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
                mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
                mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19) AND intCategoryID=1"
                mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
                mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
                mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
                'If val(txtTreasury.Tag) <> 1 Then
                    mSQL = mSQL + " Union All"
                    mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                    mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19)  AND intCategoryID=1"
                    mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID
                'End If
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " SELECT fltAmount fltAmtReceived  FROM faAllotmentLetters"
                mSQL = mSQL + " WHERE ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND intSourceOfFundID = 1"
                mSQL = mSQL + " and faAllotmentLetters.intFinancialYearID = " & mYearID & "  "
                mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) =30"
                mSQL = mSQL + " )A"
            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                GoTo SKIPRSCP:
            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                GoTo SKIPRTSP:
            End If
            
         Case 16, 17 'Road / Non Road
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (16,17) "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
            'If val(txtTreasury.Tag) <> 1 Then
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (16,17) "
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID
            'End If
            mSQL = mSQL + "  )A"
        Case 3 ' B-Fund
            mSQL = "Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            'mSql = mSql + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
            mSQL = mSQL + " Where IsNull(tnyCancelledFlag, 0) = 0 And intSchemeID = " & val(txtScheme.Tag)
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
        
        Case 10, 11, 12, 13, 14, 29
SKIPRSCP:
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 29) AND intCategoryID = 2"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
            'If val(txtTreasury.Tag) <> 1 Then
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 29) AND intCategoryID = 2"
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID
            'End If
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " SELECT fltAmount fltAmtReceived  FROM faAllotmentLetters"
            mSQL = mSQL + " WHERE ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND intSourceOfFundID IN (10, 11, 12, 13, 14, 29)"
            mSQL = mSQL + " and faAllotmentLetters.intFinancialYearID = " & mYearID & "  "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) =30"
            mSQL = mSQL + " )A"
        Case 10, 11, 12, 13, 14, 30
SKIPRTSP:
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 30)  AND intCategoryID = 3"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
            'If val(txtTreasury.Tag) <> 1 Then
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 30)  AND intCategoryID = 3"
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID
            'End If
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " SELECT fltAmount fltAmtReceived  FROM faAllotmentLetters"
            mSQL = mSQL + " WHERE ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND intSourceOfFundID IN (10, 11, 12, 13, 14, 30)"
            mSQL = mSQL + " and faAllotmentLetters.intFinancialYearID = " & mYearID & "  "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) =30"
            mSQL = mSQL + " )A"
            
        Case 4
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID=" & val(txtInstalmentNo.Tag) & "  "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
            'If val(txtTreasury.Tag) <> 1 Then
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID=" & val(txtInstalmentNo.Tag) & " "
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID
            'End If
            mSQL = mSQL + " )A"
        Case Else
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters "
            mSQL = mSQL + " Inner Join faIdemandTbl On faAllotmentLetters.intAllotmentID=faIdemandTbl.numSubLedgerID"
            mSQL = mSQL + " Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID=" & val(txtInstalmentNo.Tag) & "  "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND faAllotmentLetters.intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) NOT IN (30,40,90)"
            'If val(txtTreasury.Tag) <> 1 Then
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID=" & val(txtInstalmentNo.Tag) & " "
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID
            'End If
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " SELECT fltAmount fltAmtReceived  FROM faAllotmentLetters"
            mSQL = mSQL + " WHERE ISNULL(faAllotmentLetters.tnyStatus,0) = 1 AND intSourceOfFundID = 1"
            mSQL = mSQL + " and faAllotmentLetters.intFinancialYearID = " & mYearID & "  "
            mSQL = mSQL + " AND ISNULL(faAllotmentLetters.tnyGroupID,0) =30"
            mSQL = mSQL + " )A"
    End Select
 

    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!AmountReceived), "0", Rec!AmountReceived)
    End If
    Rec.Close
            
    '''''''''''''''''''''''''''Current Allotment'''''''''''''''''''''''''''''''''''''''''''''''''
    vsGrid.TextMatrix(3, 1) = val(txtAmountAuthorized.Text)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''Balance Available'''''''''''''''''''''''''''''''''''''''''''''''''
    vsGrid.TextMatrix(4, 1) = val(vsGrid.TextMatrix(1, 1)) - val(vsGrid.TextMatrix(2, 1)) - val(vsGrid.TextMatrix(3, 1))
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        ' TOTAL ALLOCATION FOR THE IMPLEMENTING OFFICER IN THE CURRENT YEAR
        
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
'''        Case 1, 21, 27, 28, 10, 11, 12, 13, 14, 19

         Case 1, 21, 27, 28, 19            '21 -Best Panchayat
            If val(txtRequisitiontDate.Tag) = 1 Or val(txtRequisitiontDate.Tag) = 2 Or val(txtRequisitiontDate.Tag) = 3 Then
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                
                
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN ( 21,27, 28, 10, 11, 12, 13, 14,19 ) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + "  AND intFundCategoryID = 1 "
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + "  Union All"
                
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN ( 1 ) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
               'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " ) A"
                  
            End If
            
        Case 10, 11, 12, 13, 14
            If val(txtRequisitiontDate.Tag) = 1 Then
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN  (21,1, 27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND intFundCategoryID = 1"
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN  (21,1, 27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
                mSQL = mSQL + " )A"
                
                
            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                GoTo GOSCP:
            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                GoTo GOTSP:
            End If
            
        Case 16, 17
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
        Case 3
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  = 1  AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & "  AND intSchemeID = " & val(txtScheme.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
        Case 10, 11, 12, 13, 14, 29
GOSCP:
            mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
            mSQL = mSQL + " ("
            mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + "  AND intFundCategoryID = 2 "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
            mSQL = mSQL + "  Union All"
            mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (29) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
            mSQL = mSQL + " )A"
        Case 10, 11, 12, 13, 14, 30
GOTSP:
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14,30) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + "  AND intFundCategoryID = 3 "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
        Case Else
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND ISNULL(intTreasuryID,0)<> 1"
    End Select
        Rec.Open mSQL, mCnn
        
        If Not (Rec.EOF And Rec.BOF) Then
             numTotalAllotmentIssuedToIMPO = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        mExpExcluding = Abs(numTotalAllotmentIssuedToIMPO - vsGrid.TextMatrix(1, 1))
        Rec.Close
        
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
                     
    Private Sub CalculateAmount_OLD_01Aug2015()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        
        numTotalAllotmentIssuedToIMPO = Null
        numTotalExpenditureExludingThisBill = Null
        
        Dim mYearID As Integer
        
        If mPreviousYearMode = 0 Then
            mYearID = gbFinancialYearID
        Else
            mYearID = gbFinancialYearID - 1
        End If
        
        '*********************************************************************************************'
        'Procedure to Calculate the Total Allotment Received, Total Amount Issued & Balance Available '
        '*********************************************************************************************'
        'On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
       
       
        ' 2 : TOTAL ALLOTMENT ISSUED TO ALL IMPLEMENTING OFFICERS
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' 4         Own Fund                                                                       '
        ' 1,27,28   Development Fund- Special Grant, Road renovation                               '
        ' 16,17-    Maintenance                                                                    '
        ' 25        CFC Grant                                                                      '
        ' 26        KLGSDP Grant                                                                   '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 19, 21              '21 -Best Panchayat
            If val(txtRequisitiontDate.Tag) = 1 Or val(txtRequisitiontDate.Tag) = 2 Or val(txtRequisitiontDate.Tag) = 3 Then
'''                mSql = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
'''                mSql = mSql + " AND intSourceID IN (1, 27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & txtRequisitiontDate.Text & "' And  tnyStage = 2 "
'''                mSql = mSql + " AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
'''                'msql = msql + " AND intFundCategoryID = 1 "
'''                'msql = msql + " AND ISNULL(intSchemeID,0) = 0"
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1"
                mSQL = mSQL + " AND intSourceID IN (21,27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID= " & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "'  And  tnyStage = 2"  'txtRequisitiontDate.Text
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND intFundCategoryID = 1"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1"
                mSQL = mSQL + " AND intSourceID IN (1) AND intFinancialYearID= " & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2"
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " )A"
            End If
            
        Case 10, 11, 12, 13, 14
            If val(txtRequisitiontDate.Tag) = 1 Then
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
                mSQL = mSQL + " AND intSourceID IN (21,27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND intFundCategoryID = 1"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
                mSQL = mSQL + " AND intSourceID IN (1) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " )A"
                
                
            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                GoTo SCP:
            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                GoTo TSP:
            End If
         
        Case 16, 17
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & "  AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
        Case 3
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  = 1  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & "  AND intSchemeID = " & val(txtScheme.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "

        Case 10, 11, 12, 13, 14, 29
SCP:
            mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
            mSQL = mSQL + " ("
            mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND intFundCategoryID IN (2)"
            mSQL = mSQL + " Union ALL"
            mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (29) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " )A"
          
        Case 10, 11, 12, 13, 14, 30
TSP:
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14,30) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " AND intFundCategoryID = 3"
            'msql = msql + " AND ISNULL(intSchemeID,0) = 0"
            
        Case Else
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
        End Select
    
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        Rec.Close
        
        ' 1: TOTAL ALLOTMENT RECEIVED

        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 19, 21 ' Development Fund (Gen/SPC/TSP) + Special Grant + Road Renovation              10, 11, 12, 13, 14,
            If val(txtRequisitiontDate.Tag) = 1 Or val(txtRequisitiontDate.Tag) = 2 Or val(txtRequisitiontDate.Tag) = 3 Then
                mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19) AND intCategoryID=1"
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
                mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
                mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                'msql = msql + " AND ISNULL(intSchemeID,0) = 0"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19)  AND intCategoryID=1"
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
'            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
'                GoTo SKIPRSCP:
'            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
'                GoTo SKIPRTSP
            End If
        Case 10, 11, 12, 13, 14
        If val(txtRequisitiontDate.Tag) = 1 Then
                mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19) AND intCategoryID=1"
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
                mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
                mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                'msql = msql + " AND ISNULL(intSchemeID,0) = 0"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
                mSQL = mSQL + " AND intSourceofFundID in (1,21,27,28, 10, 11, 12, 13, 14,19)  AND intCategoryID=1"
                mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                GoTo SKIPRSCP:
            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                GoTo SKIPRTSP:
            End If
            
         Case 16, 17 'Road / Non Road
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (16,17) "
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (16,17) "
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
        Case 3 ' B-Fund
            mSQL = "Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0 AND intSchemeID = " & val(txtScheme.Tag)
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
        
        Case 10, 11, 12, 13, 14, 29
SKIPRSCP:
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 29) AND intCategoryID = 2"
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 29) AND intCategoryID = 2"
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
             
        Case 10, 11, 12, 13, 14, 30
SKIPRTSP:
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 30)  AND intCategoryID = 3"
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID in (10, 11, 12, 13, 14, 30)  AND intCategoryID = 3"
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
        Case Else
            mSQL = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID=" & val(txtInstalmentNo.Tag) & "  "
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & "  "
            mSQL = mSQL + " AND dtAllotmentDate <= '" & DdMmmYy(gbTransactionDate) & "'"
            mSQL = mSQL + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSQL = mSQL + " AND intSourceofFundID=" & val(txtInstalmentNo.Tag) & " "
            mSQL = mSQL + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
    End Select
 

        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!AmountReceived), "0", Rec!AmountReceived)
        End If
        Rec.Close
                
        '''''''''''''''''''''''''''Current Allotment'''''''''''''''''''''''''''''''''''''''''''''''''
        vsGrid.TextMatrix(3, 1) = val(txtAmountAuthorized.Text)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''Balance Available'''''''''''''''''''''''''''''''''''''''''''''''''
        vsGrid.TextMatrix(4, 1) = val(vsGrid.TextMatrix(1, 1)) - val(vsGrid.TextMatrix(2, 1)) - val(vsGrid.TextMatrix(3, 1))
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ' TOTAL ALLOCATION FOR THE IMPLEMENTING OFFICER IN THE CURRENT YEAR
        
        Select Case val(txtInstalmentNo.Tag) 'SOURCE OF FUND
'''        Case 1, 21, 27, 28, 10, 11, 12, 13, 14, 19

         Case 1, 21, 27, 28, 19            '21 -Best Panchayat
            If val(txtRequisitiontDate.Tag) = 1 Or val(txtRequisitiontDate.Tag) = 2 Or val(txtRequisitiontDate.Tag) = 3 Then
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                
                
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN ( 21,27, 28, 10, 11, 12, 13, 14,19 ) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + "  AND intFundCategoryID = 1 "
                mSQL = mSQL + "  Union All"
                
                mSQL = mSQL + "  Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN ( 1 ) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
               'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " ) A"
                  
            End If
            
        Case 10, 11, 12, 13, 14
            If val(txtRequisitiontDate.Tag) = 1 Then
                mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                mSQL = mSQL + " ("
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN  (21,1, 27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " AND intFundCategoryID = 1"
                mSQL = mSQL + " Union All"
                mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
                mSQL = mSQL + " AND intSourceID IN  (21,1, 27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " "
                mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
                'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
                mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                mSQL = mSQL + " )A"
                
                
            ElseIf val(txtRequisitiontDate.Tag) = 2 Then
                GoTo GOSCP:
            ElseIf val(txtRequisitiontDate.Tag) = 3 Then
                GoTo GOTSP:
            End If
            
        Case 16, 17
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
        Case 3
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  = 1  AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & "  AND intSchemeID = " & val(txtScheme.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
        Case 10, 11, 12, 13, 14, 29
GOSCP:
            mSQL = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
            mSQL = mSQL + " ("
            mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + "  AND intFundCategoryID = 2 "
            mSQL = mSQL + "  Union All"
            mSQL = mSQL + " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (29) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + " )A"
        Case 10, 11, 12, 13, 14, 30
GOTSP:
            mSQL = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & " "
            mSQL = mSQL + " AND intSourceID IN (10, 11, 12, 13, 14,30) AND intFinancialYearID=" & mYearID & " "
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            mSQL = mSQL + "  AND intFundCategoryID = 3 "
        Case Else
            mSQL = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & val(txtIMPOName.Tag) & "  "
            mSQL = mSQL + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSQL = mSQL + " AND dtAuthorizationDate <= '" & DdMmmYy(gbTransactionDate) & "' And  tnyStage = 2 "
            'mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2) "
            mSQL = mSQL + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
    End Select
        Rec.Open mSQL, mCnn
        
        If Not (Rec.EOF And Rec.BOF) Then
             numTotalAllotmentIssuedToIMPO = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        mExpExcluding = Abs(numTotalAllotmentIssuedToIMPO - vsGrid.TextMatrix(1, 1))
        Rec.Close
        
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    
    Private Sub FetchRequisitionDetails()
        Dim mRequisitionID  As Variant
        Dim objDB           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mCnn            As New ADODB.Connection
        
        '*********************************************************************************************'
        '                       Procedure to fetch Requisition Details                                '
        '*********************************************************************************************'
        'On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If CheckPreviousYearRequisitions(RequisitionID) = 1 Then
            mRequisitionID = Array(RequisitionID, gbFinancialYearID - 1)
            mPreviousYearMode = 1
        Else
            mRequisitionID = Array(RequisitionID, gbFinancialYearID)
            mPreviousYearMode = 0
        End If
        'mRequisitionID = Array(RequisitionID, gbFinancialYearID)
        
        Set Rec = objDB.ExecuteSP("spRptViewAllotmentLetter", mRequisitionID, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
            txtRequisition.Text = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
            txtRequisition.Tag = IIf(IsNull(Rec!intID), "", Rec!intID)
            txtRequisitiontDate.Text = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
            txtRequisitiontDate.Tag = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
            txtInstalmentNo.Text = IIf(IsNull(Rec!tnyInstallmentNo), "", Rec!tnyInstallmentNo)
            txtInstalmentNo.Tag = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
            txtLSGICode.Text = IIf(IsNull(Rec!chvLocalBodyCode), "", Rec!chvLocalBodyCode) 'Rec!chrLocalBodyCode)
            txtLSGIName.Text = IIf(IsNull(Rec!vchLocalBody), "", Rec!vchLocalBody)
            txtLSGIName.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
            txtIMPOName.Text = IIf(IsNull(Rec!vchNameofIMPO), "", Rec!vchNameofIMPO)
            txtIMPOName.Tag = IIf(IsNull(Rec!intImplementingOfficersID), "", Rec!intImplementingOfficersID)
            txtAmountAuthorized.Text = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
            txtAmountInWords.Text = Words(val(txtAmountAuthorized.Text))
            
            txtTreasury.Text = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
            txtTreasury.Tag = IIf(IsNull(Rec!intTreasuryID), "", Rec!intTreasuryID)
            
            txtTreasuryCode.Text = IIf(IsNull(Rec!vchTreasuryCode), "", Rec!vchTreasuryCode)
            MaskAccHead.Text = IIf(IsNull(Rec!vchGHeadofAccount), "", Rec!vchGHeadofAccount)
            
            MaskDetailAccHead.MaxLength = 500
            MaskDetailAccHead.Text = IIf(IsNull(Rec!vchGBudgetHead), "", Rec!vchGBudgetHead)
            txtScheme.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            txtScheme.Tag = IIf(IsNull(Rec!intSchemeID), "", Rec!intSchemeID)
            txtAuthorisedDate.Text = DdMmmYy(IIf(IsNull(Rec!dtAuthorizationDate), "", Rec!dtAuthorizationDate))
            txtAuthorisedDate.Tag = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
            txtNewMode.Text = IIf(IsNull(Rec!intTreasuryID), "", Rec!intTreasuryID)
            
        End If
        Rec.Close
        
        'RequisitionID = ""
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub
    
    Private Sub cmdSave_Click()
    
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim mSQL        As String
        Dim mArrIn      As Variant
        Dim mYearID     As Variant
        
        '*********************************************************************************************'
        '                       Procedure to Save the Allotment Details                               '
        '*********************************************************************************************'
        On Error GoTo err
        Dim mCnnSulekha As New ADODB.Connection
        If Not (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
            MsgBox "Connection To Plan [Sulekha] Module not found", vbCritical
            Exit Sub
        End If
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        If CheckPreviousYearRequisitions(RequisitionID) = 1 Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        'mExpExcluding = numTotalAllotmentIssuedToIMPO
        If mExpExcluding < 0 Then mExpExcluding = 0
        mArrIn = Array(RequisitionID, _
                        val(vsGrid.TextMatrix(1, 1)), _
                        val(vsGrid.TextMatrix(2, 1)), _
                        val(vsGrid.TextMatrix(3, 1)), _
                        val(vsGrid.TextMatrix(4, 1)), _
                        numTotalAllotmentIssuedToIMPO, _
                        mExpExcluding, _
                        0, gbTransactionDate)
        
        objDB.ExecuteSP "spSaveIssueLetterOfAllotment", mArrIn, , , mCnn, adCmdStoredProc
        cmdSave.Enabled = False
        mSQL = "Update faPendingTaskRequest set  tnyStatus=8 Where intKeyID=" & RequisitionID & "  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        Call UpdateSulekhaReqDetials(RequisitionID)
        frmViewAllotmentLetter.Mode = 3
        frmViewAllotmentLetter.ArrayIn = Array(CStr(val(txtRequisition.Tag)), CStr(mYearID))
        Unload Me
        frmViewAllotmentLetter.Show vbModal
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Function UpdateSulekhaReqDetials(mReqID As Integer)
        Dim mCnnSulekha     As New ADODB.Connection
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim RecSulekha      As New ADODB.Recordset
        Dim mArrIn          As Variant
        Dim mSQL            As String
        Dim msqlSulekha     As String
        Dim mAllotmentNo    As Long
        Dim dtAllotmentDate As Date
        Dim mSourceID As Integer
        Dim mStatus As Integer
        Dim mProjectID As Variant
        Dim mYearID As Integer
        
        If mPreviousYearMode = 1 Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "Select * from faAllotments Where intID=" & mReqID
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mAllotmentNo = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                dtAllotmentDate = DdMmmYy(IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate))
                mSourceID = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
                If mSourceID = 41 Then  ''' Added on 28 Dec 2016 KLGSDP Fund
                    mSourceID = 26
                End If
                mStatus = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
                mProjectID = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
            End If
        End If
        
        If mProjectID <> 0 Then
            If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                mArrIn = Array(mReqID, _
                                mAllotmentNo, _
                                dtAllotmentDate, _
                                val(vsGrid.TextMatrix(3, 1)), _
                                mSourceID, _
                                mProjectID, _
                                gbLBID, _
                                mYearID, _
                                mStatus, _
                                0)
                                
                objDB.ExecuteSP "RequisitionDetails_I", mArrIn, , , mCnnSulekha, adCmdStoredProc
            Else
                MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
                Exit Function
            End If
        End If
    End Function
    Private Sub Form_Load()
        Call FormInitialize
        vsGrid.HighLight = flexHighlightNever
        vsGrid.MergeCells = flexMergeFree
        vsGrid.MergeRow(0) = True
        vsGrid.Cell(flexcpFontBold, 0, 0, , 1) = True
        Call FetchRequisitionDetails
        If txtNewMode.Text = 1 Then
            Call CalculateAmountNewMode
        Else
            Call CalculateAmount
        End If
    End Sub
    Public Function CheckPreviousYearRequisitions(mReqID As Integer)
        Dim mSQL        As String
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            If mLoadModeUnAuth = 10 Then
                mSQL = "Select * From faPendingTaskRequest Where intTaskID = 16 And intKeyID=" & mReqID
            'ElseIf mSourceID = 3 And intTaskID = 13 Then 'Modified by Aiby on  14-Jul-2014
            ElseIf mSourceID = 3 Then
                mSQL = "Select * From faPendingTaskRequest Where intTaskID IN (3,13) And intKeyID=" & mReqID
            Else
                mSQL = "Select * From faPendingTaskRequest Where intTaskID = 3 And intKeyID=" & mReqID
            End If
            Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                CheckPreviousYearRequisitions = 1
            Else
                CheckPreviousYearRequisitions = 0
            End If
            Rec.Close
        End If
    End Function
    Public Property Let RequisitionID(mData As Variant)
        ReqID = mData
    End Property
    
    Public Property Get RequisitionID() As Variant
        RequisitionID = ReqID
    End Property

    Private Sub Form_Unload(Cancel As Integer)
        If mLoadModeUnAuth = 10 Then
            frmListOfRequisitions.LoadMode = 10
            frmListOfRequisitions.Visible = True
            frmListOfRequisitions.ZOrder (0)
            mLoadModeUnAuth = 0
        Else
            frmListOfRequisitions.Visible = True
            frmListOfRequisitions.ZOrder (0)
        End If
    End Sub
    Public Property Let LoadMode(mData As Integer)
        mLoadModeUnAuth = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = mLoadModeUnAuth
    End Property

    Public Property Let SourceID(mData As Integer)
        mSourceID = mData
    End Property
    
    Public Property Get SourceID() As Integer
        SourceID = mSourceID
    End Property

