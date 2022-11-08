VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransactionsMapping 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mapping of Transactions with Function, Functionary & Account Head"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransactionsMapping.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcbTransaction 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   15
      ScaleHeight     =   3180
      ScaleWidth      =   15210
      TabIndex        =   18
      Top             =   6150
      Width           =   15210
      Begin VB.Frame fraTransactionType 
         Appearance      =   0  'Flat
         Caption         =   "Transaction Details"
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   11280
         TabIndex        =   19
         Top             =   -45
         Width           =   3930
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   870
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   2430
            Width           =   1245
         End
         Begin VB.CommandButton cmdClose1 
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2130
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   2430
            Width           =   1245
         End
         Begin VB.TextBox txtDate 
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
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   675
            Width           =   1695
         End
         Begin VB.TextBox txtReceiptNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   375
            Width           =   1680
         End
         Begin VB.TextBox txtTransactionTypeR 
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
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1050
            Width           =   1905
         End
         Begin VB.CommandButton cmdTransactionTypeR 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3570
            TabIndex        =   25
            Top             =   1065
            Width           =   285
         End
         Begin VB.TextBox txtFunctionaryR 
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
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1725
            Width           =   1905
         End
         Begin VB.CommandButton cmdFunctionR 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3570
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2040
            Width           =   285
         End
         Begin VB.CommandButton cmdFunctionaryR 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3570
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1725
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtFunctionR 
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
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   2025
            Width           =   1905
         End
         Begin VB.TextBox txtAmountR 
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
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1350
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Type"
            Height          =   270
            Left            =   90
            TabIndex        =   34
            Top             =   1035
            Width           =   1515
         End
         Begin VB.Label lblReceiptNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt No"
            Height          =   270
            Left            =   660
            TabIndex        =   33
            Top             =   375
            Width           =   945
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   270
            Left            =   1200
            TabIndex        =   32
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Function"
            Height          =   270
            Left            =   855
            TabIndex        =   31
            Top             =   2040
            Width           =   750
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Functionary"
            Height          =   270
            Left            =   585
            TabIndex        =   30
            Top             =   1710
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   270
            Left            =   930
            TabIndex        =   29
            Top             =   1335
            Width           =   675
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsReceipt 
         Height          =   2760
         Left            =   15
         TabIndex        =   35
         Top             =   15
         Width           =   11265
         _cx             =   19870
         _cy             =   4868
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
         BackColor       =   16318457
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16318457
         BackColorAlternate=   16318457
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransactionsMapping.frx":1CCA
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
         TextStyleFixed  =   1
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
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Search Criteria"
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   11295
      TabIndex        =   9
      Top             =   -15
      Width           =   3930
      Begin VB.TextBox txtAmount1 
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
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   3135
         Width           =   2460
      End
      Begin VB.TextBox txtFunction 
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1730
         Width           =   3150
      End
      Begin VB.TextBox txtFunctionary 
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1020
         Width           =   3150
      End
      Begin VB.TextBox txtTransactionType 
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2440
         Width           =   3150
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4545
         Width           =   1245
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4545
         Width           =   1245
      End
      Begin VB.CommandButton cmdFunction 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdFunctionary 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1020
         Width           =   285
      End
      Begin VB.CheckBox chkListToApprove 
         Caption         =   "List of Corrected Transactions"
         Height          =   270
         Left            =   90
         TabIndex        =   7
         Top             =   5550
         Width           =   2970
      End
      Begin VB.TextBox txtDateFrom 
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
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   3855
         Width           =   1260
      End
      Begin VB.TextBox txtDateTo 
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
         Height          =   285
         Left            =   2145
         TabIndex        =   5
         Top             =   3855
         Width           =   1260
      End
      Begin VB.CommandButton cmdTransactionType 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3300
         TabIndex        =   2
         Top             =   2460
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   3405
         TabIndex        =   6
         Top             =   3855
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63700993
         CurrentDate     =   40197
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   3855
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63700993
         CurrentDate     =   40197
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         Height          =   270
         Left            =   135
         TabIndex        =   15
         Top             =   1455
         Width           =   750
      End
      Begin VB.Label lblFunctionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
         Height          =   270
         Left            =   135
         TabIndex        =   14
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label lblFromDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   270
         Left            =   135
         TabIndex        =   13
         Top             =   3600
         Width           =   915
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   270
         Left            =   2115
         TabIndex        =   12
         Top             =   3600
         Width           =   690
      End
      Begin VB.Label lblTransactionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         Height          =   270
         Left            =   135
         TabIndex        =   11
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   270
         Left            =   135
         TabIndex        =   10
         Top             =   2895
         Width           =   675
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6030
      Left            =   15
      TabIndex        =   8
      Top             =   15
      Width           =   11265
      _cx             =   19870
      _cy             =   10636
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransactionsMapping.frx":1DE3
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
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   -30
      Top             =   6105
      Width           =   15255
   End
End
Attribute VB_Name = "frmTransactionsMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    '*********************************************************************************************'
    '       Form to correct the Transactions with wrong Function,Functionary & Trns.Type          '
    '*********************************************************************************************'
    Private Sub ClearTransactionDetails()
        txtReceiptNo.Text = ""
        txtReceiptNo.Tag = ""
        txtDate.Text = ""
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        txtAmountR.Text = ""
        txtFunctionary.Text = ""
        txtFunctionary.Tag = ""
        txtFunction.Text = ""
        txtFunction.Tag = ""
        vsReceipt.Clear 1, 1
    End Sub
    
    Public Sub Calculate()
        Dim mAmtArrear  As Double
        Dim mAmtCurrent As Double
        Dim mRoundOff   As Double
        Dim mTotal      As Double
        Dim mCount As Long
        
        On Error GoTo err
        For mCount = 1 To vsReceipt.Rows - 1
            If val(vsReceipt.TextMatrix(mCount, 4)) Then
                mAmtArrear = mAmtArrear + val(vsReceipt.Cell(flexcpText, mCount, 4))
            Else
                mAmtCurrent = mAmtCurrent + val(vsReceipt.Cell(flexcpText, mCount, 5))
            End If
        Next
        mTotal = Format(mAmtArrear + mAmtCurrent, "0.00")
        mRoundOff = Format(RoundOffAdjustment(mTotal), "0.00")
        txtAmountR.Text = Format(val(mTotal) + val(mRoundOff), "0.00")
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Function FillGrid()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSQL        As String
        Dim mRowCount   As Double
        
        On Error GoTo err
        
        vsGrid.Rows = 1
        mRowCount = 1
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "Select intTransactionID,dtDate,faTransactionChildForBudget.intFunctionID[FunctionID],vchFunction,faFunctionaryFunctions.intFunctionaryID[FunctionaryID],vchFunctionary,faTransactionChildForBudget.intTransactionTypeID[TransactionTypeID],vchTransactionType,Sum(fltAmount)As Amount,tnyFlag,intYearID From faTransactionChildForBudget "
            mSQL = mSQL + " Left Join faFunctionaryFunctions on faTransactionChildForBudget.intFunctionID =faFunctionaryFunctions.intFunctionID"
            mSQL = mSQL + " Left Join faFunctions On faTransactionChildForBudget.intFunctionID = faFunctions.intFunctionID"
            mSQL = mSQL + " Left Join faFunctionaries On faFunctionaryFunctions.intFunctionaryID = faFunctionaries.intFunctionaryID"
            mSQL = mSQL + " Left Join faTransactionType On faTransactionChildForBudget.intTransactionTypeID = faTransactionType.intTransactionTypeID"
            mSQL = mSQL + " Where tnyVoucherGroupID = 10"
            mSQL = mSQL + " And intSerialNo <> 1"
            mSQL = mSQL + " AND NOT faTransactionChildForBudget.intFunctionID is Null"
            'mSQL = mSQL + " AND  faFunctionaryFunctions.intFunctionID is Null"
            If txtFunctionary.Tag <> "" Then
                mSQL = mSQL + " AND faFunctionaryFunctions.intFunctionaryID = " & txtFunctionary.Tag
            End If
            If txtFunction.Tag <> "" Then
                mSQL = mSQL + " AND faTransactionChildForBudget.intFunctionID = " & txtFunction.Tag
            End If
            If txtTransactionType.Tag <> "" Then
                mSQL = mSQL + " AND faTransactionChildForBudget.intTransactionTypeID = " & txtTransactionType.Tag
            End If
            If txtAmount1.Text <> "" Then
                mSQL = mSQL + " AND fltAmount = " & val(txtAmount1.Text)
            End If
            mSQL = mSQL + " Group By intTransactionID,dtDate,faTransactionChildForBudget.intFunctionID,vchFunction,faFunctionaryFunctions.intFunctionaryID,vchFunctionary,faTransactionChildForBudget.intTransactionTypeID,vchTransactionType,tnyFlag,intYearID"
            mSQL = mSQL + " Order By intTransactionID"
            Rec.Open mSQL, mCnn
            While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCount, 0) = mRowCount
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
                vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
                vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!FunctionID), "", Rec!FunctionID)
                vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!FunctionaryID), "", Rec!FunctionaryID)
                vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!TransactionTypeID), "", Rec!TransactionTypeID)
                vsGrid.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!tnyFlag), "", Rec!tnyFlag)
                Rec.MoveNext
                mRowCount = mRowCount + 1
            Wend
            Rec.Close
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Function
err:
        MsgBox err.Description
    End Function

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdClose1_Click()
        Me.Height = 6645
        ClearTransactionDetails
    End Sub

    Private Sub cmdFunction_Click()
        On Error GoTo err:
        frmSearchFunction.Show vbModal
        If Not gbSearchStr = "" Then
            txtFunction.Text = mID(gbSearchStr, 10)
            txtFunction.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdFunctionary_Click()
        On Error GoTo err:
            frmSearchFunctionary.Show vbModal
            If Not gbSearchStr = "" Then
                txtFunctionary.Text = mID(gbSearchStr, 9)
                txtFunctionary.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdFunctionaryR_Click()
        On Error GoTo err:
        frmSearchFunctionary.Show vbModal
        If Not gbSearchStr = "" Then
            txtFunctionaryR.Text = mID(gbSearchStr, 9)
            txtFunctionaryR.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub cmdFunctionR_Click()
        On Error GoTo err:
        frmSearchFunction.Show vbModal
        If Not gbSearchStr = "" Then
            txtFunctionR.Text = mID(gbSearchStr, 10)
            txtFunctionR.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSearch_Click()
        FillGrid
    End Sub

    Private Sub cmdTransactionType_Click()
        On Error GoTo err:
            frmSearchTransactionType.Show vbModal
            If Not gbSearchStr = "" Then
                txtTransactionType.Text = gbSearchStr
                txtTransactionType.Tag = gbSearchID
                txtTransactionType.SetFocus
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdTransactionTypeR_Click()
        On Error GoTo err:
        frmSearchTransactionType.Show vbModal
        If Not gbSearchStr = "" Then
            txtTransactionTypeR.Text = gbSearchStr
            txtTransactionTypeR.Tag = gbSearchID
            txtTransactionTypeR.SetFocus
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdUpdate_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim mSQL        As String
        Dim mArrIn      As Variant
        Dim mLooop      As Integer
        
        '*********************************************************************************************'
        '       Procedure to update the Transaction with right Function, Functionary & Trns.Type     '
        '*********************************************************************************************'
        On Error GoTo err
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mArrIn = Array(val(txtDate.Tag), _
                        val(txtFunctionR.Tag), _
                        val(txtTransactionTypeR.Tag), _
                        1 _
                        )
            objDB.ExecuteSP "spUpdateTransactionChildForBudget", mArrIn, , , mCnn, adCmdStoredProc
            For mLooop = 1 To vsReceipt.Rows - 1
                mSQL = "Update faTransactionChildForBudget"
                mSQL = mSQL + " Set intAccountHeadID = " & vsReceipt.TextMatrix(mLooop, 7)
                mSQL = mSQL + " Where intTransactionID = " & txtDate.Tag
                mSQL = mSQL + " And intAccountHeadID = " & vsReceipt.TextMatrix(mLooop, 6)
                mCnn.Execute mSQL
            Next
            MsgBox "Successfully Updated", vbInformation
        End If
        Call FillGrid
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Me.Width = 15360
        Me.Height = 6645
    End Sub
    
    Private Sub Form_Load()
       ' txtDateFrom.Text = Date
       ' txtDateTo.Text = Date
        FillGrid
        vsReceipt.ColComboList(0) = "|..."
    End Sub

    Private Sub pcbTransaction_LostFocus()
'        Me.Height = 6645
    End Sub

    Private Sub txtAmount1_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDateFrom_LostFocus()
        If txtDateFrom.Text <> "" Then
            txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
        End If
    End Sub
    
    Private Sub txtDateTo_LostFocus()
         If txtDateTo.Text <> "" Then
            txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
        End If
    End Sub

    Private Sub txtFunction_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then 'Delete Key
            txtFunction.Text = ""
            txtFunction.Tag = ""
        Else
            txtFunction.Locked = True
        End If

    End Sub

    Private Sub txtFunctionary_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = 46 Then 'Delete Key
            txtFunctionary.Text = ""
            txtFunctionary.Tag = ""
        Else
            txtFunctionary.Locked = True
        End If
    End Sub

    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = 46 Then 'Delete Key
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        Else
            txtTransactionType.Locked = True
        End If
    End Sub

    Private Sub VSGrid_DblClick()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim mSQL        As String
        Dim Rec         As New ADODB.Recordset
        Dim mRowCount   As Integer
        Dim mPeriodID   As Variant
        Dim mYearID     As Variant
        Dim mArrearFlag As Variant
        
        On Error GoTo err
        If vsGrid.Rows > 1 Then
            If vsGrid.TextMatrix(vsGrid.Row, 6) <> "" Then
                If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                    mSQL = "Select faVouchers.intVoucherNo[VoucherNo],faVouchers.intVoucherID[VoucherID],faTransactionChildForBudget.dtDate,vchTransactionType,faTransactionChildForBudget.intTransactionTypeID[TransactionTypeID],vchFunctionary,faTransactions.intFunctionaryID[FunctionaryID],vchFunction,faTransactionChildForBudget.intFunctionID[FunctionID]"
                    mSQL = mSQL + " From faTransactionChildForBudget"
                    mSQL = mSQL + " Left Join faTransactions On faTransactions.intTransactionID = faTransactionChildForBudget.intTransactionID"
                    mSQL = mSQL + " Left Join faVouchers On faTransactions.intVoucherID = faVouchers.intVoucherID"
                    mSQL = mSQL + " Left Join faTransactionType On faTransactions.intTransactionTypeID = faTransactionType.intTransactionTypeID"
                    mSQL = mSQL + " Left Join faFunctions On faTransactionChildForBudget.intFunctionID = faFunctions.intFunctionID"
                    mSQL = mSQL + " Left Join faFunctionaries On faTransactions.intFunctionaryID = faFunctionaries.intFunctionaryID"
                    mSQL = mSQL + " Where faTransactions.intTransactionID = " & vsGrid.TextMatrix(vsGrid.Row, 6)
                    mSQL = mSQL + " And faTransactionChildForBudget.intFunctionID Is not null"
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        Me.Height = 9450
                        txtReceiptNo.Text = IIf(IsNull(Rec!VoucherNo), "", Rec!VoucherNo)
                        txtReceiptNo.Tag = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                        txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                        txtDate.Tag = val(vsGrid.TextMatrix(vsGrid.Row, 6))
                        txtTransactionTypeR.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                        txtTransactionTypeR.Tag = IIf(IsNull(Rec!TransactionTypeID), "", Rec!TransactionTypeID)
                        txtFunctionaryR.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                        txtFunctionaryR.Tag = IIf(IsNull(Rec!FunctionaryID), "", Rec!FunctionaryID)
                        txtFunctionR.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                        txtFunctionR.Tag = IIf(IsNull(Rec!FunctionID), "", Rec!FunctionID)
                    End If
                    Rec.Close
                    
                    If txtReceiptNo.Tag <> "" Then
'                        mSQL = "Select * From faVoucherChild"
'                        mSQL = mSQL + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
'                        '-----------------------------------------
'                        'Added By Anisha On 30.09.10 to Diplay Period
'                        mSQL = mSQL + " left Join faPeriodicity On faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
'                        '-------------------------------------------
'                        mSQL = mSQL + " Where intVoucherID=" & txtReceiptNo.Tag
                        mSQL = "Select faTransactionChildForBudget.intAccountHeadID[AccountHeadID],vchAccountHeadCode,vchAccountHead,faVoucherChild.intYearID[YearID],tnyPeriodID,faTransactionChildForBudget.fltAmount[Amount],tnyArrearFlag,vchPeriodicity From faTransactionChildForBudget"
                        mSQL = mSQL + " Inner Join faAccountHeads On faTransactionChildForBudget.intAccountHeadID = faAccountHeads.intAccountHeadID"
                        mSQL = mSQL + " Inner Join faTransactions On faTransactionChildForBudget.intTransactionID = faTransactions.intTransactionID"
                        mSQL = mSQL + " Inner Join faVoucherChild On faTransactions.intVoucherID = faVoucherChild.intVoucherID And faVoucherChild.intAccountHeadID = faTransactionCHildForBudget.intAccountHeadID"
                        mSQL = mSQL + " Left Join faPeriodicity On faVoucherChild.tnyPeriodID = faPeriodicity.intPeriodicityID"
                        mSQL = mSQL + " Where faTransactionChildForBudget.intTransactionID = " & txtDate.Tag
                        mSQL = mSQL + " And intSerialNo <> 1"
                        mSQL = mSQL + " Group By faTransactionChildForBudget.intAccountHeadID,vchAccountHeadCode,vchAccountHead,faVoucherChild.intYearID,tnyPeriodID,faTransactionChildForBudget.fltAmount,tnyArrearFlag,vchPeriodicity"
                        mSQL = mSQL + " Order By tnyArrearFlag Desc"

                        Rec.Open mSQL, mCnn
                        mRowCount = 1
                        vsReceipt.Rows = 1
                        While Not Rec.EOF
                            vsReceipt.Rows = vsReceipt.Rows + 1
                            vsReceipt.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                            vsReceipt.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                            
                            ''''''''''''''''''''''''To be Removed'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            If txtTransactionTypeR.Tag = 12 And Rec!vchAccountHeadCode = 140130400 Then
                                vsReceipt.TextMatrix(mRowCount, 0) = "140130200"
                                vsReceipt.TextMatrix(mRowCount, 1) = "Fees for Delayed Registration - Birth & DeathCertificate"
                            End If
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            
                            'mPeriodID = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                            mYearID = IIf(IsNull(Rec!YearID), 0, Rec!YearID)
                            If mYearID <> 0 Then
                                vsReceipt.TextMatrix(mRowCount, 2) = mYearID & "-" & mYearID + 1
                            End If
                            
                            '-----------------------------------------
                            'Added By Anisha On 30.09.10 to Diplay Period
    '                        If mPeriodID = 1 Then
    '                            vsreceipt.TextMatrix(mRowCount, 3) = "1st Half"
    '                        End If
    '                        If mPeriodID = 2 Then
    '                            vsreceipt.TextMatrix(mRowCount, 3) = "2nd Half"
    '                        End If
    '                        If mPeriodID = 3 Then
    '                            vsreceipt.TextMatrix(mRowCount, 3) = "Full Year"
    '                        End If
                            
                            vsReceipt.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity)
                            '--------------------------------------------------------
                            mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), "", Rec!tnyArrearFlag)
                            If mArrearFlag = 0 Then
                                vsReceipt.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
                            End If
                            If mArrearFlag = 1 Then
                                vsReceipt.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
                            End If
                            vsReceipt.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!AccountHeadID), "", Rec!AccountHeadID)
                            'vsReceipt.Rows = vsReceipt.Rows + 1
                            mRowCount = mRowCount + 1
                            Rec.MoveNext
                        Wend
                        Rec.Close
                        Call Calculate
                    End If
                Else
                    MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
                End If
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub vsReceipt_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsReceipt.Row > 1 Then
                If vsReceipt.TextMatrix(vsReceipt.Row - 1, 0) = "" Or _
                   (val(vsReceipt.TextMatrix(vsReceipt.Row - 1, 4)) <= 0 And _
                   val(vsReceipt.TextMatrix(vsReceipt.Row - 1, 5)) <= 0) Then
                   Cancel = True
                   Exit Sub
                End If
            End If
            
            If Col = 4 Or Col = 5 Then
                If Trim(vsReceipt.TextMatrix(Row, 0)) = "" Then
                    Cancel = True
                End If
            End If
            
            If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    vsReceipt.TextMatrix(Row, 0) = objAccHead.AccountCode
                    vsReceipt.TextMatrix(Row, 1) = objAccHead.AccountHead
                    vsReceipt.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
                vsReceipt.Col = vsReceipt.Col + 2
                vsReceipt.Redraw = flexRDDirect
                gbSearchStr = ""
            ElseIf Col = 1 Then
                Cancel = True
            End If
    End Sub

    Private Sub vsReceipt_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If vsReceipt.Row <= 9 Then
            Dim mSQL As String
            If val(txtTransactionTypeR.Tag) > 0 And val(txtTransactionTypeR.Tag) < 9999 Then
                mSQL = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
                mSQL = mSQL + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                mSQL = mSQL + " Where intTransactionTypeID = " & val(txtTransactionTypeR.Tag) & " And faAccountHeads.tinHiddenFlag = 0 And faAccountHeads.intGroupID is Null Order By faTransactionTypeChild.intOrder"
                frmSearchAccountHeads.SQLString = mSQL '"Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            Else
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By faAccountHeads.vchAccountHeadCode"
            End If
            frmSearchAccountHeads.VoucherMode = 100
            frmSearchAccountHeads.Show vbModal
        Else
            MsgBox "Can't print more than 9 rows in this Demand", vbInformation
        End If
    End Sub

    Private Sub vsReceipt_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        'If vsGrid.Row > 0 Then
        If Row > 0 Then
            If Col = 0 Then
                Set objAccHead = New clsAccounts
                objAccHead.SetAccountCode (Trim(vsReceipt.TextMatrix(Row, 0)))
                If objAccHead.AccountHeadID > 0 Then
                    vsReceipt.TextMatrix(Row, 0) = objAccHead.AccountCode
                    vsReceipt.TextMatrix(Row, 1) = objAccHead.AccountHead
                    vsReceipt.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                Else
                    '------------------------------------------'''''Added and edited By Sinoj'''''''
                    If vsReceipt.TextMatrix(Row, 1) <> "" Then
                        vsReceipt.RemoveItem (Row)
                    End If
'                    vsGrid.TextMatrix(Row, 0) = ""
'                    vsGrid.TextMatrix(Row, 1) = ""
'                    vsGrid.TextMatrix(Row, 6) = ""
'                    vsGrid.TextMatrix(Row, 4) = ""
'                    vsGrid.TextMatrix(Row, 5) = ""
                    Call Calculate
                    '------------------------------------------'''''Added and editted By Sinoj'''''''
                End If
            ElseIf Col = 1 And vsReceipt.ComboIndex > -1 Then
                Set objAccHead = New clsAccounts
                If objAccHead.FindAccountByHead(Trim(vsReceipt.ComboItem)) Then
                vsReceipt.TextMatrix(Row, 0) = objAccHead.AccountCode
                vsReceipt.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
'            ElseIf Col = 4 Then
'                If mRoundOffDecimalPlace Then
'                    vsReceipt.TextMatrix(Row, 4) = Format(val(vsReceipt.TextMatrix(Row, 4)), "#0")
'                Else
'                    vsReceipt.TextMatrix(Row, 4) = Format(val(vsReceipt.TextMatrix(Row, 4)), "0.00")
'                End If
'                If val(vsReceipt.TextMatrix(Row, 4)) > 0 Then
'                vsReceipt.TextMatrix(Row, 5) = ""
'                End If
'                Call Calculate
'            ElseIf Col = 5 Then
'                If mRoundOffDecimalPlace Then
'                    vsReceipt.TextMatrix(Row, 5) = Format(val(vsReceipt.TextMatrix(Row, 5)), "#0")
'                Else
'                    vsReceipt.TextMatrix(Row, 5) = Format(val(vsReceipt.TextMatrix(Row, 5)), "0.00")
'                End If
'                If val(vsReceipt.TextMatrix(Row, 5)) > 0 Then
'                vsReceipt.TextMatrix(Row, 4) = ""
'                End If
                
            End If
            'Call Calculate
            'Call ValuesForHiddenColumns
        End If
    End Sub

    Private Sub vsReceipt_KeyPress(KeyAscii As Integer)
        If vsReceipt.Col = 0 Then
            vsReceipt.Editable = flexEDKbdMouse
        Else
            'vsReceipt.Editable = flexEDNone
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsReceipt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsReceipt.Col = 0 Then
            vsReceipt.Editable = flexEDKbdMouse
        Else
            'vsReceipt.Editable = flexEDNone
            KeyAscii = 0
        End If
    End Sub
