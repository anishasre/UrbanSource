VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPublish 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Publish Monthly Transactions"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid vsMonthly 
      Height          =   3765
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6555
      _cx             =   11562
      _cy             =   6641
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      BackColorAlternate=   12640511
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
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPublish.frx":0000
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
      Begin VB.CommandButton Command24 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Top             =   270
         Width           =   1695
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   23
         Top             =   555
         Width           =   1695
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   22
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Top             =   1125
         Width           =   1695
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Top             =   1410
         Width           =   1695
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Top             =   1695
         Width           =   1695
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   1980
         Width           =   1695
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   17
         Top             =   2265
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   16
         Top             =   2550
         Width           =   1695
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Top             =   2835
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Publish"
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Top             =   3405
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   12
         Top             =   3405
         Width           =   1650
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   11
         Top             =   3120
         Width           =   1650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   10
         Top             =   2835
         Width           =   1650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   9
         Top             =   2550
         Width           =   1650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   8
         Top             =   2265
         Width           =   1650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   7
         Top             =   1980
         Width           =   1650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   6
         Top             =   1695
         Width           =   1650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   5
         Top             =   1410
         Width           =   1650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   4
         Top             =   1125
         Width           =   1650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   3
         Top             =   840
         Width           =   1650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   2
         Top             =   555
         Width           =   1650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Data Preparation"
         Height          =   285
         Left            =   2835
         TabIndex        =   1
         Top             =   270
         Width           =   1650
      End
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   6435
         Top             =   3675
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
   End
End
Attribute VB_Name = "frmPublish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub Form_Load()
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mRowCount As Integer
        
       XPC.InitIDESubClassing
       
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mRowCount = 1
            mSQL = "Select * From faPeriodicity Where intTypeID = 9"
            Rec.Open mSQL, mCnn
            While Not Rec.EOF
                vsMonthly.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity)
                vsMonthly.Cell(flexcpFontBold, mRowCount, 0) = True
                mRowCount = mRowCount + 1
                Rec.MoveNext
            Wend
            Rec.Close
        End If
    End Sub
