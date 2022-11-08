VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmAFSClosingCashBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Closing Cash/Bank Books"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   Icon            =   "frmAFSClosingCashBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   270
      TabIndex        =   8
      Top             =   990
      Width           =   10950
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   7380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   90
         Width           =   3525
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1755
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   90
         Width           =   5550
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   90
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   90
         Width           =   1680
      End
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
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
      Left            =   3780
      TabIndex        =   3
      Text            =   "31-March-2013"
      Top             =   180
      Width           =   1635
   End
   Begin VB.CommandButton cmdNect 
      Caption         =   "Next"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      TabIndex        =   2
      Top             =   5490
      Width           =   825
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6075
      TabIndex        =   1
      Top             =   5490
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4365
      TabIndex        =   0
      Top             =   5490
      Width           =   825
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3525
      Left            =   225
      TabIndex        =   4
      Top             =   1845
      Width           =   10950
      _cx             =   19315
      _cy             =   6218
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAFSClosingCashBook.frx":1CCA
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
      Caption         =   "Bank"
      Height          =   240
      Left            =   270
      TabIndex        =   7
      Top             =   1530
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Cash"
      Height          =   330
      Left            =   270
      TabIndex        =   6
      Top             =   585
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Balance of Cash/Bank/Treasury As On"
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
      Left            =   315
      TabIndex        =   5
      Top             =   180
      Width           =   3480
   End
End
Attribute VB_Name = "frmAFSClosingCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub Form_Load()
        Call FillGrid
    End Sub

    Private Sub FillGrid()
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mRec        As New ADODB.Recordset
        Dim objDB      As New clsDB
        Dim arrIn       As Variant
        Dim mCnt       As Integer
        Dim mSql       As Integer
        vsGrid.Clear 2, 1
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            vsGrid.Rows = 4
            mCnt = 0
            mSql = "Select * From faAccountHeads Where intGroupId=1  Union All"
            mSql = mSql + " Select faAccountHeads.* From faAccountHeads Inner Join faBanks ON faAccountHeads.intAccountHeadID=faBanks.intAccountHeadID"
            mSql = mSql + " Where intGroupId=2"
            Set Rec = objDB.ExecuteSP(mSql, , , , mCnn, adCmdStoredProc)
            While Not (Rec.EOF)
                arrIn = Array(Rec!intAccountHeadID)
                Set mRec = objDB.ExecuteSP("spGetClosingBalance", arrIn, , , mCnn, adCmdStoredProc)
                If Not (mRec.EOF And mRec.BOF) Then
                    Rec.MoveNext
                    vsGrid.TextMatrix(mCnt, 0) = mRec!intAccountHeadID
                    'vsGrid.TextMatrix(mCnt, 0) = mRec!intAccountHeadID
                End If
                
                
            Wend
        End If
    End Sub


