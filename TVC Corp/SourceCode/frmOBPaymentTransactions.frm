VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmOBPaymentTransactions 
   BorderStyle     =   0  'None
   Caption         =   "Payments"
   ClientHeight    =   7920
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   14325
   Icon            =   "frmOBPaymentTransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSub2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12240
      TabIndex        =   18
      Top             =   6885
      Width           =   1950
   End
   Begin VB.CommandButton cmdFullView 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Full View"
      Height          =   375
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7470
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7470
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
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
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7470
      Width           =   1185
   End
   Begin VB.TextBox txtSub1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12150
      TabIndex        =   9
      Top             =   5040
      Width           =   1950
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
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
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7470
      Width           =   1185
   End
   Begin VB.CommandButton cmdSavePayment 
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
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7470
      Width           =   1185
   End
   Begin VB.TextBox txtPtotal 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   12195
      TabIndex        =   0
      Top             =   7245
      Width           =   1995
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   12870
      Top             =   8595
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fraPart1 
      Caption         =   "  PART I (PANCHAYAT FUNDS)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   14235
      Begin VSFlex8LCtl.VSFlexGrid vsGridPPanchayatFunds 
         Height          =   1230
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   14055
         _cx             =   24791
         _cy             =   2170
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOBPaymentTransactions.frx":1CCA
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
   End
   Begin VB.Frame fraPart2 
      Caption         =   "  PART II (DEBT HEADS)  "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   0
      TabIndex        =   5
      Top             =   1710
      Width           =   14280
      Begin VSFlex8LCtl.VSFlexGrid vsGridPaymentDebtHeads 
         Height          =   1185
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   14100
         _cx             =   24871
         _cy             =   2090
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOBPaymentTransactions.frx":1E45
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
   End
   Begin VB.Frame fraDE 
      Caption         =   "  HEADS NOT IN SE   "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   45
      TabIndex        =   6
      Top             =   3285
      Width           =   14235
      Begin VSFlex8LCtl.VSFlexGrid vsDE 
         Height          =   1455
         Left            =   90
         TabIndex        =   12
         Top             =   225
         Width           =   14055
         _cx             =   24791
         _cy             =   2566
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOBPaymentTransactions.frx":1FC8
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
   End
   Begin VB.Frame fraRecoveries 
      Caption         =   "  RECOVERIES  "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   45
      TabIndex        =   8
      Top             =   5400
      Width           =   14235
      Begin VSFlex8LCtl.VSFlexGrid vsGridPRecoveries 
         Height          =   1005
         Left            =   90
         TabIndex        =   11
         Top             =   225
         Width           =   14100
         _cx             =   24871
         _cy             =   1773
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOBPaymentTransactions.frx":20DD
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
   End
   Begin VB.Label Label3 
      Caption         =   "SubTotal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11430
      TabIndex        =   19
      Top             =   6975
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "SubTotal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11340
      TabIndex        =   10
      Top             =   5085
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   510
      Left            =   5895
      TabIndex        =   7
      Top             =   4050
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total  Amount"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10665
      TabIndex        =   3
      Top             =   7335
      Width           =   1545
   End
End
Attribute VB_Name = "frmOBPaymentTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private Sub cmdBackPayment_Click()
        'frmOBReceiptsTransactions.Show
    End Sub

    Private Sub cmdCancel_Click()
        If MsgBox("You haven't finished the Wizard, are you sure you want to quit?   ", vbQuestion + vbYesNo, "Close Wizard") = vbYes Then
            Unload Me
            frmOpeningWizard.cmdCancel_Click
        End If
    End Sub
    Private Sub cmdFullView_Click()
        cmdFullView.Visible = False
        If cmdFullView.Tag = 1 Then
            If cmdFullView.Caption = "Full View" Then
                fraPart1.Height = 6920
                txtSub1.Visible = False
                vsGridPPanchayatFunds.Height = 6920
            Else
                fraPart1.Height = 1680
                txtSub1.Visible = True
                vsGridPPanchayatFunds.Height = 1230
            End If
        ElseIf cmdFullView.Tag = 2 Then
            If cmdFullView.Caption = "Full View" Then
                fraPart2.Top = fraPart1.Top
                fraPart1.Visible = False
                fraPart2.Height = 6920
                txtSub1.Visible = False
                vsGridPaymentDebtHeads.Height = 6920
            Else
                fraPart1.Visible = True
                fraPart2.Top = fraPart1.Height
                fraPart2.Height = 1500
                txtSub1.Visible = True
                vsGridPaymentDebtHeads.Height = 1185
            End If
        ElseIf cmdFullView.Tag = 3 Then
            If cmdFullView.Caption = "Full View" Then
            
                fraDE.Top = fraPart1.Top
                
                fraPart1.Visible = False
                fraPart2.Visible = False
                fraDE.Height = 6920
                txtSub1.Visible = False
                vsDE.Height = 6920
            Else
                fraPart1.Visible = True
                fraPart2.Visible = True
                fraDE.Top = 3285
                fraDE.Height = 1770
                txtSub1.Visible = True
                vsDE.Height = 1455
            End If
        ElseIf cmdFullView.Tag = 4 Then
            If cmdFullView.Caption = "Full View" Then
                fraRecoveries.Top = fraPart1.Top
                fraPart1.Visible = False
                fraPart2.Visible = False
                fraDE.Visible = False
                fraRecoveries.Height = 6920
                txtSub1.Visible = False
                vsGridPRecoveries.Height = 6920
            Else
                fraPart1.Visible = True
                fraPart2.Visible = True
                fraDE.Visible = True
                fraRecoveries.Top = 5445
                fraRecoveries.Height = 1950
                txtSub1.Visible = True
                vsGridPRecoveries.Height = 1365
            End If
        End If
    End Sub
    Private Sub cmdNext_Click()
        Me.Hide
        frmOpeningWizard.cmdNext_Click
    End Sub
    Private Sub cmdPrevious_Click()
         Me.Hide
'         Unload Me
         frmOBReceiptTransactions.Form_Load
         frmOpeningWizard.cmdPre_Click
    End Sub
    Private Sub cmdSavePayment_Click()
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objCL       As New clsAccounts
        Dim objFun      As New clsFunction
        Dim mSql        As String
        Dim mCnt        As Integer
        Dim mintOBRPTransactionsID  As Double
        Dim AccID       As Integer
        Dim mAccCode    As String
        Dim mArrIN      As Variant
        Dim mFunID      As Integer
        Dim mFunCode    As String
        Dim mAmount     As Double
        Dim mSaveFlag   As Boolean
        Dim mVrType     As Integer
        Dim mHeadType   As Integer  '1=Panchat Head,2-Debt head
        Dim mRecovry    As Integer  'Recovey=1 else 0
        Dim SEAccID     As Integer
        Dim SEAccCode   As String
        
        If cmdFullView.Caption = "Orginal View" Then
            
        End If
        '-------
        'Heads Not in Single Entry
        '-------
        Me.MousePointer = vbArrowHourglass
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        For mCnt = 1 To vsDE.Rows - 1
            If vsDE.TextMatrix(mCnt, 0) <> "" Then
                 If val(vsDE.TextMatrix(mCnt, 3)) <> 0 Then
                     If vsDE.TextMatrix(mCnt, 7) = "" Or val(vsDE.TextMatrix(mCnt, 7)) = 0 Then
                         mintOBRPTransactionsID = -1
                     Else
                          mintOBRPTransactionsID = val(vsDE.TextMatrix(mCnt, 7))
                     End If
                    mAmount = vsDE.TextMatrix(mCnt, 3)
                    AccID = val(vsDE.TextMatrix(mCnt, 5))
                    objCL.SetAccounts (AccID)
                    mAccCode = objCL.AccountCode
                    mFunID = val(vsDE.TextMatrix(mCnt, 6))
                    objFun.SetFunctionByID (mFunID)
                    mFunCode = objFun.FunctionCode
                    mArrIN = Array(mintOBRPTransactionsID, Null, Null, AccID, mAccCode, mFunID, mFunCode, mAmount, _
                    20, 0, 3, Null, Null, "Panchat Head Amount" & gbTransactionDate)
                    objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                    mSaveFlag = True
                Else
                    MsgBox "Amount Not Entered..", vbApplicationModal
                    Exit Sub
                End If
            End If
        Next
        
        mintOBRPTransactionsID = -1
        ''''-------
          'Panchayat Fund Part I
        '----------
            For mCnt = 1 To vsGridPPanchayatFunds.Rows - 1
                If vsGridPPanchayatFunds.TextMatrix(mCnt, 0) <> "" Then
                    If val(vsGridPPanchayatFunds.TextMatrix(mCnt, 5)) <> 0 Then
                         If vsGridPPanchayatFunds.TextMatrix(mCnt, 10) = "" Or val(vsGridPPanchayatFunds.TextMatrix(mCnt, 10)) = 0 Then
                             mintOBRPTransactionsID = -1
                         Else
                              mintOBRPTransactionsID = val(vsGridPPanchayatFunds.TextMatrix(mCnt, 10))
                         End If
                        SEAccID = val(vsGridPPanchayatFunds.TextMatrix(mCnt, 7))
                        SEAccCode = vsGridPPanchayatFunds.TextMatrix(mCnt, 0)
                        mAmount = vsGridPPanchayatFunds.TextMatrix(mCnt, 5)
                        AccID = val(vsGridPPanchayatFunds.TextMatrix(mCnt, 8))
                        objCL.SetAccounts (AccID)
                        mAccCode = objCL.AccountCode
                        mFunID = val(vsGridPPanchayatFunds.TextMatrix(mCnt, 9))
                        objFun.SetFunctionByID (mFunID)
                        mFunCode = objFun.FunctionCode
                        mRecovry = 0
                        mHeadType = 1
                        mVrType = 20
                        mArrIN = Array(mintOBRPTransactionsID, SEAccID, SEAccCode, AccID, mAccCode, mFunID, mFunCode, mAmount, _
                        mVrType, mRecovry, mHeadType, Null, Null, "Panchat Head Amount " & gbTransactionDate)
                        objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                         mSaveFlag = True
                    Else
                        MsgBox "Plesae Enter Amount..   ", vbApplicationModal
                        Exit Sub
                    End If
                    
                End If
            Next
        ''''-------
          'Panchayat Debt Fund Part II
        '----------
            mintOBRPTransactionsID = -1
            For mCnt = 1 To vsGridPaymentDebtHeads.Rows - 1
                If vsGridPaymentDebtHeads.TextMatrix(mCnt, 0) <> "" Then
                    If val(vsGridPaymentDebtHeads.TextMatrix(mCnt, 5)) <> 0 Then
                         If vsGridPaymentDebtHeads.TextMatrix(mCnt, 10) = "" Or val(vsGridPaymentDebtHeads.TextMatrix(mCnt, 10)) = 0 Then
                             mintOBRPTransactionsID = -1
                         Else
                              mintOBRPTransactionsID = val(vsGridPaymentDebtHeads.TextMatrix(mCnt, 10))
                         End If
                        SEAccID = val(vsGridPaymentDebtHeads.TextMatrix(mCnt, 7))
                        SEAccCode = vsGridPaymentDebtHeads.TextMatrix(mCnt, 0)
                        mAmount = vsGridPaymentDebtHeads.TextMatrix(mCnt, 5)
                        AccID = val(vsGridPaymentDebtHeads.TextMatrix(mCnt, 8))
                        objCL.SetAccounts (AccID)
                        mAccCode = objCL.AccountCode
                        mFunID = val(vsGridPaymentDebtHeads.TextMatrix(mCnt, 9))
                        objFun.SetFunctionByID (mFunID)
                        mFunCode = objFun.FunctionCode
                        mRecovry = 0
                        mHeadType = 2
                        mVrType = 20
                        mArrIN = Array(mintOBRPTransactionsID, SEAccID, SEAccCode, AccID, mAccCode, mFunID, mFunCode, mAmount, _
                        mVrType, mRecovry, mHeadType, Null, Null, "Panchat Head Amount " & gbTransactionDate)
                        objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                         mSaveFlag = True
                    Else
                        MsgBox "Plesae Enter Amount..   ", vbApplicationModal
                        Exit Sub
                    End If
                End If
            Next
        '----------
          'Recoveries
        '----------
                mintOBRPTransactionsID = -1
                For mCnt = 1 To vsGridPRecoveries.Rows - 1
                If vsGridPRecoveries.TextMatrix(mCnt, 0) <> "" Then
                    If val(vsGridPRecoveries.TextMatrix(mCnt, 5)) <> 0 Then
                         If vsGridPRecoveries.TextMatrix(mCnt, 10) = "" Or val(vsGridPRecoveries.TextMatrix(mCnt, 10)) = 0 Then
                             mintOBRPTransactionsID = -1
                         Else
                              mintOBRPTransactionsID = val(vsGridPRecoveries.TextMatrix(mCnt, 10))
                         End If
                        SEAccID = val(vsGridPRecoveries.TextMatrix(mCnt, 7))
                        SEAccCode = vsGridPRecoveries.TextMatrix(mCnt, 0)
                        mAmount = vsGridPRecoveries.TextMatrix(mCnt, 5)
                        AccID = val(vsGridPRecoveries.TextMatrix(mCnt, 8))
                        objCL.SetAccounts (AccID)
                        mAccCode = objCL.AccountCode
                        mFunID = val(vsGridPRecoveries.TextMatrix(mCnt, 9))
                        objFun.SetFunctionByID (mFunID)
                        mFunCode = objFun.FunctionCode
                        mRecovry = 1
                        mHeadType = 0
                        mVrType = 0
                        mArrIN = Array(mintOBRPTransactionsID, SEAccID, SEAccCode, AccID, mAccCode, mFunID, mFunCode, mAmount, _
                        mVrType, mRecovry, mHeadType, Null, Null, "Panchat Head Amount " & gbTransactionDate)
                        objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                        mSaveFlag = True
                    Else
                        MsgBox "Plesae Enter Amount..   ", vbApplicationModal
                        Exit Sub
                    End If
                End If
            Next
            
        '----------
'        If SaveGridData(vsGridPPanchayatFunds) = False Then
'            Exit Sub
'        Else
'            mSaveFlag = True
'        End If
'        If SaveGridData(vsGridPaymentDebtHeads) = False Then
'            Exit Sub
'        Else
'            mSaveFlag = True
'        End If
'        If SaveGridData(vsGridPRecoveries) = False Then
'
'            Exit Sub
'        Else
'            mSaveFlag = True
'        End If
        If mSaveFlag = True Then
            MsgBox "Saved sucessfully", vbApplicationModal
        End If
        Call FillGridData(vsGridPPanchayatFunds)
        Call FillGridData(vsGridPaymentDebtHeads)
        Call FillGridData(vsGridPRecoveries)
        Call FillGridData(vsDE)
        
        Me.MousePointer = vbArrow
        Me.Hide
        frmOpeningWizard.cmdNext_Click
    End Sub
    Private Sub FillGridData(vsGrid As VSFlexGrid)
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objAcc      As New clsAccounts
        Dim objFun      As New clsFunction
        Dim mSql        As String
        Dim mCnt        As Integer
        Dim mSEHead     As String
        Dim mDEHead     As String
        Dim mFunction   As String
        If (vsGrid.Name = vsGridPPanchayatFunds.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where intVoucherTypeID=20 AND tnyHeadType=1  Order By tnyHeadType,intOBRPTransactionsID"
        ElseIf (vsGrid.Name = vsGridPaymentDebtHeads.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where intVoucherTypeID=20 AND tnyHeadType=2  Order By tnyHeadType,intOBRPTransactionsID"
        ElseIf (vsGrid.Name = vsGridPRecoveries.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where tnyRecovery=1  Order By tnyHeadType,intOBRPTransactionsID"
        ElseIf (vsGrid.Name = vsDE.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where  intVoucherTypeID=20 AND tnyHeadType=3  Order By tnyHeadType,intOBRPTransactionsID"
        End If
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        vsGrid.Clear 1, 0
        vsGrid.Rows = 2
        mCnt = 1
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                If (vsGrid.Name = vsDE.Name) Then
                    objAcc.SetAccountID (IIf(IsNull(Rec!intAccountHeadID), -1, Rec!intAccountHeadID))
                    mDEHead = objAcc.AccountHead
                    objFun.SetFunctionByID (IIf(IsNull(Rec!intFunctionID), -1, Rec!intFunctionID))
                    mFunction = objFun.FunctionName
                    vsGrid.TextMatrix(mCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    vsGrid.TextMatrix(mCnt, 1) = mDEHead
                    vsGrid.TextMatrix(mCnt, 2) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode) + " " + mFunction 'Fuction Code + Function
                    vsGrid.TextMatrix(mCnt, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    If (IIf(IsNull(Rec!vchVoucherNo), 0, Rec!vchVoucherNo)) > 0 Then
                        vsGrid.ColHidden(4) = False
                    End If
                    vsGrid.TextMatrix(mCnt, 4) = IIf(IsNull(Rec!vchVoucherNo), "", Rec!vchVoucherNo)
                    vsGrid.TextMatrix(mCnt, 5) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                    vsGrid.TextMatrix(mCnt, 6) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    vsGrid.TextMatrix(mCnt, 7) = IIf(IsNull(Rec!intOBRPTransactionsID), "", Rec!intOBRPTransactionsID)
                Else
                    mSql = ""
                    mSql = "SELECT vchSEHead FROM faSEAccountHeads WHERE intSEHeadID= " & IIf(IsNull(Rec!intSEAccountHeadID), -1, Rec!intSEAccountHeadID)
                    RecChild.Open mSql, mCnn
                    If Not (RecChild.EOF And RecChild.BOF) Then
                        mSEHead = IIf(IsNull(RecChild!vchSEHead), "", RecChild!vchSEHead)
                    End If
                    RecChild.Close
                    objAcc.SetAccountID (IIf(IsNull(Rec!intAccountHeadID), -1, Rec!intAccountHeadID))
                    mDEHead = objAcc.AccountHead
                    objFun.SetFunctionByID (IIf(IsNull(Rec!intFunctionID), -1, Rec!intFunctionID))
                    mFunction = objFun.FunctionName
                    vsGrid.TextMatrix(mCnt, 0) = IIf(IsNull(Rec!vchSEAccountHeadCode), "", Rec!vchSEAccountHeadCode)
                    vsGrid.TextMatrix(mCnt, 1) = mSEHead
                    vsGrid.TextMatrix(mCnt, 2) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    vsGrid.TextMatrix(mCnt, 3) = mDEHead
                    vsGrid.TextMatrix(mCnt, 4) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode) + " " + mFunction 'Fuction Code + Function
                    vsGrid.TextMatrix(mCnt, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsGrid.TextMatrix(mCnt, 6) = IIf(IsNull(Rec!vchVoucherNo), "", Rec!vchVoucherNo)
                    vsGrid.TextMatrix(mCnt, 7) = IIf(IsNull(Rec!intSEAccountHeadID), "", Rec!intSEAccountHeadID)
                    vsGrid.TextMatrix(mCnt, 8) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                    vsGrid.TextMatrix(mCnt, 9) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    vsGrid.TextMatrix(mCnt, 10) = IIf(IsNull(Rec!intOBRPTransactionsID), "", Rec!intOBRPTransactionsID)
                End If
                vsGrid.AddItem ""
                mCnt = mCnt + 1
                Rec.MoveNext
                
            Wend
        End If
        Rec.Close
        mCnn.Close
        Call CalculateTotal
    End Sub
'    Private Function SaveGridData(vsGrid As VSFlexGrid) As Boolean
'        Dim Rec         As New ADODB.Recordset
'        Dim mCnn        As New ADODB.Connection
'        Dim objdb       As New clsDB
'        Dim objCL       As New clsAccounts
'        Dim objFun      As New clsFunction
'        Dim mSql        As String
'        Dim mCnt        As Integer
'        Dim mintOBRPTransactionsID  As Double
'        Dim AccID       As Integer
'        Dim SEAccID     As Integer
'        Dim SEAccCode   As String
'        Dim mAccCode    As String
'        Dim mArrIn      As Variant
'        Dim mFunID      As Integer
'        Dim mFunCode    As String
'        Dim mAmount     As Double
'        Dim mVrType     As Integer
'        Dim mHeadType   As Integer  '1=Panchat Head,2-Debt head
'        Dim mRecovry    As Integer  'Recovey=1 else 0
'        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'         For mCnt = 1 To vsGrid.Rows - 1
'            If vsGrid.TextMatrix(mCnt, 0) <> "" Then
'                If val(vsGrid.TextMatrix(mCnt, 5)) <> 0 Then
'                     If vsGrid.TextMatrix(mCnt, 10) = "" Or val(vsGrid.TextMatrix(mCnt, 10)) = 0 Then
'                         mintOBRPTransactionsID = -1
'                     Else
'                          mintOBRPTransactionsID = val(vsGrid.TextMatrix(mCnt, 10))
'                     End If
'                    SEAccID = val(vsGrid.TextMatrix(mCnt, 7))
'                    SEAccCode = vsGrid.TextMatrix(mCnt, 0)
'                    mAmount = vsGrid.TextMatrix(mCnt, 5)
'                    AccID = val(vsGrid.TextMatrix(mCnt, 8))
'                    objCL.SetAccounts (AccID)
'                    mAccCode = objCL.AccountCode
'                    mFunID = val(vsGrid.TextMatrix(mCnt, 9))
'                    objFun.SetFunctionByID (mFunID)
'                    mFunCode = objFun.FunctionCode
'                    If vsGrid = vsGridPRecoveries Then
'                         mRecovry = 1
'                         mHeadType = 0
'                         mVrType = 0
'                    ElseIf vsGrid = vsGridPPanchayatFunds Then
'                         mRecovry = 0
'                         mHeadType = 1
'                         mVrType = 20
'                    ElseIf vsGrid = vsGridPaymentDebtHeads Then
'                         mRecovry = 0
'                         mHeadType = 2
'                         mVrType = 20
'                    End If
'                    mArrIn = Array(mintOBRPTransactionsID, SEAccID, SEAccCode, AccID, mAccCode, mFunID, mFunCode, mAmount, _
'                    mVrType, mRecovry, mHeadType, Null, Null, "Panchat Head Amount " & gbTransactionDate)
'                    objdb.ExecuteSP "spSaveOBRPTransactions", mArrIn, , , mCnn, adCmdStoredProc
'                Else
'                    MsgBox "Plesae Enter Amount..   ", vbApplicationModal
'                    SaveGridData = False
'                    Exit Function
'                End If
'            End If
'        Next
'        SaveGridData = True
'        mCnn.Close
'    End Function
    Private Sub CalculateTotal()
        Dim mPAmt       As Double
        Dim mDAmt       As Double
        Dim mDeamt      As Double
        Dim mSub1       As Double
        Dim mRecvry     As Double
        Dim mSub2       As Double
        Dim mTot        As Double
        Dim mCnt        As Integer
        mPAmt = 0
        mDAmt = 0
        mDeamt = 0
        mSub1 = 0
        mRecvry = 0
        mSub2 = 0
        mTot = 0
        For mCnt = 1 To vsGridPPanchayatFunds.Rows - 1
            If vsGridPPanchayatFunds.TextMatrix(mCnt, 5) <> "" Then
                mPAmt = mPAmt + vsGridPPanchayatFunds.TextMatrix(mCnt, 5)
            End If
        Next
        For mCnt = 1 To vsGridPaymentDebtHeads.Rows - 1
            If vsGridPaymentDebtHeads.TextMatrix(mCnt, 5) <> "" Then
                mDAmt = mDAmt + vsGridPaymentDebtHeads.TextMatrix(mCnt, 5)
            End If
        Next
        For mCnt = 1 To vsDE.Rows - 1
            If vsDE.TextMatrix(mCnt, 3) <> "" Then
                mDeamt = mDeamt + vsDE.TextMatrix(mCnt, 3)
            End If
        Next
        For mCnt = 1 To vsGridPRecoveries.Rows - 1
            If vsGridPRecoveries.TextMatrix(mCnt, 5) <> "" Then
                mRecvry = mRecvry + vsGridPRecoveries.TextMatrix(mCnt, 5)
            End If
        Next
        mSub1 = mPAmt + mDAmt + mDeamt
        mSub2 = mRecvry
        mTot = mSub1 + mSub2
        txtSub1.Text = mSub1
        txtSub2.Text = mSub2
        txtPtotal.Text = mTot
    End Sub

    Private Sub Form_Activate()
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSavePayment.Enabled = False
        End If
    End Sub

    Public Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        vsGridPPanchayatFunds.ColComboList(0) = "|..."
        vsDE.ColComboList(0) = "|..."
        vsGridPaymentDebtHeads.ColComboList(0) = "|..."
        vsGridPRecoveries.ColComboList(0) = "|..."
        Call FillGridData(vsGridPPanchayatFunds)
        Call FillGridData(vsGridPaymentDebtHeads)
        Call FillGridData(vsGridPRecoveries)
        Call FillGridData(vsDE)
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSavePayment.Enabled = False
        End If
    End Sub
    Private Sub txtPtotal_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub
    Private Sub txtSub1_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub
    Private Sub txtSub2_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub
    Private Sub vsDE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsDE.SetFocus
        If Button = 2 Then
            If fraDE.Height > 2000 Then
                cmdFullView.Caption = "Orginal View"
            Else
                cmdFullView.Caption = "Full View"
            End If
                cmdFullView.Top = X
                cmdFullView.Left = Y
                cmdFullView.Visible = True
                cmdFullView.Tag = 3
        End If
    End Sub
    Private Sub vsGridPaymentDebtHeads_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsGridPaymentDebtHeads.SetFocus
        If Button = 2 Then
            If fraPart2.Height > 2000 Then
                cmdFullView.Caption = "Orginal View"
            Else
                cmdFullView.Caption = "Full View"
            End If
                cmdFullView.Top = X
                cmdFullView.Left = Y
                cmdFullView.Visible = True
                cmdFullView.Tag = 2
        End If
    End Sub
    Private Sub vsGridPPanchayatFunds_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 5 Then
            If val(vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 5)) > 0 Then
                If vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 2) = "" Or vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 4) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 5) = ""
                    Exit Sub
                End If
            End If
            Call CalculateTotal
        End If
    End Sub
    Private Sub vsDE_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 3 Then
            If val(vsDE.TextMatrix(vsDE.Row, 3)) > 0 Then
                If vsDE.TextMatrix(vsDE.Row, 0) = "" Or vsDE.TextMatrix(vsDE.Row, 6) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsDE.TextMatrix(vsDE.Row, 3) = ""
                    Exit Sub
                End If
            End If
            Call CalculateTotal
        End If
    End Sub
    Private Sub vsGridPaymentDebtHeads_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 5 Then
            If val(vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 5)) > 0 Then
                If vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 2) = "" Or vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 4) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 5) = ""
                    Exit Sub
                End If
            End If
            Call CalculateTotal
        End If
    End Sub
    Private Sub vsGridPPanchayatFunds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            vsGridPPanchayatFunds.SetFocus
            If Button = 2 Then
                If fraPart1.Height > 2000 Then
                    cmdFullView.Caption = "Orginal View"
                Else
                    cmdFullView.Caption = "Full View"
                End If
                    cmdFullView.Top = X
                    cmdFullView.Left = Y
                    cmdFullView.Visible = True
                    cmdFullView.Tag = 1
            End If
    End Sub
    Private Sub vsGridPRecoveries_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 5 Then
            If val(vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 5)) > 0 Then
                If vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 2) = "" Or vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 4) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 5) = ""
                    Exit Sub
                End If
            End If
            Call CalculateTotal
        End If
    End Sub

    Private Sub vsGridPPanchayatFunds_Click()
        If val(vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 10)) > 0 Then
'            MsgBox "Data Already Exits.. Do you want to Change. "
'            vsGridPPanchayatFunds.Editable = flexEDNone
'            Exit Sub
        Else
            vsGridPPanchayatFunds.Editable = flexEDKbdMouse
        End If
    End Sub
    Private Sub vsDE_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsDE.Col = 3 Then
            vsDE.EditMaxLength = 15
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47) Then KeyAscii = 0
        ElseIf KeyAscii = 13 Then
            If vsDE.TextMatrix(vsDE.Row, 3) <> "" And vsDE.Row = vsDE.Rows - 1 Then
                vsDE.Rows = vsDE.Rows + 1
            End If
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub vsGridPRecoveries_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGridPRecoveries.Col = 5 Then
            vsGridPRecoveries.EditMaxLength = 15
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) And KeyAscii <> 47) Then KeyAscii = 0
            If KeyAscii = 13 Then
                If vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 5) <> "" And vsGridPRecoveries.Row = vsGridPRecoveries.Rows - 1 Then
                    vsGridPRecoveries.Rows = vsGridPRecoveries.Rows + 1
                End If
            End If
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub vsGridPaymentDebtHeads_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGridPaymentDebtHeads.Col = 5 Then
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) And KeyAscii <> 47) Then KeyAscii = 0
            vsGridPaymentDebtHeads.EditMaxLength = 15
            If KeyAscii = 13 Then
                If vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 5) <> "" And vsGridPaymentDebtHeads.Row = vsGridPaymentDebtHeads.Rows - 1 Then
                    vsGridPaymentDebtHeads.Rows = vsGridPaymentDebtHeads.Rows + 1
                End If
            End If
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub vsGridPPanchayatFunds_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGridPPanchayatFunds.Col = 5 Then
            vsGridPPanchayatFunds.EditMaxLength = 15
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) And KeyAscii <> 47) Then KeyAscii = 0
            If KeyAscii = 13 Then
                If vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 5) <> "" And vsGridPPanchayatFunds.Row = vsGridPPanchayatFunds.Rows - 1 Then
                    vsGridPPanchayatFunds.Rows = vsGridPPanchayatFunds.Rows + 1
                End If
            End If
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub FillDEHeas4SE(vsGrid As VSFlexGrid)
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim mCnt        As Integer
        mSql = "SELECT Distinct  faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead," & vbNewLine
        mSql = mSql + " faAccountHeads.intAccountHeadID , faFunctions.vchFunctionCode, faFunctions.vchFunction, faFunctions.intFunctionID" & vbNewLine
        mSql = mSql + " From faSEAccountHeads"
        mSql = mSql + " INNER JOIN  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode" & vbNewLine
'        mSql = mSql + " LEFT Join faTransactionTypeChild ON faTransactionTypeChild.vchAccountHeadCode=faAccountHeads.vchAccountHeadCode" & vbNewLine
'        mSql = mSql + " Inner Join faTransactionType ON faTransactionTypeChild.intTransactionTypeID=faTransactionType.intTransactionTypeID" & vbNewLine
        mSql = mSql + " Inner Join faFunctions ON faFunctions.vchFunctionCode=faSEAccountHeads.vchFunctionCode" & vbNewLine
        mSql = mSql + " Where tinHiddenFlag <> 1 And intSEHeadID = " & val(vsGrid.TextMatrix(vsGrid.Row, 7))
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.TextMatrix(vsGrid.Row, 2) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
            vsGrid.TextMatrix(vsGrid.Row, 3) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            vsGrid.TextMatrix(vsGrid.Row, 4) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode) & " " & IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            vsGrid.TextMatrix(vsGrid.Row, 8) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
            vsGrid.TextMatrix(vsGrid.Row, 9) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
        Else
            MsgBox "Double Entry Account Head not Mapped", vbApplicationModal
            vsGrid.RemoveItem (vsGrid.Row)
            vsGrid.AddItem (" ")
            Exit Sub
        End If
            
    End Sub
    Private Sub vsDE_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        If Col = 0 Then
            If val(vsDE.TextMatrix(vsDE.Row, 5)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
            Dim mSql    As String
            mSql = ""
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
            mSql = mSql + " Where tinType IN (2) AND intGroupID is Null And tinHiddenFlag <> 1 And "
            mSql = mSql + " vchAccountHeadCode not in (Select isNull(vchAccountHeadCode,'') From faSEAccountHeads)"
            mSql = mSql + " Order by vchAccountHeadCode"
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.Show vbModal
            If vsDE.FindRow(gbSearchID, 1, 5) > 0 Then
                MsgBox "Already selected this Account Head", vbApplicationModal
                Exit Sub
            End If
            If gbSearchID < 0 Then
                MsgBox "Account Head Not Selected", vbApplicationModal
                Exit Sub
            End If
            vsDE.TextMatrix(vsDE.Row, 0) = Token(gbSearchStr, " ")
            vsDE.TextMatrix(vsDE.Row, 1) = Trim(gbSearchStr)
            vsDE.TextMatrix(vsDE.Row, 5) = gbSearchID
            gbSearchID = -1
            gbSearchStr = ""
            mSql = ""
            mSql = mSql + " SELECT Distinct faFunctions.vchFunctionCode,faFunctions.vchFunction ,faFunctions.intFunctionID"
            mSql = mSql + " FROM faAccountHeads LEFT Join faTransactionTypeChild ON faTransactionTypeChild.vchAccountHeadCode=faAccountHeads.vchAccountHeadCode"
            mSql = mSql + " Inner Join faTransactionType ON faTransactionTypeChild.intTransactionTypeID=faTransactionType.intTransactionTypeID"
            mSql = mSql + " Inner Join faFunctions ON faFunctions.intFunctionID=faTransactionType.intFunctionID"
            mSql = mSql + " Where faAccountHeads.intAccountHeadID = " & val(vsDE.TextMatrix(vsDE.Row, 5))
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                vsDE.TextMatrix(vsDE.Row, 2) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode) & " " & IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                vsDE.TextMatrix(vsDE.Row, 6) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            Else
                vsDE.ColComboList(2) = "|..."
            End If
        End If
        If Col = 2 Then
            frmSearchFunction.Show vbModal
            vsDE.TextMatrix(vsDE.Row, 2) = Trim(gbSearchStr)
            vsDE.TextMatrix(vsDE.Row, 6) = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    Private Sub DeleteRecord(mID As Integer)
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        mSql = "Delete From faOBRPTransactions Where intOBRPTransactionsID=" & mID
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCnn.Execute mSql
    End Sub
     Private Sub vsGridPaymentDebtHeads_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim mSql As String
        If Col = 0 Then
            If val(vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 7)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
            frmSearchSEPanchayatAccountHeads.intModeOfTransaction = 4
            frmSearchSEPanchayatAccountHeads.Show vbModal
            If vsGridPaymentDebtHeads.FindRow(gbSearchID, 1, 7) > 0 Then
                MsgBox "Already selected this Account Head", vbApplicationModal
                Exit Sub
            End If
            If gbSearchID < 0 Then
                MsgBox "Account Head Not Selected", vbApplicationModal
                Exit Sub
            End If
            vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 0) = gbSearchCode
            vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 1) = gbSearchStr
            vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 7) = gbSearchID
            Call FillDEHeas4SE(vsGridPaymentDebtHeads)
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
        End If
    End Sub
    Private Sub vsGridPPanchayatFunds_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If Col = 0 Then
            If val(vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 7)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
            frmSearchSEPanchayatAccountHeads.intModeOfTransaction = 3
            frmSearchSEPanchayatAccountHeads.Show vbModal
            If vsGridPPanchayatFunds.FindRow(gbSearchID, 1, 7) > 0 Then
                MsgBox "Already selected this Account Head", vbApplicationModal
                Exit Sub
            End If
            If gbSearchID < 0 Then
                MsgBox "Account Head Not Selected", vbApplicationModal
                Exit Sub
            End If
            vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 0) = gbSearchCode
            vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 1) = gbSearchStr
            vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 7) = gbSearchID
            'vsGridPPanchayatFunds.ColEditMask(0) = True
            Call FillDEHeas4SE(vsGridPPanchayatFunds)
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
        End If
    End Sub
    Private Sub vsGridPRecoveries_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If Col = 0 Then
            If val(vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 7)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
            frmSearchSEPanchayatAccountHeads.intModeOfTransaction = 5
            frmSearchSEPanchayatAccountHeads.Show vbModal
            If vsGridPRecoveries.FindRow(gbSearchID, 1, 7) > 0 Then
                MsgBox "Already selected this Account Head", vbApplicationModal
                Exit Sub
            End If
            If gbSearchID < 0 Then
                MsgBox "Account Head Not Selected", vbApplicationModal
                Exit Sub
            End If
            vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 0) = gbSearchCode
            vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 1) = gbSearchStr
            vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 7) = gbSearchID
            Call FillDEHeas4SE(vsGridPRecoveries)
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
        End If
    End Sub
    Private Sub vsDE_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                If (val(vsDE.TextMatrix(vsDE.Row, 7)) > 1) Then
                    Call DeleteRecord(val(vsDE.TextMatrix(vsDE.Row, 7)))
                    vsDE.RemoveItem (vsDE.Row)
                ElseIf vsDE.Rows > 1 Then
                    vsDE.RemoveItem (vsDE.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If (vsDE.TextMatrix(vsDE.Row, 3) <> "" And vsDE.TextMatrix(vsDE.Row, 0) <> "") And vsDE.Row = vsDE.Rows - 1 Then
                vsDE.Rows = vsDE.Rows + 1
            End If
        End If
    End Sub
    Private Sub vsGridPaymentDebtHeads_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
               If (val(vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 10)) > 1) Then
                    Call DeleteRecord(val(vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 10)))
                    vsGridPaymentDebtHeads.RemoveItem (vsGridPaymentDebtHeads.Row)
                ElseIf vsGridPaymentDebtHeads.Rows > 1 Then
                    vsGridPaymentDebtHeads.RemoveItem (vsGridPaymentDebtHeads.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If vsGridPaymentDebtHeads.TextMatrix(vsGridPaymentDebtHeads.Row, 3) <> "" And vsGridPaymentDebtHeads.Row = vsGridPaymentDebtHeads.Rows - 1 Then
                vsGridPaymentDebtHeads.Rows = vsGridPaymentDebtHeads.Rows + 1
            End If
        End If
    End Sub
    Private Sub vsGridPRecoveries_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                If (val(vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 10)) > 1) Then
                    Call DeleteRecord(val(vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 10)))
                    vsGridPRecoveries.RemoveItem (vsGridPRecoveries.Row)
                ElseIf vsGridPRecoveries.Rows > 1 Then
                    vsGridPRecoveries.RemoveItem (vsGridPRecoveries.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If vsGridPRecoveries.TextMatrix(vsGridPRecoveries.Row, 5) <> "" And vsGridPRecoveries.Row = vsGridPRecoveries.Rows - 1 Then
                vsGridPRecoveries.Rows = vsGridPRecoveries.Rows + 1
            End If
        End If
    End Sub
    Private Sub vsGridPPanchayatFunds_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                If (val(vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 10)) > 1) Then
                    Call DeleteRecord(val(vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 10)))
                    vsGridPPanchayatFunds.RemoveItem (vsGridPPanchayatFunds.Row)
                ElseIf vsGridPPanchayatFunds.Rows > 1 Then
                    vsGridPPanchayatFunds.RemoveItem (vsGridPPanchayatFunds.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If vsGridPPanchayatFunds.TextMatrix(vsGridPPanchayatFunds.Row, 5) <> "" And vsGridPPanchayatFunds.Row = vsGridPPanchayatFunds.Rows - 1 Then
                vsGridPPanchayatFunds.Rows = vsGridPPanchayatFunds.Rows + 1
            End If
        End If
    End Sub
    
    Private Sub vsGridPRecoveries_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsGridPRecoveries.SetFocus
        If Button = 2 Then
            If fraRecoveries.Height > 2000 Then
                cmdFullView.Caption = "Orginal View"
            Else
                cmdFullView.Caption = "Full View"
            End If
                cmdFullView.Top = X
                cmdFullView.Left = Y
                cmdFullView.Visible = True
                cmdFullView.Tag = 4
        End If
    End Sub
