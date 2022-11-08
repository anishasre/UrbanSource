VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmOBReceiptTransactions 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13905
   Icon            =   "frmOBReceiptTransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSubTotal2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11745
      TabIndex        =   21
      Text            =   "0"
      Top             =   6975
      Width           =   1860
   End
   Begin VB.CommandButton cmdRFullView 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Full View"
      Height          =   375
      Left            =   90
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7155
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7560
      Width           =   1185
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
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   1185
   End
   Begin VB.TextBox txtTotal3 
      Height          =   420
      Left            =   855
      TabIndex        =   12
      Text            =   "0"
      Top             =   7605
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtTotal2 
      Height          =   420
      Left            =   495
      TabIndex        =   11
      Text            =   "0"
      Top             =   7605
      Visible         =   0   'False
      Width           =   375
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
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   1185
   End
   Begin VB.TextBox txtTotal1 
      Height          =   375
      Left            =   45
      TabIndex        =   9
      Text            =   "0"
      Top             =   7605
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtReceiptstotal 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   11655
      TabIndex        =   1
      Text            =   "0"
      Top             =   7335
      Width           =   1950
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   12870
      Top             =   9225
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
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
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1185
   End
   Begin VB.Frame fraPart1 
      Caption         =   "PART I(PANCHAYAT FUNDS)"
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
      Left            =   135
      TabIndex        =   3
      Top             =   180
      Width           =   13605
      Begin VSFlex8LCtl.VSFlexGrid vsGridReceiptPanchayatFunds 
         Height          =   1095
         Left            =   45
         TabIndex        =   20
         Top             =   270
         Width           =   13380
         _cx             =   23601
         _cy             =   1931
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
         FormatString    =   $"frmOBReceiptTransactions.frx":1CCA
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
   End
   Begin VB.Frame fraPart2 
      Caption         =   "PART II(DEBT HEADS)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   135
      TabIndex        =   4
      Top             =   1710
      Width           =   13605
      Begin VSFlex8LCtl.VSFlexGrid vsGridReceiptDebitHeads 
         Height          =   1275
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   13380
         _cx             =   23601
         _cy             =   2249
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
         FormatString    =   $"frmOBReceiptTransactions.frx":1E43
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
   End
   Begin VB.Frame fraRecoveries 
      Caption         =   "RECOVERIES"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   90
      TabIndex        =   5
      Top             =   5355
      Width           =   13605
      Begin VSFlex8LCtl.VSFlexGrid vsGridReceiptRecoveries 
         Height          =   1095
         Left            =   45
         TabIndex        =   15
         Top             =   315
         Width           =   13470
         _cx             =   23760
         _cy             =   1931
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
         FormatString    =   $"frmOBReceiptTransactions.frx":1FBE
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
   End
   Begin VB.TextBox txtSubTotal1 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   11700
      TabIndex        =   6
      Text            =   "0"
      Top             =   5040
      Width           =   1905
   End
   Begin VB.Frame fraDE 
      Caption         =   "HEADS NOT IN SE"
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
      Left            =   135
      TabIndex        =   7
      Top             =   3375
      Width           =   13605
      Begin VSFlex8LCtl.VSFlexGrid vsDEDebt 
         Height          =   1320
         Left            =   90
         TabIndex        =   17
         Top             =   315
         Width           =   13470
         _cx             =   23760
         _cy             =   2328
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
         FormatString    =   $"frmOBReceiptTransactions.frx":213A
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9360
            TabIndex        =   18
            Top             =   1530
            Width           =   1095
         End
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   10890
      TabIndex        =   22
      Top             =   7020
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receipt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10260
      TabIndex        =   2
      Top             =   7380
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Left            =   10980
      TabIndex        =   8
      Top             =   5040
      Width           =   780
   End
End
Attribute VB_Name = "frmOBReceiptTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmdCancel_Click()
        If MsgBox("You haven't finished the Wizard, are you sure you want to quit?   ", vbQuestion + vbYesNo, "Close Wizard") = vbYes Then
            Unload Me
            frmOpeningWizard.cmdCancel_Click
        End If
    End Sub

    Private Sub cmdNext_Click()
        Me.Hide
        frmOpeningWizard.FrameNo = 3
'        Unload Me
        frmOpeningWizard.cmdNext_Click
    End Sub

    Private Sub cmdPrevious_Click()
         Me.Hide
'         Unload Me
         frmOpeningWizard.cmdPre_Click
    End Sub

Private Sub cmdRFullView_Click()
        cmdRFullView.Visible = False
        If cmdRFullView.Tag = 1 Then
            If cmdRFullView.Caption = "Full View" Then
                fraPart1.Height = 6920
                txtSubTotal1.Visible = False
                vsGridReceiptPanchayatFunds.Height = 6920
            Else
                fraPart1.Height = 1500
                txtSubTotal1.Visible = True
                vsGridReceiptPanchayatFunds.Height = 1095
            End If
        ElseIf cmdRFullView.Tag = 2 Then
            If cmdRFullView.Caption = "Full View" Then
                fraPart2.Top = fraPart1.Top
                fraPart1.Visible = False
                fraPart2.Height = 6920
                txtSubTotal1.Visible = False
                vsGridReceiptDebitHeads.Height = 6920
            Else
                fraPart1.Visible = True
                fraPart2.Top = fraPart1.Height
                fraPart2.Height = 1635
                txtSubTotal1.Visible = True
                vsGridReceiptDebitHeads.Height = 1275
            End If
        ElseIf cmdRFullView.Tag = 3 Then
            If cmdRFullView.Caption = "Full View" Then
                fraDE.Top = fraPart1.Top
                fraPart1.Visible = False
                fraPart2.Visible = False
                fraDE.Height = 6920
                txtSubTotal1.Visible = False
                vsDEDebt.Height = 6920
            Else
                fraPart1.Visible = True
                fraPart2.Visible = True
                fraDE.Top = 3285
                fraDE.Height = 1680
                txtSubTotal1.Visible = True
                vsDEDebt.Height = 1320
            End If
        ElseIf cmdRFullView.Tag = 4 Then
            If cmdRFullView.Caption = "Full View" Then
                fraRecoveries.Top = fraPart1.Top
                fraPart1.Visible = False
                fraPart2.Visible = False
                fraDE.Visible = False
                fraRecoveries.Height = 6920
                txtSubTotal1.Visible = False
                vsGridReceiptRecoveries.Height = 6920
            Else
                fraPart1.Visible = True
                fraPart2.Visible = True
                fraDE.Visible = True
                fraRecoveries.Top = 5445
                fraRecoveries.Height = 1725
                txtSubTotal1.Visible = True
                vsGridReceiptRecoveries.Height = 1365
            End If
        End If
    End Sub
    Private Sub cmdSave_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mArrIN  As Variant
        Dim mLoop   As Integer
        Dim mintOBRPTransactionsID As Integer
        Dim mAmount As Double
        Dim mAccID As Integer
        Dim mAccCode As Double
        Dim mSEAccID As Integer
        Dim mSEAccCode As Double
        Dim objFun      As New clsFunction
        Dim mFunID As Integer
        Dim mFunCode As Double
        Dim mSql   As String
        
        If objdb.SetConnection(mCnn) Then
       
            '------------------------PART I(Panchayat Funds)----------
            For mLoop = 1 To vsGridReceiptPanchayatFunds.Rows - 1
                If vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 0) <> "" Then
                    If vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 6) <> "" Then
                          If vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 10) = "" Or val(vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 10)) = 0 Then
                            mintOBRPTransactionsID = -1
                          Else
                            mintOBRPTransactionsID = vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 10)
                          End If
                          mSEAccID = val(vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 7))
                          mSEAccCode = vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 0)
                          
                          mAccID = val(vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 8))
                          mAccCode = vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 2)
                          
                          mFunID = vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 9)
                          objFun.SetFunctionByID (mFunID)
                          mFunCode = objFun.FunctionCode
                          
                          mAmount = val(vsGridReceiptPanchayatFunds.TextMatrix(mLoop, 6))
                          
                          mArrIN = Array(mintOBRPTransactionsID, _
                                                        mSEAccID, _
                                                        mSEAccCode, _
                                                        mAccID, _
                                                        mAccCode, _
                                                        mFunID, _
                                                        mFunCode, _
                                                        mAmount, _
                                                        10, _
                                                        0, 1, _
                                                        Null, _
                                                        Null, _
                                                        Null _
                                                        )
                       objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                    Else
                        MsgBox "Please Enter the Amount", vbInformation
                        Exit Sub
                    End If
'                Else
'                    MsgBox "Please fill the row", vbInformation
'                    Exit Sub
                End If
            Next mLoop
            '---------------------------------PART II(DEBT HEADS----------------------------------------------
            
             For mLoop = 1 To vsGridReceiptDebitHeads.Rows - 1
                 If vsGridReceiptDebitHeads.TextMatrix(mLoop, 0) <> "" Then
                    If vsGridReceiptDebitHeads.TextMatrix(mLoop, 6) <> "" Then
                          If vsGridReceiptDebitHeads.TextMatrix(mLoop, 10) = "" Or val(vsGridReceiptDebitHeads.TextMatrix(mLoop, 10)) = 0 Then
                            mintOBRPTransactionsID = -1
                          Else
                            mintOBRPTransactionsID = vsGridReceiptDebitHeads.TextMatrix(mLoop, 10)
                          End If
                          mSEAccID = val(vsGridReceiptDebitHeads.TextMatrix(mLoop, 7))
                          mSEAccCode = vsGridReceiptDebitHeads.TextMatrix(mLoop, 0)
                          
                          mAccID = vsGridReceiptDebitHeads.TextMatrix(mLoop, 8)
                          mAccCode = vsGridReceiptDebitHeads.TextMatrix(mLoop, 2)
                          
                          mFunID = vsGridReceiptDebitHeads.TextMatrix(mLoop, 9)
                          objFun.SetFunctionByID (mFunID)
                          mFunCode = objFun.FunctionCode
                          
                          mAmount = val(vsGridReceiptDebitHeads.TextMatrix(mLoop, 6))
                          
                          mArrIN = Array(mintOBRPTransactionsID, _
                                                        mSEAccID, _
                                                        mSEAccCode, _
                                                        mAccID, _
                                                        mAccCode, _
                                                        mFunID, _
                                                        mFunCode, _
                                                        mAmount, _
                                                        10, _
                                                        0, 2, _
                                                        Null, _
                                                        Null, _
                                                        Null _
                                                        )
                       objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                     Else
                        MsgBox "Please Enter the Amount", vbInformation
                        Exit Sub
                    End If
'                Else
'                    MsgBox "Please fill the row", vbInformation
'                    Exit Sub
                End If
            Next mLoop
            
            '---------------------------HEADS NOT IN SE---------------------------------------------------------
             For mLoop = 1 To vsDEDebt.Rows - 1
                 If vsDEDebt.TextMatrix(mLoop, 0) <> "" Then
                        If vsDEDebt.TextMatrix(mLoop, 4) <> "" Then
                          If vsDEDebt.TextMatrix(mLoop, 7) = "" Or val(vsDEDebt.TextMatrix(mLoop, 7)) = 0 Then
                            mintOBRPTransactionsID = -1
                          Else
                            mintOBRPTransactionsID = vsDEDebt.TextMatrix(mLoop, 7)
                          End If
                         ' mSEAccID = ""
                          'mSEAccCode = ""
    
                          mAccID = val(vsDEDebt.TextMatrix(mLoop, 5))
                          mAccCode = vsDEDebt.TextMatrix(mLoop, 0)
    
                          mFunID = vsDEDebt.TextMatrix(mLoop, 6)
                          objFun.SetFunctionByID (mFunID)
                          mFunCode = objFun.FunctionCode
    
                          mAmount = val(vsDEDebt.TextMatrix(mLoop, 4))
    
                          mArrIN = Array(mintOBRPTransactionsID, _
                                                        Null, _
                                                        Null, _
                                                        mAccID, _
                                                        mAccCode, _
                                                        mFunID, _
                                                        mFunCode, _
                                                        mAmount, _
                                                        10, _
                                                        0, 3, _
                                                        Null, _
                                                        Null, _
                                                        Null _
                                                        )
                       objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                    Else
                        MsgBox "Please Enter the Amount", vbInformation
                        Exit Sub
                    End If
'                 Else
'                        MsgBox "Please fill the row", vbInformation
'                        Exit Sub
                 End If
             Next mLoop
          '---------------------------RECOVERIES--------------------------------------------------------------
             For mLoop = 1 To vsGridReceiptRecoveries.Rows - 1
                 If vsGridReceiptRecoveries.TextMatrix(mLoop, 0) <> "" Then
                    If vsGridReceiptRecoveries.TextMatrix(mLoop, 6) <> "" Then
                          If vsGridReceiptRecoveries.TextMatrix(mLoop, 10) = "" Or val(vsGridReceiptRecoveries.TextMatrix(mLoop, 10)) = 0 Then
                            mintOBRPTransactionsID = -1
                          Else
                            mintOBRPTransactionsID = vsGridReceiptRecoveries.TextMatrix(mLoop, 10)
                          End If
                          mSEAccID = val(vsGridReceiptRecoveries.TextMatrix(mLoop, 7))
                          mSEAccCode = vsGridReceiptRecoveries.TextMatrix(mLoop, 0)
    
                          mAccID = vsGridReceiptRecoveries.TextMatrix(mLoop, 8)
                          mAccCode = vsGridReceiptRecoveries.TextMatrix(mLoop, 2)
    
                          mFunID = vsGridReceiptRecoveries.TextMatrix(mLoop, 9)
                          objFun.SetFunctionByID (mFunID)
                          mFunCode = objFun.FunctionCode
    
                          mAmount = val(vsGridReceiptRecoveries.TextMatrix(mLoop, 6))
    
                          mArrIN = Array(mintOBRPTransactionsID, _
                                                        mSEAccID, _
                                                        mSEAccCode, _
                                                        mAccID, _
                                                        mAccCode, _
                                                        mFunID, _
                                                        mFunCode, _
                                                        mAmount, _
                                                        0, _
                                                        1, 0, _
                                                        Null, _
                                                        Null, _
                                                        Null _
                                                        )
                       objdb.ExecuteSP "spSaveOBRPTransactions", mArrIN, , , mCnn, adCmdStoredProc
                    Else
                        MsgBox "Please Enter the Amount", vbInformation
                        Exit Sub
                    End If
'                Else
'                    MsgBox "Please fill the row", vbInformation
'                    Exit Sub
                End If
            Next mLoop
          '---------------------------------------------------------------------------------------------------
            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
         Else
             MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
         End If
        Call FillGridData(vsGridReceiptPanchayatFunds)
        Call FillGridData(vsGridReceiptDebitHeads)
        Call FillGridData(vsDEDebt)
        Call FillGridData(vsGridReceiptRecoveries)
         Me.Hide
         frmOBPaymentTransactions.Form_Load
         'Unload Me
         frmOpeningWizard.cmdNext_Click
       
    End Sub

    Private Sub Form_Activate()
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSave.Enabled = False
        End If
    End Sub

    Public Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        vsGridReceiptPanchayatFunds.ColComboList(0) = "|..."
        vsGridReceiptPanchayatFunds.MergeRow(0) = True
        vsGridReceiptPanchayatFunds.MergeCells = flexMergeRestrictRows
        vsDEDebt.MergeRow(0) = True
        vsDEDebt.MergeCells = flexMergeRestrictRows
        vsGridReceiptDebitHeads.ColComboList(0) = "|..."
        vsDEDebt.ColComboList(0) = "|..."
        vsGridReceiptRecoveries.ColComboList(0) = "|..."
        txtSubTotal1.Text = val(txtTotal1.Text) + val(txtTotal2.Text) + val(txtTotal3.Text)
        txtReceiptstotal.Text = val(txtSubTotal1.Text) + val(txtSubTotal2.Text)
      
        Call FillGridData(vsGridReceiptPanchayatFunds)
        Call FillGridData(vsGridReceiptDebitHeads)
        Call FillGridData(vsDEDebt)
        Call FillGridData(vsGridReceiptRecoveries)
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSave.Enabled = False
        End If
        'Call CalculateTotal
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
        If (vsGrid.Name = vsGridReceiptPanchayatFunds.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where intVoucherTypeID=10 AND tnyHeadType=1  Order By tnyHeadType"
        ElseIf (vsGrid.Name = vsGridReceiptDebitHeads.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where intVoucherTypeID=10 AND tnyHeadType=2  Order By tnyHeadType"
        ElseIf (vsGrid.Name = vsGridReceiptRecoveries.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where tnyRecovery=1  Order By tnyHeadType"
        ElseIf (vsGrid.Name = vsDEDebt.Name) Then
            mSql = ""
            mSql = "Select *,isNull(tnyHeadType,0) HeadType,isNull(tnyRecovery,0) Recovery From faOBRPTransactions Where intVoucherTypeID=10 AND tnyHeadType=3  Order By tnyHeadType"
        End If
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        vsGrid.Clear 1, 0
        vsGrid.Rows = 2
        mCnt = 1
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                If (vsGrid.Name = vsDEDebt.Name) Then
                    objAcc.SetAccountID (IIf(IsNull(Rec!intAccountHeadID), -1, Rec!intAccountHeadID))
                    mDEHead = objAcc.AccountHead
                    objFun.SetFunctionByID (IIf(IsNull(Rec!intFunctionID), -1, Rec!intFunctionID))
                    mFunction = objFun.FunctionName
                    vsGrid.TextMatrix(mCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    vsGrid.TextMatrix(mCnt, 1) = mDEHead
                    vsGrid.TextMatrix(mCnt, 2) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode) + " " + mFunction 'Fuction Code + Function
                    If IIf(IsNull(Rec!vchVoucherNo), 0, Rec!vchVoucherNo) > 0 Then
                        vsGrid.ColHidden(3) = False
                    End If
                    vsGrid.TextMatrix(mCnt, 3) = IIf(IsNull(Rec!vchVoucherNo), "", Rec!vchVoucherNo)
                    vsGrid.TextMatrix(mCnt, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
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
                    If IIf(IsNull(Rec!vchVoucherNo), 0, Rec!vchVoucherNo) > 0 And Rec!tnyRecovery = 0 Then
                        vsGrid.ColHidden(5) = False
                    End If
                    vsGrid.TextMatrix(mCnt, 5) = IIf(IsNull(Rec!vchVoucherNo), "", Rec!vchVoucherNo)
                    vsGrid.TextMatrix(mCnt, 6) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
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

    Private Sub vsDEDebt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsDEDebt.ColSel = 4 Then
           If val(vsDEDebt.TextMatrix(vsDEDebt.Row, 4)) > 0 Then
                If vsDEDebt.TextMatrix(vsDEDebt.Row, 0) = "" Or vsDEDebt.TextMatrix(vsDEDebt.Row, 2) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsDEDebt.TextMatrix(vsDEDebt.Row, 4) = ""
                    Exit Sub
                End If
            End If
        End If
    
    Call CalculateTotal
    End Sub

    Private Sub vsDEDebt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If fraDE.Height > 2000 Then
                cmdRFullView.Caption = "Orginal View"
            Else
                cmdRFullView.Caption = "Full View"
            End If
            cmdRFullView.Top = X
            cmdRFullView.Left = Y
            cmdRFullView.Visible = True
            cmdRFullView.Tag = 3
        End If
    End Sub

    Private Sub vsGridReceiptDebitHeads_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If fraPart2.Height > 2000 Then
                cmdRFullView.Caption = "Orginal View"
            Else
                cmdRFullView.Caption = "Full View"
            End If
            cmdRFullView.Top = X
            cmdRFullView.Left = Y
            cmdRFullView.Visible = True
            cmdRFullView.Tag = 2
        End If
    End Sub

    Private Sub vsGridReceiptPanchayatFunds_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If vsGridReceiptPanchayatFunds.ColSel = 6 Then
           If val(vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 6)) > 0 Then
                If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 0) = "" Or vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 4) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 6) = ""
                    Exit Sub
                End If
            End If
        End If
        Call CalculateTotal
    End Sub
    Private Sub vsGridReceiptDebitHeads_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vsGridReceiptDebitHeads.ColSel = 6 Then
           If val(vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 6)) > 0 Then
                If vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 0) = "" Or vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 4) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 6) = ""
                    Exit Sub
                End If
            End If
        End If
    Call CalculateTotal
    End Sub

    Private Sub vsGridReceiptPanchayatFunds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If fraPart1.Height > 2000 Then
                cmdRFullView.Caption = "Orginal View"
            Else
                cmdRFullView.Caption = "Full View"
            End If
            cmdRFullView.Top = X
            cmdRFullView.Left = Y
            cmdRFullView.Visible = True
            cmdRFullView.Tag = 1
        End If
    End Sub

    Private Sub vsGridReceiptRecoveries_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vsGridReceiptRecoveries.ColSel = 6 Then
           If val(vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 6)) > 0 Then
                If vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 0) = "" Or vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 4) = "" Then
                    MsgBox "Please Fill Previous Column.", vbApplicationModal
                    vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 6) = ""
                    Exit Sub
                End If
            End If
        End If
    Call CalculateTotal
    End Sub
  Private Sub vsDEDebt_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim objdb     As New clsDB
        Dim mCnn      As New ADODB.Connection
        Dim Rec       As New ADODB.Recordset
        Dim objAcc    As New clsAccounts
        Dim mSql      As String
        
        If vsDEDebt.ColSel = 0 Then
            If val(vsDEDebt.TextMatrix(vsDEDebt.Row, 5)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
            mSql = ""
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
            mSql = mSql + " Where tinType IN (1,3,4) AND intGroupID is Null And tinHiddenFlag <> 1 And "
            mSql = mSql + " vchAccountHeadCode not in (Select isNull(vchAccountHeadCode,'') From faSEAccountHeads)"
            mSql = mSql + " Order by vchAccountHeadCode"
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.Show vbModal
            If gbSearchID <> -1 Then
                If vsDEDebt.FindRow(gbSearchID, 0, 5) > 0 Then
                    MsgBox "Already selected this Account Head"
                    Exit Sub
                End If
                vsDEDebt.TextMatrix(vsDEDebt.Row, 0) = Token(gbSearchStr, " ")
                vsDEDebt.TextMatrix(vsDEDebt.Row, 1) = Trim(gbSearchStr)
                vsDEDebt.TextMatrix(vsDEDebt.Row, 5) = gbSearchID
                gbSearchID = -1
                gbSearchStr = ""
                mSql = ""
                mSql = mSql + " SELECT Distinct faFunctions.vchFunctionCode,faFunctions.vchFunction ,faFunctions.intFunctionID"
                mSql = mSql + " FROM faAccountHeads LEFT Join faTransactionTypeChild ON faTransactionTypeChild.vchAccountHeadCode=faAccountHeads.vchAccountHeadCode"
                mSql = mSql + " Inner Join faTransactionType ON faTransactionTypeChild.intTransactionTypeID=faTransactionType.intTransactionTypeID"
                mSql = mSql + " Inner Join faFunctions ON faFunctions.intFunctionID=faTransactionType.intFunctionID"
                mSql = mSql + " Where faAccountHeads.intAccountHeadID = " & val(vsDEDebt.TextMatrix(vsDEDebt.Row, 5))
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF And Rec.BOF) Then
                    vsDEDebt.TextMatrix(vsDEDebt.Row, 2) = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode) & " " & IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                    vsDEDebt.TextMatrix(vsDEDebt.Row, 6) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                End If
            End If
        End If
        If val(vsDEDebt.TextMatrix(vsDEDebt.Row, 6)) = 0 Then
            vsDEDebt.ColComboList(2) = "|..."
             If vsDEDebt.ColSel = 2 Then
               frmSearchFunction.Show vbModal
               vsDEDebt.TextMatrix(vsDEDebt.Row, 2) = Trim(gbSearchStr)
               vsDEDebt.TextMatrix(vsDEDebt.Row, 6) = gbSearchID
               gbSearchStr = ""
               gbSearchID = -1
             End If
         Else
            vsDEDebt.ColComboList(2) = ""
         End If
    End Sub
    Private Sub vsDEDebt_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim mSql   As String
        Dim objdb As New clsDB
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                If val(vsDEDebt.TextMatrix(vsDEDebt.Row, 7)) <> 0 Then
                    mSql = "Delete FROM faOBRPTransactions Where intOBRPTransactionsID= " & vsDEDebt.TextMatrix(vsDEDebt.Row, 7)
                    vsDEDebt.RemoveItem (vsDEDebt.Row)
                    mCnn.Execute (mSql)
                ElseIf vsDEDebt.Rows > 1 Then
                    vsDEDebt.RemoveItem (vsDEDebt.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If vsDEDebt.TextMatrix(vsDEDebt.Row, 4) <> "" And vsDEDebt.Row = vsDEDebt.Rows - 1 Then
               vsDEDebt.Rows = vsDEDebt.Rows + 1
            End If
        End If
        'Call CalculatePart3
        Call CalculateTotal
    End Sub
    Private Sub vsDEDebt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
         If vsDEDebt.ColSel = 4 Then
         If KeyAscii = 13 Then
            If vsDEDebt.TextMatrix(vsDEDebt.Row, 2) = "" Then
               vsDEDebt.TextMatrix(vsDEDebt.Row, 4) = ""
               MsgBox "Please select the Function", vbInformation
               
            ElseIf vsDEDebt.Row = vsDEDebt.Rows - 1 Then
                vsDEDebt.Rows = vsDEDebt.Rows + 1
                'Call CalculatePart3
                Call CalculateTotal
            End If
         End If
        End If
        If vsDEDebt.Col = 4 Then
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47) Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    End Sub
'    Private Sub CalculatePart3()
'        Dim mCnt    As Integer
'        Dim mTotal  As Double
'        Dim mSubTotal As Double
'        mTotal = 0
'        For mCnt = 1 To vsDEDebt.Rows - 1
'            mTotal = mTotal + val(vsDEDebt.TextMatrix(mCnt, 4))
'        Next
'        txtTotal3.Text = Format(mTotal, "#.00")
'        mSubTotal = val(txtTotal1.Text) + val(txtTotal2.Text) + val(txtTotal3.Text)
'        txtSubTotal1.Text = Format(mSubTotal, "#.00")
'    End Sub
    Private Sub vsGridReceiptDebitHeads_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim objdb     As New clsDB
        Dim mCnn      As New ADODB.Connection
        Dim Rec       As New ADODB.Recordset
        Dim objAcc    As New clsAccounts
        Dim mSql      As String
        
       If vsGridReceiptDebitHeads.ColSel = 0 Then
            If val(vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 7)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
             frmSearchSEPanchayatAccountHeads.intModeOfTransaction = 2
             frmSearchSEPanchayatAccountHeads.Show vbModal
             If gbSearchID <> -1 Then
                 If vsGridReceiptDebitHeads.FindRow(gbSearchID, 0, 7) > 0 Then
                     MsgBox "Already selected this Account Head"
                     Exit Sub
                 End If
                 vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 0) = gbSearchCode
                 vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 1) = gbSearchStr
                 vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 7) = gbSearchID
                 Call FillDEHeas4SE(vsGridReceiptDebitHeads)
                 gbSearchID = -1
                 gbSearchStr = ""
                 gbSearchCode = ""
            End If
       End If
       If vsGridReceiptDebitHeads.ColSel = 2 Then
            mSql = "select( faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead,faAccountHeads.intAccountHeadID From faSEAccountHeads inner join  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode where intSEHeadID=" & vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 7) & " "
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If (Rec.EOF And Rec.BOF) Then
               MsgBox " Double Entry Account Head is not Mapped", vbInformation
               Exit Sub
            End If
            If vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 0) <> "" Then
                frmSearchAccountHeads.SQLString = "select( faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead,faAccountHeads.intAccountHeadID From faSEAccountHeads inner join  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode where intSEHeadID=" & vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 7) & " "
                frmSearchAccountHeads.Show vbModal
                If gbSearchID <> -1 Then
                          If vsGridReceiptDebitHeads.FindRow(gbSearchID, 0, 7) > 0 Then
                              MsgBox "Already selected this Account Head"
                              Exit Sub
                          End If
                          vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 2) = Token(gbSearchStr, " ")
                          vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 3) = Trim(gbSearchStr)
                          vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 8) = gbSearchID
                          gbSearchID = -1
                          gbSearchStr = ""
                End If
            Else
                MsgBox "Please select the SE AccountHead, vbInformation"
            End If
       End If
       If vsGridReceiptDebitHeads.ColSel = 4 Then
        If vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 2) <> "" Then
            frmSearchFunction.Show vbModal
            vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 4) = Trim(gbSearchStr)
            vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 9) = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        Else
            MsgBox "Please select the DE AccountHead", vbInformation
        End If
       End If
    End Sub
    Private Sub vsGridReceiptDebitHeads_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim mSql   As String
        Dim objdb As New clsDB
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                If val(vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 10)) <> 0 Then
                    mSql = "Delete FROM faOBRPTransactions Where intOBRPTransactionsID= " & vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 10)
                    vsGridReceiptDebitHeads.RemoveItem (vsGridReceiptDebitHeads.Row)
                    mCnn.Execute (mSql)
                ElseIf vsGridReceiptDebitHeads.Rows > 1 Then
                   vsGridReceiptDebitHeads.RemoveItem (vsGridReceiptDebitHeads.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 6) <> "" And vsGridReceiptDebitHeads.Row = vsGridReceiptDebitHeads.Rows - 1 Then
               vsGridReceiptDebitHeads.Rows = vsGridReceiptDebitHeads.Rows + 1
            End If
        End If
        'Call CalculatePart2
        Call CalculateTotal
    End Sub
'    Private Sub CalculatePart2()
'        Dim mCnt    As Integer
'        Dim mTotal  As Double
'        Dim mSubTotal As Double
'
'        mTotal = 0
'        For mCnt = 1 To vsGridReceiptDebitHeads.Rows - 1
'            mTotal = mTotal + val(vsGridReceiptDebitHeads.TextMatrix(mCnt, 6))
'        Next
'        txtTotal2.Text = Format(mTotal, "#.00")
'        mSubTotal = val(txtTotal1.Text) + val(txtTotal2.Text) + val(txtTotal3.Text)
'        txtSubTotal1.Text = Format(mSubTotal, "#.00")
'    End Sub
    Private Sub vsGridReceiptDebitHeads_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGridReceiptDebitHeads.ColSel = 6 Then
         If KeyAscii = 13 Then
            If vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 4) = "" Then
               vsGridReceiptDebitHeads.TextMatrix(vsGridReceiptDebitHeads.Row, 6) = ""
               MsgBox "Please select the Function", vbInformation
            ElseIf vsGridReceiptDebitHeads.Row = vsGridReceiptDebitHeads.Rows - 1 Then
                vsGridReceiptDebitHeads.Rows = vsGridReceiptDebitHeads.Rows + 1
                'Call CalculatePart2
                Call CalculateTotal
            End If
         End If
        End If
        If vsGridReceiptDebitHeads.Col = 6 Then
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47) Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub vsGridReceiptPanchayatFunds_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'        If Col = 0 Or Col = 2 Then
'            If vsGridReceiptPanchayatFunds.TextMatrix(Row, 1) = "" Then
'                MsgBox "Please Select Account Head..."
'                Exit Sub
'            End If
'        End If
  
    End Sub
    Private Sub vsGridReceiptPanchayatFunds_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        
        Dim objdb     As New clsDB
        Dim mCnn      As New ADODB.Connection
        Dim Rec       As New ADODB.Recordset
        Dim objAcc    As New clsAccounts
        Dim mSql      As String
        Dim mLoop     As Integer
        
       If vsGridReceiptPanchayatFunds.ColSel = 0 Then
'             If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 7) <> "" Then
'
'             End If
             If val(vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 7)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
             End If
             frmSearchSEPanchayatAccountHeads.intModeOfTransaction = 1
             frmSearchSEPanchayatAccountHeads.Show vbModal
             If gbSearchID <> -1 Then
                 If vsGridReceiptPanchayatFunds.FindRow(gbSearchID, 0, 7) > 0 Then
                     MsgBox "Already selected this Account Head"
                     Exit Sub
                 End If
                 vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 0) = gbSearchCode
                 vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 1) = gbSearchStr
                 vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 7) = gbSearchID
                 Call FillDEHeas4SE(vsGridReceiptPanchayatFunds)
                 gbSearchID = -1
                 gbSearchStr = ""
                 gbSearchCode = ""
            End If
       End If
       If vsGridReceiptPanchayatFunds.ColSel = 2 Then
            mSql = "select( faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead,faAccountHeads.intAccountHeadID From faSEAccountHeads inner join  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode where intSEHeadID=" & vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 7) & " "
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If (Rec.EOF And Rec.BOF) Then
               MsgBox "No Double Entry Account Head is Mapped", vbInformation
               Exit Sub
            End If
            If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 0) <> "" Then
                frmSearchAccountHeads.SQLString = "select( faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead,faAccountHeads.intAccountHeadID From faSEAccountHeads inner join  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode where intSEHeadID=" & vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 7) & " "
                frmSearchAccountHeads.Show vbModal
                If gbSearchID <> -1 Then
                          If vsGridReceiptPanchayatFunds.FindRow(gbSearchID, 0, 7) > 0 Then
                              MsgBox "Already selected this Account Head"
                              Exit Sub
                          End If
                          vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 2) = Token(gbSearchStr, " ")
                          vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 3) = Trim(gbSearchStr)
                          vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 8) = gbSearchID
                          gbSearchID = -1
                          gbSearchStr = ""
                End If
            Else
                MsgBox "Please select the SE AccountHead, vbInformation"
            End If
       End If
       If vsGridReceiptPanchayatFunds.ColSel = 4 Then
        If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 2) <> "" Then
            frmSearchFunction.Show vbModal
            vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 4) = Trim(gbSearchStr)
            vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 9) = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        Else
            MsgBox "Please select the DE AccountHead", vbInformation
        End If
       End If
    End Sub
    Private Sub vsGridReceiptPanchayatFunds_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim mSql   As String
        Dim objdb As New clsDB
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
               If val(vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 10)) <> 0 Then
                    mSql = "Delete FROM faOBRPTransactions Where intOBRPTransactionsID= " & vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 10) & " "
                    mCnn.Execute (mSql)
                    vsGridReceiptPanchayatFunds.RemoveItem (vsGridReceiptPanchayatFunds.Row)
               ElseIf vsGridReceiptPanchayatFunds.Rows > 1 Then
                    vsGridReceiptPanchayatFunds.RemoveItem (vsGridReceiptPanchayatFunds.Row)
               End If
            End If
        End If
        If KeyCode = 13 Then
            If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 6) <> "" And vsGridReceiptPanchayatFunds.Row = vsGridReceiptPanchayatFunds.Rows - 1 Then
               vsGridReceiptPanchayatFunds.Rows = vsGridReceiptPanchayatFunds.Rows + 1
            End If
        End If
        'Call CalculatePart1
        Call CalculateTotal
    End Sub
    Private Sub vsGridReceiptPanchayatFunds_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGridReceiptPanchayatFunds.ColSel = 6 Then
'           If val(vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 6)) > 0 Then
'                If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 0) = "" Or vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 4) = "" Then
'                    MsgBox "Please Fill Previous Column.", vbApplicationModal
'                    vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 6) = ""
'                    Exit Sub
'                End If
'            End If
         
         If KeyAscii = 13 Then
            If vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 4) = "" Then
               vsGridReceiptPanchayatFunds.TextMatrix(vsGridReceiptPanchayatFunds.Row, 6) = ""
               MsgBox "Please select the Function", vbInformation
            ElseIf vsGridReceiptPanchayatFunds.Row = vsGridReceiptPanchayatFunds.Rows - 1 Then
                vsGridReceiptPanchayatFunds.Rows = vsGridReceiptPanchayatFunds.Rows + 1
                Call CalculateTotal
            End If
         End If
        End If
        If vsGridReceiptPanchayatFunds.Col = 6 Then
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47) Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
       
    End Sub
    Private Sub vsGridReceiptRecoveries_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
        Dim objdb     As New clsDB
        Dim mCnn      As New ADODB.Connection
        Dim Rec       As New ADODB.Recordset
        Dim objAcc    As New clsAccounts
        Dim mSql      As String
        
       If vsGridReceiptRecoveries.ColSel = 0 Then
            If val(vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 7)) <> 0 Then
                If MsgBox("Data Already Exists in this Row.. Do you want to Replace", vbYesNo, "Saankhya") = vbNo Then
                    Exit Sub
                End If
            End If
             frmSearchSEPanchayatAccountHeads.intModeOfTransaction = 5
             frmSearchSEPanchayatAccountHeads.Show vbModal
             If gbSearchID <> -1 Then
                 If vsGridReceiptRecoveries.FindRow(gbSearchID, 0, 7) > 0 Then
                     MsgBox "Already selected this Account Head"
                     Exit Sub
                 End If
                 vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 0) = gbSearchCode
                 vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 1) = gbSearchStr
                 vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 7) = gbSearchID
                 Call FillDEHeas4SE(vsGridReceiptRecoveries)
                 gbSearchID = -1
                 gbSearchStr = ""
                 gbSearchCode = ""
            End If
       End If
       If vsGridReceiptRecoveries.ColSel = 2 Then
            mSql = "select( faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead,faAccountHeads.intAccountHeadID From faSEAccountHeads inner join  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode where intSEHeadID=" & vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 7) & " "
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If (Rec.EOF And Rec.BOF) Then
               MsgBox "No Double Entry Account Head is Mapped", vbInformation
               Exit Sub
            End If
            If vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 0) <> "" Then
                frmSearchAccountHeads.SQLString = "select( faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead,faAccountHeads.intAccountHeadID From faSEAccountHeads inner join  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode where intSEHeadID=" & vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 7) & " "
                frmSearchAccountHeads.Show vbModal
                If gbSearchID <> -1 Then
                          If vsGridReceiptRecoveries.FindRow(gbSearchID, 0, 7) > 0 Then
                              MsgBox "Already selected this Account Head"
                              Exit Sub
                          End If
                          vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 2) = Token(gbSearchStr, " ")
                          vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 3) = Trim(gbSearchStr)
                          vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 8) = gbSearchID
                          gbSearchID = -1
                          gbSearchStr = ""
                End If
            Else
                MsgBox "Please select the SE AccountHead, vbInformation"
            End If
       End If
       If vsGridReceiptRecoveries.ColSel = 4 Then
        If vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 2) <> "" Then
            frmSearchFunction.Show vbModal
            vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 4) = Trim(gbSearchStr)
            vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 9) = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        Else
            MsgBox "Please select the DE AccountHead", vbInformation
        End If
       End If

    End Sub
    Private Sub vsGridReceiptRecoveries_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim mSql   As String
        Dim objdb As New clsDB
         
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                If val(vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 10)) <> 0 Then
                    mSql = "Delete FROM faOBRPTransactions Where intOBRPTransactionsID= " & vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 10)
                    vsGridReceiptRecoveries.RemoveItem (vsGridReceiptRecoveries.Row)
                    mCnn.Execute (mSql)
                ElseIf vsGridReceiptRecoveries.Rows > 1 Then
                    vsGridReceiptRecoveries.RemoveItem (vsGridReceiptRecoveries.Row)
                End If
            End If
        End If
        If KeyCode = 13 Then
            If vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 6) <> "" And vsGridReceiptRecoveries.Row = vsGridReceiptRecoveries.Rows - 1 Then
               vsGridReceiptRecoveries.Rows = vsGridReceiptRecoveries.Rows + 1
            End If
        End If
        'Call CalculateRecoveries
        Call CalculateTotal
    End Sub
    Private Sub vsGridReceiptRecoveries_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGridReceiptRecoveries.ColSel = 6 Then
         If KeyAscii = 13 Then
            If vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 4) = "" Then
                vsGridReceiptRecoveries.TextMatrix(vsGridReceiptRecoveries.Row, 6) = ""
               MsgBox "Please select the Function", vbInformation
            ElseIf vsGridReceiptRecoveries.Row = vsGridReceiptRecoveries.Rows - 1 Then
                vsGridReceiptRecoveries.Rows = vsGridReceiptRecoveries.Rows + 1
                'Call CalculateRecoveries
                Call CalculateTotal
            End If
         End If
        End If
        If vsGridReceiptRecoveries.Col = 6 Then
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47) Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    End Sub
'    Private Sub CalculateRecoveries()
'        Dim mCnt    As Integer
'        Dim mTotal  As Double
'
'
'        mTotal = 0
'        For mCnt = 1 To vsGridReceiptRecoveries.Rows - 1
'            mTotal = mTotal + val(vsGridReceiptRecoveries.TextMatrix(mCnt, 6))
'        Next
'        txtSubTotal2.Text = Format(mTotal, "#.00")
'    End Sub
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
        For mCnt = 1 To vsGridReceiptPanchayatFunds.Rows - 1
            If vsGridReceiptPanchayatFunds.TextMatrix(mCnt, 6) <> "" Then
                mPAmt = mPAmt + vsGridReceiptPanchayatFunds.TextMatrix(mCnt, 6)
            End If
        Next
        For mCnt = 1 To vsGridReceiptDebitHeads.Rows - 1
            If vsGridReceiptDebitHeads.TextMatrix(mCnt, 6) <> "" Then
                mDAmt = mDAmt + vsGridReceiptDebitHeads.TextMatrix(mCnt, 6)
            End If
        Next
        For mCnt = 1 To vsDEDebt.Rows - 1
            If vsDEDebt.TextMatrix(mCnt, 4) <> "" Then
                mDeamt = mDeamt + vsDEDebt.TextMatrix(mCnt, 4)
            End If
        Next
        For mCnt = 1 To vsGridReceiptRecoveries.Rows - 1
            If vsGridReceiptRecoveries.TextMatrix(mCnt, 6) <> "" Then
                mRecvry = mRecvry + vsGridReceiptRecoveries.TextMatrix(mCnt, 6)
            End If
        Next
        mSub1 = mPAmt + mDAmt + mDeamt
        mSub2 = mRecvry
        mTot = mSub1 + mSub2
        txtSubTotal1.Text = mSub1
        txtSubTotal2.Text = mSub2
        txtReceiptstotal.Text = mTot
    End Sub
    Private Sub FillDEHeas4SE(vsGrid As VSFlexGrid)
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objAcc      As New clsAccounts
        Dim objFun      As New clsFunction
        Dim mSql        As String
        Dim mCnt        As Integer
        mSql = "SELECT Distinct  faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead," & vbNewLine
        mSql = mSql + " faAccountHeads.intAccountHeadID , faFunctions.vchFunctionCode, faFunctions.vchFunction, faFunctions.intFunctionID" & vbNewLine
        mSql = mSql + " From faSEAccountHeads"
        'mSql = mSql + " INNER JOIN  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode" & vbNewLine
        'mSql = mSql + " LEFT Join faTransactionTypeChild ON faTransactionTypeChild.vchAccountHeadCode=faAccountHeads.vchAccountHeadCode" & vbNewLine
        'mSql = mSql + " Inner Join faTransactionType ON faTransactionTypeChild.intTransactionTypeID=faTransactionType.intTransactionTypeID" & vbNewLine
        'mSql = mSql + " Inner Join faFunctions ON faFunctions.intFunctionID=faTransactionType.intFunctionID" & vbNewLine
        mSql = mSql + " Left JOIN  faAccountHeads on faAccountHeads.vchAccountHeadCode=faSEAccountHeads.vchAccountHeadCode"
        mSql = mSql + " Left Join faFunctions ON faFunctions.vchFunctionCode=faSEAccountHeads.vchFunctionCode"
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
            vsGrid.AddItem ("")
'            vsGrid.TextMatrix(vsGrid.Row, 2) = ""
'            vsGrid.TextMatrix(vsGrid.Row, 3) = ""
'            vsGrid.TextMatrix(vsGrid.Row, 4) = ""
'            vsGrid.TextMatrix(vsGrid.Row, 8) = ""
'            vsGrid.TextMatrix(vsGrid.Row, 9) = ""
            Exit Sub
        End If
            
    End Sub

    Private Sub vsGridReceiptRecoveries_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If fraRecoveries.Height > 2000 Then
                cmdRFullView.Caption = "Orginal View"
            Else
                cmdRFullView.Caption = "Full View"
            End If
            cmdRFullView.Top = X
            cmdRFullView.Left = Y
            cmdRFullView.Visible = True
            cmdRFullView.Tag = 4
        End If
    End Sub
