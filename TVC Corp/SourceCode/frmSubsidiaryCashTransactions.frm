VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSubsidiaryCashTransactions 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subsidiary Cash Book Transactions"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   13740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7305
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   34
      Text            =   "0.00"
      Top             =   7395
      Width           =   2130
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   13740
      TabIndex        =   10
      Top             =   735
      Width           =   13740
      Begin VB.TextBox txtCashBookID 
         Height          =   315
         Left            =   13650
         TabIndex        =   46
         Top             =   165
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CheckBox Check1 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11850
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   135
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   285
         Left            =   11400
         TabIndex        =   13
         Top             =   210
         Width           =   285
      End
      Begin VB.TextBox txtSubsidiaryCashBook 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5970
         TabIndex        =   12
         Top             =   195
         Width           =   5400
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subsidiary Cash Book:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4050
         TabIndex        =   15
         Top             =   255
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1845
         TabIndex        =   14
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   13680
      TabIndex        =   9
      Top             =   7215
      Width           =   13740
      Begin VB.CheckBox optApprove 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Approve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4545
         TabIndex        =   45
         Top             =   120
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.OptionButton optApprove1 
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4740
         TabIndex        =   39
         Top             =   135
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optReject 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   135
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmdRemitBack 
         Caption         =   "Remit Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1485
      End
      Begin VB.CommandButton cmdPayment 
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5790
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1485
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7305
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1485
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   35
         Top             =   180
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13740
      TabIndex        =   8
      Top             =   0
      Width           =   13740
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1020
      Left            =   0
      TabIndex        =   16
      Top             =   1365
      Width           =   13740
      Begin VB.TextBox txtFunction 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7575
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   540
         Width           =   3795
      End
      Begin VB.TextBox txtFunctionary 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   540
         Width           =   3795
      End
      Begin VB.TextBox txtAccountHeadCode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   210
         Width           =   1560
      End
      Begin VB.TextBox txtAccountHead 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   210
         Width           =   7485
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6795
         TabIndex        =   42
         Top             =   615
         Width           =   810
      End
      Begin VB.Label lblFunctionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1290
         TabIndex        =   41
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head of Account:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   255
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Height          =   1425
      Left            =   0
      TabIndex        =   17
      Top             =   2325
      Width           =   13740
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   4
         Top             =   915
         Width           =   9045
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2325
         MaxLength       =   9
         TabIndex        =   2
         Top             =   585
         Width           =   2190
      End
      Begin VB.TextBox txtPaidTo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   0
         Top             =   255
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   8010
         TabIndex        =   22
         Top             =   255
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   300
         Left            =   11415
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtType 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8730
         TabIndex        =   1
         Top             =   255
         Width           =   2640
      End
      Begin VB.TextBox txtReference 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8730
         MaxLength       =   50
         TabIndex        =   3
         Top             =   585
         Width           =   2640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid To "
         Height          =   195
         Left            =   1710
         TabIndex        =   27
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Left            =   8325
         TabIndex        =   26
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
         Height          =   195
         Left            =   1545
         TabIndex        =   25
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   1725
         TabIndex        =   24
         Top             =   645
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref"
         Height          =   195
         Left            =   8430
         TabIndex        =   23
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3900
      Left            =   0
      TabIndex        =   28
      Top             =   3300
      Width           =   13740
      Begin VB.TextBox txtAmountReceived 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   3540
         Width           =   2190
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9255
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   3540
         Width           =   2130
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2985
         Left            =   1635
         TabIndex        =   30
         Top             =   480
         Width           =   10095
         _cx             =   17806
         _cy             =   5265
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSubsidiaryCashTransactions.frx":0000
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
         TabBehavior     =   1
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
      Begin VB.Label lblDemandNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   5805
         TabIndex        =   40
         Top             =   3570
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8745
         TabIndex        =   33
         Top             =   3600
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Received:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   795
         TabIndex        =   32
         Top             =   3600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSubsidiaryCashTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '*********************************************************************************************'
    '                       Form to make disbursement of Subsidiary Cash Book                     '
    '*********************************************************************************************'
    Dim mID             As Variant
    Dim mintTransferID     As Variant          'TransferID
    Dim mSubsidiaryAccountHeadID As Variant
    Dim mTypeID         As Variant
    Dim mDate           As Date
    Dim mUserID         As Variant
    Dim mSeatID         As Variant
    Dim mAccountHeadID  As Variant
    Dim mFunctionaryID  As Variant
    Dim mFunctionID     As Variant
    Dim mAmount         As Variant
    Dim mApprovedUserID As Variant
    Dim mApprovalDate   As Variant
    Dim mReference      As Variant
    Dim mRemarks        As Variant
    Dim mLinkID         As Variant
    Dim mStatus         As Variant
    Dim mExpenditure    As Variant
    Dim mVoucherID      As Variant
        
    Dim mintID          As Integer
    Dim mTransferID     As Integer
    Dim mAmtReceived    As Double
    Dim mDemandID       As Variant
    Dim marrOut     As Variant
            
    Private Sub FillDetails()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim mSql        As String
        Dim msqlchild   As String
        Dim Rec         As New ADODB.Recordset
        Dim RecChild    As New ADODB.Recordset
        Dim mRowCount   As Integer
        
        On Error GoTo err
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select *,faSubsidiaryCashBook.intFunctionaryID[FunctionaryID],faSubsidiaryCashBook.intFunctionID[FunctionID],faSubsidiaryCashBook.numUserID[UserID] From faSubsidiaryCashBook"
            mSql = mSql + " Left Join faSubsidiaryAccountHeads On faSubsidiaryCashBook.intSubsidiaryAccountHeadID = faSubsidiaryAccountHeads.intSubsidiaryAccountHeadID"
            mSql = mSql + " Left Join faAccountHeads On faSubsidiaryCashBook.intAccountHeadID = faAccountHeads.intAccountHeadID"
            mSql = mSql + " Left Join faFunctionaries On faSubsidiaryCashBook.intFunctionaryID = faFunctionaries.intFunctionaryID"
            mSql = mSql + " Left Join faFunctions On faSubsidiaryCashBook.intFunctionID = faFunctions.intFunctionID"
            mSql = mSql + " Left Join faUser On faSubsidiaryCashBook.numUserID= faUser.numUserID"
            mSql = mSql + " Where intID = " & intID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtSubsidiaryCashBook.Text = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode) + " " + IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
                txtSubsidiaryCashBook.Tag = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
                txtDate.Tag = IIf(IsNull(Rec!intID), "", Rec!intID)
                txtCashBookID.Text = IIf(IsNull(Rec!intTransferID), "", Rec!intTransferID)
                txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                txtFunctionary.Tag = IIf(IsNull(Rec!FunctionaryID), "", Rec!FunctionaryID)
                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                txtFunction.Tag = IIf(IsNull(Rec!FunctionID), "", Rec!FunctionID)
                txtRemarks.Tag = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName) 'UserName
                txtAmountReceived.Text = AmtReceived
                txtAmountReceived.Tag = IIf(IsNull(Rec!UserID), "", Rec!UserID) 'UserID
                
                vsGrid.Rows = 1
                mRowCount = 1

                msqlchild = "Select * From faSubsidiaryCashBookChild"
                msqlchild = msqlchild + " Where intID =" & intID
                msqlchild = msqlchild + " Order By intSerialNo"
                RecChild.Open msqlchild, mCnn
                While Not RecChild.EOF
                    vsGrid.AddItem ""
                    vsGrid.TextMatrix(mRowCount, 0) = mRowCount
                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecChild!vchPayee), "", RecChild!vchPayee)
                    vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecChild!vchReference), "", RecChild!vchReference)
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount)
                    vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecChild!intSerialNo), "", RecChild!intSerialNo)
                    vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(RecChild!vchRemarks), "", RecChild!vchRemarks)
                    txtTotal.Text = val(txtTotal.Text) + val(IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount))
                    RecChild.MoveNext
                    mRowCount = mRowCount + 1
                Wend
                txtBalance.Text = val(txtAmountReceived.Text) - val(txtTotal.Text)
                txtBalance.Tag = DemandID 'DemandID
                RecChild.Close
            End If
            Rec.Close
            intID = 0
            AmtReceived = 0
            DemandID = ""
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub GenerateAutomatedJournalForNetSalaryPayable(mCnn As ADODB.Connection)
        Dim objAcc              As New clsAccounts
        Dim objTranType         As New clsTransactionType
        Dim objDb               As New clsDB

        Dim mExtCnn             As New ADODB.Connection
        Dim mConStr             As String

        Dim Rec                 As New ADODB.Recordset
        Dim RecTranType         As New ADODB.Recordset
'        Dim mCnn As New ADODB.Connection

        Dim arrInputMaster      As Variant
        Dim arrInput            As Variant
        Dim mRows               As Long
        Dim mintByLedgerID      As Long
        Dim arrOutPut As Variant
        Dim mLoopCrl As Integer

        Dim mintFundID          As Variant
        Dim mintFunctionID      As Variant
        Dim mintFunctionaryID   As Variant
        Dim mintFieldID         As Variant
        Dim mintVoucherID       As Variant
        Dim mintTransactionID   As Variant
        Dim mTransactionTypeID  As Variant
        Dim mVoucherTypeID      As Variant
        Dim mSubLedgerID        As Variant
        Dim mintKeyID1          As Variant
        Dim mintKeyID2          As Variant
        Dim mintProcessID       As Long
        Dim mLoop               As Long
        Dim mBudgetCentreID     As Variant
        Dim mtinDebitOrCredit   As Integer
        Dim mAmount             As Double
        Dim mintOrder           As Integer
        Dim mSql                As String
        Dim mVoucherGroupID     As Integer
        Dim mRPLinkID           As Variant
        
        On Error GoTo err
'        If (objDb.CreateNewConnection(mcnn, enuSourceString.Saankhya)) Then
'            mcnn.BeginTrans
            
            'mTransactionTypeID = 1211
            Select Case val(txtAccountHeadCode.Tag)
                Case gbAcHeadIDNetSalaryPayable
                    mTransactionTypeID = gbTransactionTypePayBills
                Case gbAcHeadIDUnemploymentWages
                    mTransactionTypeID = gbTransactionTypeBFundSSSFund
            End Select
            
            mVoucherTypeID = 40
            mVoucherGroupID = 2
            mSubLedgerID = txtSubsidiaryCashBook.Tag 'intID
            mintKeyID1 = val(txtAccountHeadCode.Tag) ' AccountHeadID
            mintKeyID2 = val(txtAmountReceived.Tag) 'UserID
            mBudgetCentreID = Null
            mintProcessID = 0
            'mVoucherGroupID = 0
            mRPLinkID = ""
            'mVoucherGroupID = 2
            mRPLinkID = Null
            mintFundID = Null
            mintFunctionaryID = txtFunctionary.Tag
            mintFunctionID = txtFunction.Tag
            mintFieldID = Null
    
    
            arrInput = Array( _
                               -1, _
                                gbLocalBodyID, _
                                Null, _
                                mTransactionTypeID, _
                                mVoucherTypeID, _
                                Null, _
                                Null, _
                                gbTransactionDate, _
                                val(txtBalance.Text), _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                gbUserID, _
                                gbCounterID, _
                                Null, _
                                mintKeyID1, _
                                mintKeyID2, 115, 1, gbFinancialYearID, Null, Null, Null, Null, Null, mintFundID, gbSeatID, Null, Null, Null, Null, Null, Null, Null, mVoucherGroupID, mRPLinkID)
    
    
            '-------------------------------------------------------'
            ' Connection And Transaction Begins                     '
            '-------------------------------------------------------'
            'objDb.SetConnection mCnn
            'mCnn.BeginTrans
            'On Error GoTo ErrRollBack:
    
            objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
            If IsNumeric(arrOutPut(0, 0)) Then
                mintVoucherID = arrOutPut(0, 0)
                If mintVoucherID <> "" Then
                    mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        'txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    End If
                    Rec.Close
                End If
            Else
                GoTo err
            End If
    
            '-------------------------------------------------------'
            ' faVoucher Child
            '-------------------------------------------------------'
            'Dim mintVoucherID_1         As Double  '
            Dim mintLocalBodyID_2       As Long
            Dim mintSlNo_3              As Long
            Dim mintAccountHeadID_4     As Long
            Dim mtnyDebitOrCredit_5     As Integer
            Dim mintYearID_6            As Long
            Dim mtnyPeriodID_7          As Integer
            Dim mtnyArrearFlag_8        As Integer
            Dim mnumDemandID_9          As Variant
            Dim mfltAmount_10           As Double
    
            mCnn.Execute "Delete From faVoucherChild Where intVoucherID =" & mintVoucherID
    
            mintLocalBodyID_2 = gbLocalBodyID
            mintSlNo_3 = 1
    
            mintAccountHeadID_4 = gbAcHeadIDMiscAdvance
    '                       mtnyDebitOrCredit_5 = 0
            mintYearID_6 = gbFinancialYearID
            mtnyPeriodID_7 = 3
            mtnyArrearFlag_8 = 0
            mnumDemandID_9 = Null
            mfltAmount_10 = val(txtBalance.Text)
            mtinDebitOrCredit = 0           '  selected it sets mtinDebitCredit
                          '  0 = Credit  and 1 = Debit
            '------------------------------------------------'
                'faVoucherChild Parameters
            '------------------------------------------------'
    
    '                        @intVoucherID_1     [bigint],
    '                        @intLocalBodyID_2  [int],
    '                        @intSlNo_3     [int],
    '                        @intAccountHeadID_4    [int],
    '                        @tnyDebitOrCredit_5    [tinyint],
    '                        @intYearID_6   [int],
    '                        @tnyPeriodID_7     [tinyint],
    '                        @tnyArrearFlag_8   [tinyint],
    '                        @numDemandID_9     [numeric],
    '                        @fltAmount_10      [float] = 0
    
            arrInput = Array( _
                                mintVoucherID, _
                                mintLocalBodyID_2, _
                                mintSlNo_3, _
                                mintAccountHeadID_4, _
                                mtnyDebitOrCredit_5, _
                                mintYearID_6, _
                                mtnyPeriodID_7, _
                                mtnyArrearFlag_8, _
                                mnumDemandID_9, _
                                mfltAmount_10 _
                                )
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
            '-------------------------------------------------------'
            ' faVoucher Address
            '-------------------------------------------------------'
    
                                                                             '                Else
            '----------------------------------------------------------------'
            '         G e n e r a l - J o u r n a l  P o s t i n g           '
            '----------------------------------------------------------------'
            mtinDebitOrCredit = 0           '  selected it sets mtinDebitCredit
            '  0 = Credit  and 1 = Debit
            mintProcessID = 0                   ' Used for Automation - Recurring Process
    
            mBudgetCentreID = Null
            '-------------------------------------'
            ' Data for Transaction Table          '
            '-------------------------------------'
            arrInputMaster = Array( _
                                    -1, _
                                    gbLocalBodyID, _
                                    gbFinancialYearID, _
                                    Format(gbTransactionDate, "DD/MMM/YYYY"), _
                                    0, _
                                    0, _
                                    mintFunctionID, _
                                    mintFunctionaryID, _
                                    mintFieldID, _
                                    mintFundID, _
                                    mBudgetCentreID, _
                                    Null, _
                                    mTransactionTypeID, _
                                    mintProcessID, _
                                    "JV", _
                                    40, _
                                    Null, _
                                    mSubLedgerID, _
                                    gbUserID, _
                                    mintVoucherID, _
                                    mVoucherGroupID)
    
            objDb.ExecuteSP "spSaveTransactions", arrInputMaster, arrOutPut, , mCnn
            '----------------------------------------'
            ' Data for TransactionChild              '
            '----------------------------------------'
            If IsNumeric(arrOutPut(0, 0)) Then
                mintTransactionID = arrOutPut(0, 0)
                If mintTransactionID = "" Then
                    GoTo err
                End If
            End If
            mCnn.Execute "Delete From faTransactionChild Where intTransactionID =" & mintTransactionID
            Select Case mintKeyID1  'intAccountHeadID
'                Case 689            'Programmes/Expenditures of Transferred Functions/Schemes - Unemployment Wages
'                    mintOrder = 1
'                    mtinDebitOrCredit = 1
'                    arrInput = Array(mintTransactionID, _
'                            mintOrder, _
'                            mintKeyID1, _
'                            Format(val(txtTotal.Text), "0.00"), _
'                            mtinDebitOrCredit, _
'                            "", _
'                            "", _
'                            mintFundID _
'                            )
'                        objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
'
'                    mintOrder = 2
'                    mtinDebitOrCredit = 0
'                    arrInput = Array(mintTransactionID, _
'                            mintOrder, _
'                            1550, _
'                            Format(val(txtTotal.Text), "0.00"), _
'                            mtinDebitOrCredit, _
'                            mintKeyID1, _
'                            "", _
'                            mintFundID _
'                            )
'                        objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                'Case 1078        'Net Salary Payable
                Case gbAcHeadIDNetSalaryPayable
                    mintOrder = 1
                    mtinDebitOrCredit = 1
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            mintKeyID1, _
                            Format(val(txtBalance.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            "", _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                        
                    mintOrder = 2
                    mtinDebitOrCredit = 0
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            gbAcHeadIDUnpaidSalaries, _
                            Format(val(txtBalance.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            mintKeyID1, _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                    
    '                mintOrder = 3
    '                mtinDebitOrCredit = 0
    '                arrInput = Array(mintTransactionID, _
    '                        mintOrder, _
    '                        1079, _
    '                        Format(val(txtBalance.Text), "0.00"), _
    '                        mtinDebitOrCredit, _
    '                        mintKeyID1, _
    '                        "", _
    '                        mintFundID _
    '                        )
    '                    objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            End Select
           ' GenerateJournal = mintVoucherID
'            If mintVoucherID <> "" Then
'                frmViewVoucher.MultipleVouchers = False
'                frmViewVoucher.FormName = "frmSubsidiaryCashBook"
'                frmViewVoucher.ArrayIn = Array(CStr(mintVoucherID))
'                frmViewVoucher.Show vbModal
'            End If
'        Else
'            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
'        End If
'        mcnn.CommitTrans
        Exit Sub
err:
        'mDataSavedFlag = False
        MsgBox err.Description
    End Sub

'
'    Private Sub SaveData(arrInputMaster As Variant, mInput As Variant)
'        Dim objDb               As New clsDB
'        Dim mCnn                As ADODB.Connection
'        Dim arrOutPut           As Variant
'        Dim arrInput(7)         As Variant
'        Dim mintTransactionID   As Long
'        Dim mLoop               As Long
'        Dim mCount              As Long
'
'        objDb.SetConnection mCnn
'            'mCnn.BeginTrans
'            On Error GoTo ErrRollBack:
'            Call objDb.ExecuteSP("spSaveTransactions", arrInputMaster, arrOutPut, , mCnn)
'            If IsNumeric(arrOutPut(0, 0)) Then
'                mintTransactionID = arrOutPut(0, 0)
'            Else
'                GoTo ErrRollBack:
'            End If
'            mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mintTransactionID
'            For mLoop = 0 To ((UBound(mInput) + 1) / 8) - 1
'                arrInput(0) = mintTransactionID
'                arrInput(1) = mInput(mCount + 1)
'                arrInput(2) = mInput(mCount + 2)
'                arrInput(3) = mInput(mCount + 3)
'                arrInput(4) = mInput(mCount + 4)
'                arrInput(5) = mInput(mCount + 5)
'                arrInput(6) = mInput(mCount + 6)
'                arrInput(7) = mInput(mCount + 7)
'                mCount = mCount + 8
'                objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
'            Next mLoop
'            'mCnn.CommitTrans
'            mDataSavedFlag = True
'
'        Exit Sub
'ErrRollBack:
'        Debug.Print Error$
'        'mCnn.RollbackTrans
'        mDataSavedFlag = False
'    End Sub
    Private Function GenerateJournal(mCnn As ADODB.Connection) As Variant

        Dim objAcc              As New clsAccounts
        Dim objTranType         As New clsTransactionType
        Dim objDb               As New clsDB

        Dim mExtCnn             As New ADODB.Connection
        Dim mConStr             As String

        Dim Rec                 As New ADODB.Recordset
        Dim RecTranType         As New ADODB.Recordset
'        Dim mCnn As New ADODB.Connection

        Dim arrInputMaster      As Variant
        Dim arrInput            As Variant
        Dim mRows               As Long
        Dim mintByLedgerID      As Long
        Dim arrOutPut As Variant
        Dim mLoopCrl As Integer

        Dim mintFundID          As Variant
        Dim mintFunctionID      As Variant
        Dim mintFunctionaryID   As Variant
        Dim mintFieldID         As Variant
        Dim mintVoucherID       As Variant
        Dim mintTransactionID   As Variant
        Dim mTransactionTypeID  As Variant
        Dim mVoucherTypeID      As Variant
        Dim mSubLedgerID        As Variant
        Dim mintKeyID1          As Variant
        Dim mintKeyID2          As Variant
        Dim mintProcessID       As Long
        Dim mLoop               As Long
        Dim mBudgetCentreID     As Variant
        Dim mtinDebitOrCredit   As Integer
        Dim mAmount             As Double
        Dim mintOrder           As Integer
        Dim mSql                As String
        Dim mVoucherGroupID     As Integer
        Dim mRPLinkID           As Variant
        
        On Error GoTo err
'        If (objDb.CreateNewConnection(mcnn, enuSourceString.Saankhya)) Then
'            mcnn.BeginTrans
            Select Case val(txtAccountHeadCode.Tag)
                Case gbAcHeadIDNetSalaryPayable
                    mTransactionTypeID = gbTransactionTypePayBills
                Case gbAcHeadIDUnemploymentWages
                    mTransactionTypeID = gbTransactionTypeBFundSSSFund
            End Select
            
            mTransactionTypeID = 1211
            mVoucherTypeID = 40
            mVoucherGroupID = 2
            mSubLedgerID = txtSubsidiaryCashBook.Tag 'intID
            mintKeyID1 = val(txtAccountHeadCode.Tag) ' AccountHeadID
            mintKeyID2 = val(txtAmountReceived.Tag) 'UserID
            mBudgetCentreID = Null
            mintProcessID = 0
            'mVoucherGroupID = 0
            mRPLinkID = ""
            'mVoucherGroupID = 2
            mRPLinkID = Null
            mintFundID = Null
            mintFunctionaryID = txtFunctionary.Tag
            mintFunctionID = txtFunction.Tag
            mintFieldID = Null
    
    
            arrInput = Array( _
                               -1, _
                                gbLocalBodyID, _
                                Null, _
                                mTransactionTypeID, _
                                mVoucherTypeID, _
                                Null, _
                                Null, _
                                gbTransactionDate, _
                                val(txtTotal.Text), _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                gbUserID, _
                                gbCounterID, _
                                Null, _
                                mintKeyID1, _
                                mintKeyID2, 115, 1, gbFinancialYearID, Null, Null, Null, Null, Null, mintFundID, gbSeatID, Null, Null, Null, Null, Null, Null, Null, mVoucherGroupID, mRPLinkID)
    
    
            '-------------------------------------------------------'
            ' Connection And Transaction Begins                     '
            '-------------------------------------------------------'
            'objDb.SetConnection mCnn
            'mCnn.BeginTrans
            'On Error GoTo ErrRollBack:
    
            objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
            If IsNumeric(arrOutPut(0, 0)) Then
                mintVoucherID = arrOutPut(0, 0)
                If mintVoucherID <> "" Then
                    mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        'txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    End If
                    Rec.Close
                End If
            Else
                GoTo err
            End If
    
            '-------------------------------------------------------'
            ' faVoucher Child
            '-------------------------------------------------------'
            'Dim mintVoucherID_1         As Double  '
            Dim mintLocalBodyID_2       As Long
            Dim mintSlNo_3              As Long
            Dim mintAccountHeadID_4     As Long
            Dim mtnyDebitOrCredit_5     As Integer
            Dim mintYearID_6            As Long
            Dim mtnyPeriodID_7          As Integer
            Dim mtnyArrearFlag_8        As Integer
            Dim mnumDemandID_9          As Variant
            Dim mfltAmount_10           As Double
    
            mCnn.Execute "Delete From faVoucherChild Where intVoucherID =" & mintVoucherID
    
            mintLocalBodyID_2 = gbLocalBodyID
            mintSlNo_3 = 1
    
            mintAccountHeadID_4 = gbAcHeadIDMiscAdvance
    '                       mtnyDebitOrCredit_5 = 0
            mintYearID_6 = gbFinancialYearID
            mtnyPeriodID_7 = 3
            mtnyArrearFlag_8 = 0
            mnumDemandID_9 = Null
            mfltAmount_10 = val(txtTotal.Text)
            mtinDebitOrCredit = 0           '  selected it sets mtinDebitCredit
                          '  0 = Credit  and 1 = Debit
            '------------------------------------------------'
                'faVoucherChild Parameters
            '------------------------------------------------'
    
    '                        @intVoucherID_1     [bigint],
    '                        @intLocalBodyID_2  [int],
    '                        @intSlNo_3     [int],
    '                        @intAccountHeadID_4    [int],
    '                        @tnyDebitOrCredit_5    [tinyint],
    '                        @intYearID_6   [int],
    '                        @tnyPeriodID_7     [tinyint],
    '                        @tnyArrearFlag_8   [tinyint],
    '                        @numDemandID_9     [numeric],
    '                        @fltAmount_10      [float] = 0
    
            arrInput = Array( _
                                mintVoucherID, _
                                mintLocalBodyID_2, _
                                mintSlNo_3, _
                                mintAccountHeadID_4, _
                                mtnyDebitOrCredit_5, _
                                mintYearID_6, _
                                mtnyPeriodID_7, _
                                mtnyArrearFlag_8, _
                                mnumDemandID_9, _
                                mfltAmount_10 _
                                )
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
            '-------------------------------------------------------'
            ' faVoucher Address
            '-------------------------------------------------------'
    
                                                                             '                Else
            '----------------------------------------------------------------'
            '         G e n e r a l - J o u r n a l  P o s t i n g           '
            '----------------------------------------------------------------'
            mtinDebitOrCredit = 0           '  selected it sets mtinDebitCredit
            '  0 = Credit  and 1 = Debit
            mintProcessID = 0                   ' Used for Automation - Recurring Process
    
            mBudgetCentreID = Null
            '-------------------------------------'
            ' Data for Transaction Table          '
            '-------------------------------------'
            arrInputMaster = Array( _
                                    -1, _
                                    gbLocalBodyID, _
                                    gbFinancialYearID, _
                                    Format(gbTransactionDate, "DD/MMM/YYYY"), _
                                    0, _
                                    0, _
                                    mintFunctionID, _
                                    mintFunctionaryID, _
                                    mintFieldID, _
                                    mintFundID, _
                                    mBudgetCentreID, _
                                    Null, _
                                    mTransactionTypeID, _
                                    mintProcessID, _
                                    "JV", _
                                    40, _
                                    Null, _
                                    mSubLedgerID, _
                                    gbUserID, _
                                    mintVoucherID, _
                                    mVoucherGroupID)
    
            objDb.ExecuteSP "spSaveTransactions", arrInputMaster, arrOutPut, , mCnn
            '----------------------------------------'
            ' Data for TransactionChild              '
            '----------------------------------------'
            If IsNumeric(arrOutPut(0, 0)) Then
                mintTransactionID = arrOutPut(0, 0)
                If mintTransactionID = "" Then
                    GoTo err
                End If
            End If
            mCnn.Execute "Delete From faTransactionChild Where intTransactionID =" & mintTransactionID
            Select Case mintKeyID1  'intAccountHeadID
'                Case 689            'Programmes/Expenditures of Transferred Functions/Schemes - Unemployment Wages
                Case gbAcHeadIDUnemploymentWages             'Programmes/Expenditures of Transferred Functions/Schemes - Unemployment Wages
                    mintOrder = 1
                    mtinDebitOrCredit = 1
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            mintKeyID1, _
                            Format(val(txtTotal.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            "", _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                        
                    mintOrder = 2
                    mtinDebitOrCredit = 0
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            gbAcHeadIDMiscAdvance, _
                            Format(val(txtTotal.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            mintKeyID1, _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                'Case 1078        'Net Salary Payable
                Case gbAcHeadIDNetSalaryPayable         'Net Salary Payable
                    mintOrder = 1
                    mtinDebitOrCredit = 1
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            mintKeyID1, _
                            Format(val(txtTotal.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            "", _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                        
                    mintOrder = 2
                    mtinDebitOrCredit = 0
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            gbAcHeadIDMiscAdvance, _
                            Format(val(txtTotal.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            mintKeyID1, _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                    If val(txtBalance.Text) > 0 Then
                        Call GenerateAutomatedJournalForNetSalaryPayable(mCnn)
                    End If
                Case Else
                    mintOrder = 1
                    mtinDebitOrCredit = 1
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            mintKeyID1, _
                            Format(val(txtTotal.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            "", _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn

                    mintOrder = 2
                    mtinDebitOrCredit = 0
                    arrInput = Array(mintTransactionID, _
                            mintOrder, _
                            gbAcHeadIDMiscAdvance, _
                            Format(val(txtTotal.Text), "0.00"), _
                            mtinDebitOrCredit, _
                            mintKeyID1, _
                            "", _
                            mintFundID _
                            )
                        objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
    '                mintOrder = 3
    '                mtinDebitOrCredit = 0
    '                arrInput = Array(mintTransactionID, _
    '                        mintOrder, _
    '                        1079, _
    '                        Format(val(txtBalance.Text), "0.00"), _
    '                        mtinDebitOrCredit, _
    '                        mintKeyID1, _
    '                        "", _
    '                        mintFundID _
    '                        )
    '                    objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            End Select
            GenerateJournal = mintVoucherID
'            If mintVoucherID <> "" Then
'                frmViewVoucher.MultipleVouchers = False
'                frmViewVoucher.FormName = "frmSubsidiaryCashBook"
'                frmViewVoucher.ArrayIn = Array(CStr(mintVoucherID))
'                frmViewVoucher.Show vbModal
'            End If
'        Else
'            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
'        End If
'        mcnn.CommitTrans
        Exit Function
err:
        'mDataSavedFlag = False
        MsgBox err.Description
'        mcnn.RollbackTrans
    End Function
    Private Sub CheckDemand()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objDb   As New clsDB
        
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
        
            If txtBalance.Tag <> "" Then
                mSql = "Select vchDemandNo From faIDemandTBL"
                mSql = mSql + " Where numDemandID =" & txtBalance.Tag
                mSql = mSql + " And tnyStatus <> 9"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    lblDemandNo.Visible = True
                    lblDemandNo.Caption = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
                    cmdSave.Enabled = False
                End If
                Rec.Close
            End If
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
    End Sub
    
    Private Sub GenerateDemand(mCnn As ADODB.Connection)
        Dim intLBID             As Variant
        Dim tnyExtAppID         As Variant
        Dim tnyExtModuleID      As Variant
        Dim tnyDemandType       As Variant
        Dim intTransactionTypeID As Variant
        Dim intYearID           As Variant
        Dim tnyPeriodID         As Variant
        Dim dtDemandDate        As Variant
        Dim numSubLedgerID      As Variant
        Dim intKeyID            As Variant
        Dim intKeyID2           As Variant
        Dim vchRemarks          As Variant
        Dim tnyStatus           As Variant
        Dim intVoucherID        As Variant
        Dim dtVoucherDate       As Variant
        Dim tnyArrearFlag       As Variant
        Dim dtExpiryDate        As Variant
        Dim numDemandID         As Variant
        Dim vchDemandNo         As Variant
        Dim intSectionID        As Variant
        Dim numSeatID           As Variant
        Dim vchAdminNote        As Variant
        Dim intAccountHeadID    As Variant
        Dim vchAccountHeadCode  As Variant
        Dim fltAmount           As Variant
        Dim dtOnDate            As Variant
        Dim numZoneID           As Variant
        Dim intWardNo           As Variant
        Dim intDoorNo           As Variant
        Dim vchDoorNo2          As Variant
        Dim numForwardedSeatID  As Variant
        Dim ForwardedSeatID    As String
        Dim dtDueDate           As Variant

        Dim intInstrumentTypeID As Variant
        Dim vchInstrumentNo     As Variant
        Dim dtInstrumentDate    As Variant
        Dim vchDrawnFrom        As Variant
        Dim vchDrawnPlace       As Variant
        Dim tnyAccrualType      As Variant
        
        Dim intFunctionaryID    As Variant
        Dim intFunctionID       As Variant

        Dim vchName_6           As String
        Dim vchInit1_7          As String
        Dim vchInit2_8          As String
        Dim vchInit3_9          As String
        Dim vchInit4_10         As String
        Dim vchHouseName_11     As String
        Dim vchStreet_12        As String
        Dim vchLocalPlace_13    As String
        Dim vchMainPlace_14     As String
        Dim vchPost_15          As String
        Dim vchPin_16           As String
        Dim vchPhone_17         As String

        Dim mLoop               As Variant
        Dim arrInput            As Variant
        Dim arrOutPut           As Variant

        Dim objDb               As New clsDB
'        Dim mCnn                As New ADODB.Connection

        Dim dtTransactionDate As Variant
        Dim intDemandMode As Variant

        intSectionID = Null
        intTransactionTypeID = 1211
        numForwardedSeatID = Null


        '---------------------------------------------------------'
        '**         Input Variable to spSaveIDemandTBL          **'
        '---------------------------------------------------------'
        '
        '@intLBID            int,
        '@tnyExtAppID        As Integer
        '@tnyExtModuleID     As Integer
        '@tnyDemandType      As Integer
        '@intTransactionTypeID   smallint,
        '@intYearID          smallint,
        '@tnyPeriodID        tinyint,
        '@dtDemandDate       smalldatetime,
        '@numSubLedgerID     Numeric,
        '@intKeyID           Numeric,
        '@intKeyID2          Numeric,
        '@vchRemarks         varChar(100),
        '@tnyStatus          TinyInt,
        '@intVoucherID       Int ,
        '@dtVoucherDate      SmallDateTime,
        '@tnyArrearFlag      TinyInt = 0,
        '@dtExpiryDate       SmallDateTime = Null,
        '@numDemandID        Numeric = Null Output
        '
        '@numSeatID          Numeric     = Null,
        '@intSectionID       Int         = Null,
        '@numUserID          Numeric     = Null,
        '@numCounterID       Numeric     = Null,
        '@vchAdminNote       varChar(100)    = Null,
        '@vchDemandNo        varChar(20)     = Null,
        '@numZoneID          Numeric     = Null,
        '@intWardNo          Int         = Null,
        '@intDoorNo         int         = Null,
        '@vchDoorNo2        varChar(10)     = Null
        '@intForwaredSeat   Numeric         = Null
        '@dtDueDate         SmallDateTime = Null

        intLBID = gbLocalBodyID
        tnyExtAppID = AppID.Saankhya
        tnyExtModuleID = 99
        tnyDemandType = 10
        dtDemandDate = gbTransactionDate
        numSubLedgerID = txtSubsidiaryCashBook.Tag 'intID in faSubsidiaryCashBook
        intKeyID = 1504
        intKeyID2 = txtAmountReceived.Tag 'UserID
        vchRemarks = Null
        tnyStatus = 0
        intVoucherID = Null
        dtVoucherDate = Null
        tnyArrearFlag = Null
        dtExpiryDate = gbTransactionDate
        numSeatID = gbSeatID
        'intSectionID = gbSectionID
        dtDueDate = gbTransactionDate

        vchAdminNote = Null
        vchDemandNo = Null
        numZoneID = Null
        intWardNo = Null
        intDoorNo = Null
        vchDoorNo2 = Null

        intInstrumentTypeID = gbInstrumentCash
        vchInstrumentNo = Null
        dtInstrumentDate = Null
        vchDrawnFrom = Null
        vchDrawnPlace = Null
        
        intFunctionaryID = txtFunctionary.Tag
        intFunctionID = txtFunction.Tag
        
        If txtBalance.Tag <> "" Then
            numDemandID = txtBalance.Tag
        Else
            numDemandID = ""
        End If
        
        dtTransactionDate = gbTransactionDate   ' Added On 19.10.11 By Poornima
        intDemandMode = 0

        arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
        tnyDemandType, intTransactionTypeID, _
        intYearID, tnyPeriodID, dtDemandDate, _
        numSubLedgerID, _
        intKeyID, _
        intKeyID2, _
        vchRemarks, _
        tnyStatus, _
        intVoucherID, _
        dtVoucherDate, _
        tnyArrearFlag, _
        dtExpiryDate, _
        IIf(numDemandID = "", Null, val(numDemandID)), _
        gbFinancialYearID, _
        numSeatID, _
        intSectionID, _
        gbUserID, _
        gbCounterID, _
        vchAdminNote, _
        vchDemandNo, _
        numZoneID, _
        intWardNo, _
        intDoorNo, _
        vchDoorNo2, _
        numForwardedSeatID, dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, vchDrawnFrom, vchDrawnPlace, Null, gbLocationID, intFunctionaryID, intFunctionID, Null, dtTransactionDate, intDemandMode)


'        If Not objDb.CreateNewConnection(mcnn, enuSourceString.Saankhya) Then
'            MsgBox "Didn't able to establish a connection with Database server!", vbInformation
'            Exit Sub
'        End If
'        mcnn.BeginTrans
        'On Error GoTo ErrRollBack
        objDb.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc

        If IsArray(arrOutPut) Then
            numDemandID = arrOutPut(0, 0)
            txtBalance.Tag = numDemandID
            vchDemandNo = arrOutPut(1, 0)
            lblDemandNo.Visible = True
            lblDemandNo.Caption = vchDemandNo
        Else
            MsgBox "Didn't able to Generate Demand ID!", vbInformation
            GoTo ErrRollBack:
        End If

        '---------------------------------------------------------'
        '**         Input Variable to spSaveIDemandChild        **'
        '---------------------------------------------------------'
        '@numDemandID_1      [numeric]   ,
        '@intLBID_2      [int]       ,
        '@tnySlNo_3      [tinyint]   ,
        '@intAccountHeadID_4     [int]       ,
        '@vchAccountHeadCode_5   [varchar](20)   ,
        '@fltAmount_6        [float]     ,
        '@vchRemarks_7       [varchar](100)  ,
        '@tnyStatus_8        [tinyint]   = 0 ,
        '@dtOnDate_9     [SmallDateTime] = Null
        '
        '@intYearID      [Integer] = Null,
        '@tnyPeriodID        [TinyInt] = Null,
        '@tnyArrearFlag      [TinyInt] = Null
        '---------------------------------------------------------'

        mCnn.Execute "Delete From faIDemandChild Where numDemandID=" & numDemandID

        mLoop = 1
        intAccountHeadID = 1550
        vchAccountHeadCode = 460100700
        tnyArrearFlag = Null
        fltAmount = val(txtBalance.Text)
        intYearID = Null
        tnyPeriodID = Null
        tnyStatus = 0
        dtOnDate = gbTransactionDate
        vchRemarks = ""

        arrInput = Array(numDemandID, _
        intLBID, _
        mLoop, _
        intAccountHeadID, _
        vchAccountHeadCode, _
        fltAmount, _
        vchRemarks, _
        tnyStatus, _
        dtOnDate, _
        intYearID, _
        tnyPeriodID, _
        tnyArrearFlag _
        )
        objDb.ExecuteSP "spSaveIDemandChild", arrInput, , , mCnn, adCmdStoredProc

        '---------------------------------------------------------'
        '**         Input Variable to spSaveIDemandChild        **'
        '---------------------------------------------------------'

        mCnn.Execute "Delete From faIDemandAddress Where numDemandID=" & numDemandID

        vchName_6 = Trim(txtRemarks.Tag)
        vchInit1_7 = ""
        vchInit2_8 = ""
        vchInit3_9 = ""
        vchInit4_10 = ""
        vchHouseName_11 = ""
        vchStreet_12 = ""
        vchLocalPlace_13 = ""
        vchMainPlace_14 = ""
        vchPost_15 = ""
        vchPin_16 = ""
        vchPhone_17 = ""

        arrInput = Array(numDemandID, _
        gbLocalBodyID, _
        numZoneID, _
        intWardNo, _
        intDoorNo, _
        vchDoorNo2, _
        vchName_6, _
        vchInit1_7, vchInit2_8, vchInit3_9, vchInit4_10, _
        vchHouseName_11, _
        vchStreet_12, _
        vchLocalPlace_13, _
        vchMainPlace_14, _
        vchPost_15, _
        vchPin_16, _
        vchPhone_17)

        objDb.ExecuteSP "spSaveIDemandAddress", arrInput, , , mCnn, adCmdStoredProc
        'On Error GoTo 0
        'mNewFlag = False
        'cmdSave.Enabled = False

        'If intTransactionTypeID = gbTransactionTypeProfTaxTradeAccrual Then
         '   Call AccrualJournalByDemandID(numDemandID)
        'End If

        'If chkSkipPrinting.Value = 0 Then
        Call PrintDemandSlip(numDemandID, mCnn)
        'End If
'        mcnn.CommitTrans
        Exit Sub
ErrRollBack:
        MsgBox "Unexpected Error! RollBacking!", vbInformation
'        mcnn.RollbackTrans

    End Sub
      
    Private Sub CancelDemand()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objDb   As New clsDB
        
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Update faIDemandTbl"
            mSql = mSql + " Set tnyStatus =" & 9
            mSql = mSql + " Where numDemandID = " & txtBalance.Tag
            mCnn.Execute mSql
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
    End Sub
    
    Private Sub FormInitiliaze()
        txtPaidTo.Text = ""
        txtPaidTo.Tag = ""
        txtType.Text = ""
        txtAmount.Text = ""
        txtAmount.Tag = ""
        txtReference.Text = ""
        txtRemarks.Text = ""
        marrOut = ""
        mID = ""
    End Sub
    
    Private Sub FillvsGrid()
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mRowCount   As Double
        
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mRowCount = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            txtTotal.Text = ""
            mSql = "Select * From faSubsidiaryCashBookChild"
            mSql = mSql + " Where intID =" & marrOut(0, 0)
            mSql = mSql + " Order By intSerialNo"
            Rec.Open mSql, mCnn
            While Not Rec.EOF
                vsGrid.AddItem ""
                vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intSerialNo), "", Rec!intSerialNo)
                txtPaidTo.Tag = IIf(IsNull(Rec!intSerialNo), "", Rec!intSerialNo)
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchPayee), "", Rec!vchPayee)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchReference), "", Rec!vchReference)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!intSerialNo), "", Rec!intSerialNo)
                vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                txtTotal.Text = val(txtTotal.Text) + val(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount))
                txtBalance.Text = val(txtAmountReceived.Text) - val(txtTotal.Text)
                Rec.MoveNext
                mRowCount = mRowCount + 1
            Wend
            Rec.Close
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
    End Sub
    
    Private Sub GetCashBookDetails(mLocalTypeID As Variant)
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mMAXID  As Variant
        Dim mSql    As String
        
        '*********************************************************************************************'
        '            Procedure to get the Cash book details for making Payment & Remit Back           '
        '*********************************************************************************************'
        On Error GoTo err
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select intID From faSubsidiaryCashBook"
            mSql = mSql + " Where intTransferID = " & txtCashBookID.Text
            mSql = mSql + " And intTypeID = " & mLocalTypeID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mMAXID = IIf(IsNull(Rec!intID), "", Rec!intID)
'                mMAXID = mMAXID + 1
            End If
            Rec.Close
            
            mSql = "Select * From faSubsidiaryCashBook"
            mSql = mSql + " Where intTransferID = " & txtCashBookID.Text
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If mMAXID = "" Then
                    mID = -1
                Else
                    mID = mMAXID
                End If
                mintTransferID = txtCashBookID.Text  'TransferID
                mSubsidiaryAccountHeadID = txtSubsidiaryCashBook.Tag
                mTypeID = mLocalTypeID
                mDate = gbTransactionDate
                mUserID = gbUserID
                mSeatID = gbSeatID
                mAccountHeadID = val(txtAccountHeadCode.Tag)
                mFunctionaryID = val(txtFunctionary.Tag)
                mFunctionID = val(txtFunction.Tag)
                If mLocalTypeID <> 10 Then
                    If txtTotal.Text <> "0.00" Then
                        mAmount = (val(txtTotal.Text) - val(txtAmount.Tag)) + val(txtAmount.Text)
                        'mAmount = Val(txtTotal.Text) + Val(txtAmount.Text)
                    Else
                        mAmount = val(txtAmount.Text)
                    End If
                Else
                    mAmount = val(txtAmountReceived.Text) - val(txtTotal.Text)
                End If
                mApprovedUserID = IIf(IsNull(Rec!numApprovedUserID), "", Rec!numApprovedUserID)
                mApprovalDate = IIf(IsNull(Rec!dtApprovalDate), "", Rec!dtApprovalDate)
                mReference = IIf(IsNull(Rec!vchReference), "", Rec!vchReference)
                mRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                mStatus = 0
                mExpenditure = IIf(IsNull(Rec!tnyIsExpRecorded), "", Rec!tnyIsExpRecorded)
                mVoucherID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            End If
            Rec.Close
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdSave_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim mStatus As Integer
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        Dim mID     As Variant
        Dim mintVoucherID As Variant
        
        '*********************************************************************************************'
        '                   Procedure to make approval of Remit Back                                  '
        '*********************************************************************************************'
        On Error GoTo err
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
        mCnn.BeginTrans
        'If gbUserTypeID = 2 Or gbUserTypeID = 1 Or gbUserTypeID = 4 Then
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
            mSql = "Update faSubsidiaryCashBook"
            mSql = mSql + " Set tnyStatus = 4"
            mSql = mSql + " Where intTransferID = " & txtCashBookID.Text
            mSql = mSql + " And intTypeID = 50"
            mCnn.Execute mSql
            mSql = "Update faSubsidiaryCashBook"
            mSql = mSql + " Set tnyStatus = 1,"
            mSql = mSql + " numApprovedUserID =" & gbUserID
            mSql = mSql + " ,dtApprovalDate = '" & Format(gbTransactionDate, "DD/MMM/YYYY") & "'"
            mSql = mSql + " Where intTransferID = " & txtCashBookID.Text
            mSql = mSql + " And intTypeID = 10"
            mCnn.Execute mSql
            If optApprove.value = vbChecked Then
                mStatus = 1
                If lblDemandNo.Caption = "" Then
                    txtBalance.Tag = ""
                    If Trim(txtBalance.Text) <> 0 Then 'If there is balance amount after disbursement
                        Call GenerateDemand(mCnn)
                    End If
                    mintVoucherID = GenerateJournal(mCnn)
                End If
            End If
        End If
        cmdSave.Enabled = False
        mCnn.CommitTrans
            
            '---------------------------------------'
            '-----For Generate the Journal Voucher--'
            '---------------------------------------'
            If mintVoucherID <> "" Then
                frmViewVoucher.MultipleVouchers = False
                frmViewVoucher.FormName = "frmSubsidiaryCashBook"
                frmViewVoucher.ArrayIn = Array(CStr(mintVoucherID))
                frmViewVoucher.Show vbModal
            End If
            '---------------------------------------'
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
        mCnn.RollbackTrans
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdPayment_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim mArrIn      As Variant
        Dim mArrInChild As Variant
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mRowCount   As Double
        Dim mSerialNo   As Variant
        
'        Dim mID             As Variant
'        Dim mCashBookID     As Variant
'        Dim mTypeID         As Variant
'        Dim mDate           As Date
'        Dim mUserID         As Variant
'        Dim mSeatID         As Variant
'        Dim mAccountHeadID  As Variant
'        Dim mFunctionaryID  As Variant
'        Dim mFunctionID     As Variant
'        Dim mAmount         As Variant
'        Dim mApprovedUserID As Variant
'        Dim mReference      As Variant
'        Dim mRemarks        As Variant
'        Dim mLinkID         As Variant
'        Dim mStatus         As Variant
        
        '*********************************************************************************************'
        '                           Procedure to make the Disbursement                                '
        '*********************************************************************************************'
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
        
            ''''''''''''''''''''''''''''''General Validations'''''''''''''''''''
            If Trim(txtPaidTo.Text) = "" Then
                MsgBox "Please enter the Name of Person", vbInformation
                txtPaidTo.SetFocus
                Exit Sub
            End If
            If Trim(txtAmount.Text) = "" Then
                MsgBox "Please enter the Amount", vbInformation
                txtAmount.SetFocus
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ''''''''''''''''''''''''''''''Validation of Amount''''''''''''''''''''
            If txtPaidTo.Tag = "" Then
                If val(txtAmountReceived.Text) < (val(txtTotal.Text) + val(txtAmount.Text)) Then
                    MsgBox "Amount Exceed", vbInformation
                    Exit Sub
                End If
            Else
                If val(txtAmountReceived.Text) < (val(txtTotal.Text) - val(txtAmount.Tag)) + val(txtAmount.Text) Then
                    MsgBox "Amount Exceed", vbInformation
                    Exit Sub
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Call GetCashBookDetails(20)
            mArrIn = Array(mID, _
                            mintTransferID, _
                            mSubsidiaryAccountHeadID, _
                            mTypeID, _
                            mDate, _
                            mUserID, _
                            mSeatID, _
                            mAccountHeadID, _
                            mFunctionaryID, _
                            mFunctionID, _
                            mAmount, _
                            mApprovedUserID, _
                            mApprovalDate, _
                            mReference, _
                            mRemarks, _
                            mStatus, _
                            mExpenditure, _
                            mVoucherID _
                        )
            objDb.ExecuteSP "spSaveSubsidiaryCashBook", mArrIn, marrOut, , mCnn, adCmdStoredProc
            
            If txtPaidTo.Tag <> "" Then
                mSerialNo = val(txtPaidTo.Tag)
            Else
                mSql = "Select Max(intSerialNo) As SerialNo From faSubsidiaryCashBookChild"
                mSql = mSql + " Where intID = " & marrOut(0, 0)
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mSerialNo = IIf(IsNull(Rec!SerialNo), 0, Rec!SerialNo)
                    mSerialNo = mSerialNo + 1
                End If
                Rec.Close
            End If
            
            mArrInChild = Array(marrOut(0, 0), _
                                    mSerialNo, _
                                    Null, _
                                    txtPaidTo.Text, _
                                    txtReference.Text, _
                                    val(txtAmount.Text), _
                                    txtRemarks.Text _
                                )
            objDb.ExecuteSP "spSaveSubsidiaryCashBookChild", mArrInChild, , , mCnn, adCmdStoredProc
            mSql = "Update faSubsidiaryCashBook"
            mSql = mSql + " Set tnyStatus = 2"
            mSql = mSql + " Where intTransferID =" & txtCashBookID.Text
            mSql = mSql + " And intTypeID = 50"
            mCnn.Execute mSql
            MsgBox "Successfully Saved", vbInformation
            Call FillvsGrid
            Call FormInitiliaze
    '        For mRowCount = 1 To vsGrid.Rows - 1
    '            If vsGrid.TextMatrix(mRowCount, 1) <> "" And vsGrid.TextMatrix(mRowCount, 3) <> "" Then
    '                mArrInChild = Array(mArrOut(0, 0), _
    '                                vsGrid.TextMatrix(mRowCount, 5), _
    '                                Null, _
    '                                vsGrid.TextMatrix(mRowCount, 1), _
    '                                vsGrid.TextMatrix(mRowCount, 2), _
    '                                vsGrid.TextMatrix(mRowCount, 3) _
    '                            )
    '                objDB.ExecuteSP "spSaveSubsidiaryCashBookChild", mArrInChild, , , mCnn, adCmdStoredProc
    '            Else
    '                MsgBox "Please fill the details properly", vbInformation
    '                Exit Sub
    '            End If
    '        Next
            'mCnn.Execute "Update faSubsidiaryCashBook Set intTypeID = 20 Where intID = " & txtSubsidiaryCashBook.Tag
    '        cmdPayment.Enabled = False
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
    End Sub

    Private Sub cmdRemitBack_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mArrIn  As Variant
        Dim mSql    As String
        
        '*********************************************************************************************'
        '              Procedure to make the Remit back process after disbursement                    '
        '*********************************************************************************************'
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
        
            Call GetCashBookDetails(10)
            mArrIn = Array(mID, _
                            mintTransferID, _
                            mSubsidiaryAccountHeadID, _
                            mTypeID, _
                            mDate, _
                            mUserID, _
                            mSeatID, _
                            mAccountHeadID, _
                            mFunctionaryID, _
                            mFunctionID, _
                            mAmount, _
                            mApprovedUserID, _
                            mApprovalDate, _
                            mReference, _
                            mRemarks, _
                            mStatus, _
                            mExpenditure, _
                            mVoucherID _
                        )
            objDb.ExecuteSP "spSaveSubsidiaryCashBook", mArrIn, marrOut, , mCnn, adCmdStoredProc
            mSql = "Update faSubsidiaryCashBook"
            mSql = mSql + " Set tnyStatus = 3"
            mSql = mSql + " Where intTransferID =" & txtCashBookID.Text
            mSql = mSql + " And intTypeID = 50"
            mCnn.Execute mSql
            MsgBox "Successfully Saved", vbInformation
            cmdRemitBack.Enabled = False
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
    End Sub

    Private Sub Form_Load()
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupChiefCashier Then
            cmdSave.Visible = False
            optApprove.Visible = False
            optReject.Visible = False
        End If
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
            cmdPayment.Visible = False
            cmdRemitBack.Visible = False
            cmdSave.Visible = True
            optApprove.Visible = True
            optApprove.value = vbChecked
            optReject.Visible = False
            optReject.value = False
        End If
        Call FormInitiliaze
        If intID > 0 Then
            Call FillDetails
        End If
        Call CheckDemand
        vsGrid.SelectionMode = flexSelectionByRow
        txtDate.Text = gbTransactionDate
    End Sub

    Private Sub optApprove_Click()
        cmdSave.Enabled = True
    End Sub

    Private Sub optReject_Click()
        cmdSave.Enabled = True
    End Sub

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtPaidTo_LostFocus()
        If txtPaidTo.Text <> "" Then
            txtPaidTo.Text = FormatIntoProperCase(txtPaidTo.Text)
        End If
    End Sub

    Private Sub vsGrid_DblClick()
        Call FormInitiliaze
        If vsGrid.Row <> 0 Then
            txtPaidTo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            txtPaidTo.Tag = vsGrid.TextMatrix(vsGrid.Row, 5)
            txtAmount.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            txtAmount.Tag = vsGrid.TextMatrix(vsGrid.Row, 3)
            txtReference.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            txtRemarks.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
        End If
    End Sub

'
'    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'        If CalculateAmt > Val(txtAmount.Text) Then
'            MsgBox "Amount exceed", vbInformation
'            vsGrid.TextMatrix(vsGrid.Row, vsGrid.Col) = ""
''            vsGrid.SetFocus
'            'vsGrid.RemoveItem
'        Else
'            txtTotal.Text = CalculateAmt
'        End If
'    End Sub
'
    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim mRowCount       As Integer
        Dim mCnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        
        '*********************************************************************************************'
        '                    Procedure to delete a particular disbursement                            '
        '*********************************************************************************************'
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            'If gbUserTypeID = 3 Then
            If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupChiefCashier Then
                
                
                If KeyCode = vbKeyDelete Then
                    If vsGrid.Rows <> 0 Then 'And vsGrid.Row <> 1 Then
                        If MsgBox("Do you want to delete this record ?", vbYesNo) = vbYes Then
                            Call FormInitiliaze
                            Call GetCashBookDetails(20)
                            txtTotal.Text = val(txtTotal.Text) - val(vsGrid.TextMatrix(vsGrid.Row, 3))
                            txtBalance.Text = val(txtBalance.Text) + val(vsGrid.TextMatrix(vsGrid.Row, 3))
                            mCnn.Execute "Delete From faSubsidiaryCashBookChild Where intID = " & mID & " And intSerialNo = " & vsGrid.TextMatrix(vsGrid.Row, 5)
                            mCnn.Execute "Update faSubsidiaryCashBook Set fltAmount = " & val(txtTotal.Text) & " Where intID = " & mID
                            For mRowCount = vsGrid.Row + 1 To vsGrid.Rows - 1
                                mCnn.Execute "Update faSubsidiaryCashBookChild Set intSerialNo = " & val(vsGrid.TextMatrix(mRowCount, 5)) - 1 & " Where intID = " & mID & " And intSerialNo = " & val(vsGrid.TextMatrix(mRowCount, 5))
                                vsGrid.TextMatrix(mRowCount, 0) = vsGrid.TextMatrix(mRowCount, 0) - 1
                                vsGrid.TextMatrix(mRowCount, 5) = vsGrid.TextMatrix(mRowCount, 5) - 1
                                'mCnn.Execute "Update faSubsidiaryCashBookChild Set intSerialNo = " & Val(vsGrid.TextMatrix(mRowCount, 5)) & " Where intID = " & mID
                            Next
            '                mCnn.Execute "Delete From faSubsidiaryCashBookChild Where intID = " & mID & " And intSerialNo = " & vsGrid.TextMatrix(vsGrid.Row, 5)
                            vsGrid.RemoveItem (vsGrid.Row)
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
    End Sub
    
    Public Property Let intID(Data As Integer)
        mintID = Data
    End Property

    Public Property Get intID() As Integer
        intID = mintID
    End Property
    
    Public Property Let TransferID(Data As Integer)
        mTransferID = Data
    End Property

    Public Property Get TransferID() As Integer
        TransferID = mTransferID
    End Property

    Public Property Let AmtReceived(Data As Double)
        mAmtReceived = Data
    End Property

    Public Property Get AmtReceived() As Double
        AmtReceived = mAmtReceived
    End Property

    Public Property Let DemandID(Data As Variant)
        mDemandID = Data
    End Property

    Public Property Get DemandID() As Variant
        DemandID = mDemandID
    End Property

'
'    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
'        If vsGrid.Col = 3 And vsGrid.Row = vsGrid.Rows - 1 Then
'            If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" And vsGrid.TextMatrix(vsGrid.Row, 3) <> "" Then
'                If KeyAscii = 13 Then
'                    vsGrid.Rows = vsGrid.Rows + 1
'                    vsGrid.Col = 1
'                    vsGrid.Row = vsGrid.Row + 1
'                    vsGrid.TextMatrix(vsGrid.Row, 0) = vsGrid.TextMatrix(vsGrid.Row - 1, 0) + 1
'                    vsGrid.TextMatrix(vsGrid.Row, 5) = Val(vsGrid.TextMatrix(vsGrid.Row, 0))
'                End If
'            End If
'        End If
'    End Sub
'
'    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'        If vsGrid.Col = 3 Then
'            If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 13 Or KeyAscii = 8) Then
'                KeyAscii = 0
'            End If
'        End If
'    End Sub
'
'    Private Function CalculateAmt() As Variant
'        Dim mCount As Integer
'        Dim mTOt As Variant
'        mTOt = 0
'        'If Val(cmbTransactionType.Tag) = 1001 Or Val(cmbTransactionType.Tag) = 1002 Or Val(cmbTransactionType.Tag) = 1003 Then 'PayBill/workbill
'            For mCount = 1 To vsGrid.Rows - 1
'                If Trim(vsGrid.TextMatrix(mCount, 3)) = "" Then Exit For
'                mTOt = Val(mTOt) + Val(vsGrid.TextMatrix(mCount, 3))
'            Next
'            CalculateAmt = mTOt
'        'End If
'    End Function
'    Private Sub vsGrid_Click()
'        Call FormInitiliaze
'        If vsGrid.Row <> 0 Then
'            txtPaidTo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
'            txtPaidTo.Tag = vsGrid.TextMatrix(vsGrid.Row, 5)
'            txtAmount.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
'            txtAmount.Tag = vsGrid.TextMatrix(vsGrid.Row, 3)
'            txtReference.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
'            txtRemarks.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
'        End If
'End Sub
