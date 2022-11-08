VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmScheduleRatesForBirthDeath 
   Caption         =   "Schedule Rates for Birth / Death / Marriage Certificates"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   11655
   Begin VB.TextBox txtRemarks 
      Height          =   825
      Left            =   9030
      TabIndex        =   49
      Top             =   3960
      Width           =   2475
   End
   Begin VB.Frame fmeBithDeath 
      Height          =   495
      Left            =   5940
      TabIndex        =   45
      Top             =   0
      Width           =   2085
      Begin VB.OptionButton optDeath 
         Caption         =   "Death"
         Height          =   225
         Left            =   1080
         TabIndex        =   47
         Top             =   180
         Width           =   795
      End
      Begin VB.OptionButton optBirth 
         Caption         =   "Birth"
         Height          =   225
         Left            =   240
         TabIndex        =   46
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.ComboBox cmbTransactionType 
      Height          =   345
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   90
      Width           =   4065
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
      Left            =   8220
      TabIndex        =   43
      Top             =   5100
      Width           =   1875
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
      Left            =   10125
      TabIndex        =   42
      Top             =   5100
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Height          =   2490
      Left            =   30
      TabIndex        =   7
      Top             =   3060
      Width           =   7890
      Begin VB.TextBox txtWardNo 
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
         Left            =   945
         MaxLength       =   3
         TabIndex        =   26
         Top             =   555
         Width           =   1800
      End
      Begin VB.TextBox txtDoorNo1 
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
         Left            =   945
         MaxLength       =   5
         TabIndex        =   25
         Top             =   870
         Width           =   1110
      End
      Begin VB.TextBox txtDoorNo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   24
         Top             =   870
         Width           =   690
      End
      Begin VB.ComboBox cmbZone 
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
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   210
         Width           =   1800
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3885
         MaxLength       =   100
         TabIndex        =   22
         Top             =   210
         Width           =   2535
      End
      Begin VB.TextBox txtHouseName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3885
         MaxLength       =   100
         TabIndex        =   21
         Top             =   540
         Width           =   3210
      End
      Begin VB.TextBox txtStreet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3885
         MaxLength       =   100
         TabIndex        =   20
         Top             =   855
         Width           =   3210
      End
      Begin VB.TextBox txtLocalPlace 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3885
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1170
         Width           =   3210
      End
      Begin VB.TextBox txtMainPlace 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3885
         MaxLength       =   100
         TabIndex        =   18
         Top             =   1485
         Width           =   3210
      End
      Begin VB.TextBox txtInitial1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6435
         MaxLength       =   1
         TabIndex        =   17
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInitial2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6750
         MaxLength       =   1
         TabIndex        =   16
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInitial3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7065
         MaxLength       =   1
         TabIndex        =   15
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInitial4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7380
         MaxLength       =   1
         TabIndex        =   14
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtPost 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3885
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1800
         Width           =   2025
      End
      Begin VB.TextBox txtPin 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6180
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox txtPhone 
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
         Left            =   3885
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2115
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1095
         MaxLength       =   3
         TabIndex        =   10
         Top             =   2115
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox txtRefNo 
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
         Left            =   945
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1185
         Width           =   1800
      End
      Begin VB.ComboBox cmbInstrumentType 
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
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1770
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   45
         TabIndex        =   40
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   60
         TabIndex        =   39
         Top             =   915
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3360
         TabIndex        =   37
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House/Office"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   2775
         TabIndex        =   36
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3300
         TabIndex        =   35
         Top             =   915
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Place"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   2910
         TabIndex        =   34
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   2940
         TabIndex        =   33
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3465
         TabIndex        =   32
         Top             =   1845
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   5940
         TabIndex        =   31
         Top             =   1845
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3045
         TabIndex        =   30
         Top             =   2175
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drawn From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   30
         TabIndex        =   29
         Top             =   2160
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reff No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1230
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   1830
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   9525
      TabIndex        =   4
      Top             =   3390
      Width           =   1935
   End
   Begin VB.TextBox txtCerrtificateCharge 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   9525
      TabIndex        =   3
      Top             =   3090
      Width           =   1935
   End
   Begin VB.TextBox txtNoOfCertificate 
      Enabled         =   0   'False
      Height          =   330
      Left            =   10740
      TabIndex        =   1
      Top             =   120
      Width           =   705
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2430
      Left            =   30
      TabIndex        =   41
      Top             =   570
      Width           =   11520
      _cx             =   20320
      _cy             =   4286
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmScheduleRatesForBirthDeath.frx":0000
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
      Editable        =   1
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      Height          =   225
      Left            =   8130
      TabIndex        =   48
      Top             =   3960
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   225
      Left            =   8490
      TabIndex        =   6
      Top             =   3450
      Width           =   960
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Certificate Charge"
      Height          =   225
      Left            =   7950
      TabIndex        =   5
      Top             =   3150
      Width           =   1485
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No of Certificate"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9330
      TabIndex        =   2
      Top             =   165
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Type"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   1410
   End
End
Attribute VB_Name = "frmScheduleRatesForBirthDeath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private mReceiptDemandFlagSelect As Variant
    
    Dim mScheduleType As Integer
    Dim mScheduleSubID As Integer

Public Property Let mReceiptOrDemandFlag(mReceiptDemandFlag As Variant)
        mReceiptDemandFlagSelect = mReceiptDemandFlag
End Property

Private Sub Forminitialize()
    vsGrid.Clear 1, 1
    optBirth.value = False
    optDeath.value = False
End Sub
Private Sub cmbTransactionType_Click()
    fmeBithDeath.Visible = True
    If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = 11 Then 'Marriage '
        mScheduleType = 15
        mScheduleSubID = 3
    Else
        Call Forminitialize
    End If
    Call FillGrid
End Sub

Private Sub cmdCopy_Click()
    Dim mLoop As Long
    Dim mLoopChild As Long
    Dim mCount As Long
    Dim mCheck As Integer
    mCheck = 0
    For mCount = 1 To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpChecked, mCount, 0) = vbChecked Then
            mCheck = mCheck + 1
        End If
    Next
    
    If mCheck = 0 Then
        MsgBox "Check any one of the following List before copying", vbInformation
        Exit Sub
    End If
    
    If txtName.Text = "" Then
        MsgBox "Please Give the Name bofore Copying", vbInformation
        Exit Sub
    End If

    If mReceiptDemandFlagSelect = 1 Then
    
        frmReceiptsCounter.txtWardNo = txtWardNo.Text
        frmReceiptsCounter.txtDoorNo1 = txtDoorNo1.Text
        frmReceiptsCounter.txtDoorNo2 = txtDoorNo2.Text
        frmReceiptsCounter.txtRefNo = txtRefNo.Text
        frmReceiptsCounter.txtName = txtName.Text
        frmReceiptsCounter.txtInit1 = txtInitial1.Text
        frmReceiptsCounter.txtInit2 = txtInitial2.Text
        frmReceiptsCounter.txtInit3 = txtInitial3.Text
        frmReceiptsCounter.txtInit4 = txtInitial4.Text
        frmReceiptsCounter.txtHouse = txtHouseName.Text
        frmReceiptsCounter.txtStreet = txtStreet.Text
        frmReceiptsCounter.txtLocalPlace = txtLocalPlace.Text
        frmReceiptsCounter.txtMainPlace = txtMainPlace.Text
        frmReceiptsCounter.txtPost = txtPost.Text
        frmReceiptsCounter.txtPin = txtPin.Text
        frmReceiptsCounter.txtPhone = txtPhone.Text
        frmReceiptsCounter.txtTotalCurrent = txtGrandTotal.Text
        frmReceiptsCounter.txtTotal = txtGrandTotal.Text
        frmReceiptsCounter.txtDescription = txtRemarks.Text
        
        frmReceiptsCounter.vsGrid.Rows = 1
        frmReceiptsCounter.vsGrid.MergeCells = flexMergeFree
        mCount = 0
        For mLoop = 1 To vsGrid.Rows - 1
            frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 1
            If vsGrid.Cell(flexcpChecked, mLoop, 0) = 1 Then
                mCount = mCount + 1
                frmReceiptsCounter.vsGrid.Row = mCount
                If vsGrid.Cell(flexcpChecked, mLoop, 0) = vbChecked Then
                    frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 0) = vsGrid.Cell(flexcpText, mLoop, 2)
                    If mLoop = 1 And val(txtNoOfCertificate) <> 0 Then
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 5) = vsGrid.Cell(flexcpText, mLoop, 7) * val(txtNoOfCertificate.Text)
                    Else
                        frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 5) = vsGrid.Cell(flexcpText, mLoop, 7)
                    End If
                    frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, 6) = vsGrid.Cell(flexcpText, mLoop, 1)
                    frmReceiptsCounter.vsGrid.Cell(flexcpChecked, mCount, 12) = 1
                End If
            End If
        Next
    ElseIf mReceiptDemandFlagSelect = 2 Then
        
        frmDemandInterface.cmbSections.Text = "Municipal Health Department"
        If cmbTransactionType.Text = "Birth/Death Registration" Then
            'frmDemandInterface.cmbTransactionType.Text = "Birth & Death Registration"
            frmDemandInterface.txtTransactionType.Text = "Birth & Death Registration"
            frmDemandInterface.txtTransactionType.Tag = 12
        Else
            'frmDemandInterface.cmbTransactionType.Text = "Marriage Registration"
            frmDemandInterface.txtTransactionType.Text = "Marriage Registration"
            frmDemandInterface.txtTransactionType.Tag = 11
        End If
        frmDemandInterface.txtWardNo = txtWardNo.Text
        frmDemandInterface.txtDoorNo1 = txtDoorNo1.Text
        frmDemandInterface.txtDoorNo2 = txtDoorNo2.Text
        frmDemandInterface.txtReference = txtRefNo.Text
        frmDemandInterface.txtName = txtName.Text
        frmDemandInterface.txtInitial1 = txtInitial1.Text
        frmDemandInterface.txtInitial2 = txtInitial2.Text
        frmDemandInterface.txtInitial3 = txtInitial3.Text
        frmDemandInterface.txtInitial4 = txtInitial4.Text
        frmDemandInterface.txtHouseName = txtHouseName.Text
        frmDemandInterface.txtStreet = txtStreet.Text
        frmDemandInterface.txtLocalPlace = txtLocalPlace.Text
        frmDemandInterface.txtMainPlace = txtMainPlace.Text
        frmDemandInterface.txtPost = txtPost.Text
        frmDemandInterface.txtPin = txtPin.Text
        frmDemandInterface.txtPhone = txtPhone.Text
        frmDemandInterface.txtCurrentAmt = txtGrandTotal.Text
        frmDemandInterface.txtGrandTotal = txtGrandTotal.Text
        
        frmDemandInterface.vsGrid.Rows = 1
        frmDemandInterface.vsGrid.MergeCells = flexMergeFree
        mCount = 0
        For mLoop = 1 To vsGrid.Rows - 1
            frmDemandInterface.vsGrid.Rows = frmDemandInterface.vsGrid.Rows + 1
            If vsGrid.Cell(flexcpChecked, mLoop, 0) = 1 Then
                mCount = mCount + 1
                frmDemandInterface.vsGrid.Row = mCount
                If vsGrid.Cell(flexcpChecked, mLoop, 0) = vbChecked Then
                    frmDemandInterface.vsGrid.Cell(flexcpText, mCount, 0) = vsGrid.Cell(flexcpText, mLoop, 2)
                    frmDemandInterface.vsGrid.Cell(flexcpText, mCount, 1) = vsGrid.Cell(flexcpText, mLoop, 3)
                    frmDemandInterface.vsGrid.Cell(flexcpText, mCount, 5) = vsGrid.Cell(flexcpText, mLoop, 7)
                    frmDemandInterface.vsGrid.Cell(flexcpText, mCount, 6) = vsGrid.Cell(flexcpText, mLoop, 1)
                    frmDemandInterface.vsGrid.Cell(flexcpChecked, mCount, 12) = 1
                End If
            End If
        Next
    End If
    Unload Me
End Sub

Private Sub fmeBithDeath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OptionForBirthDeath
    Call FillGrid
End Sub

Private Sub Form_Activate()
    Me.Height = 7155
    Me.Width = 11940
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Load()
    Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Order By chvZoneNameEnglish", , True, True, True, DBMaster)
    Call FillGrid
    Call FillTransType
    If mReceiptDemandFlagSelect = 1 Then
        If frmReceiptsCounter.txtTransactionType.Text = "Birth & Death Registration" Then
            cmbTransactionType.Text = "Birth/Death Registration"
        Else
            cmbTransactionType.Text = "Marriage Registration"
        End If
    End If
End Sub

Private Sub FillTransType()
    cmbTransactionType.AddItem "Birth/Death Registration"
    cmbTransactionType.ItemData(cmbTransactionType.NewIndex) = 12
''    cmbTransactionType.AddItem "Death Registration"
''    cmbTransactionType.ItemData(cmbTransactionType.NewIndex) = 12
    cmbTransactionType.AddItem "Marriage Registration"
    cmbTransactionType.ItemData(cmbTransactionType.NewIndex) = 11
End Sub

Private Sub FillGrid()
    Dim mSQL As String
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mRowCount As Integer
    objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    mSQL = "SELECT  smScheduleMasters.intScheduleID, smScheduleMasters.fltFixedRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID "
    mSQL = mSQL + " FROM smScheduleMasters INNER JOIN "
    mSQL = mSQL + " smAttributes ON smScheduleMasters.intAttributeID = smAttributes.intAttributeID "
    mSQL = mSQL + " WHERE   (smScheduleMasters.intScheduleID = " & mScheduleType & ") and smAttributes.tnyGroupID = " & mScheduleSubID
    mSQL = mSQL + " ORDER By smAttributes.intOrderBy"
    Rec.Open mSQL, mCnn
    mRowCount = 1
    vsGrid.Rows = 2
    While Not Rec.EOF And Not Rec.BOF
        vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchAttributeTitle), "", Rec!vchAttributeTitle)
        vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltFixedRate), "", Format("0.00", Rec!fltFixedRate))
        mRowCount = mRowCount + 1
        vsGrid.Rows = vsGrid.Rows + 1
        Rec.MoveNext
    Wend
End Sub

Private Sub optBirth_Click()
    Call OptionForBirthDeath
    Call FillGrid
End Sub

Private Sub optDeath_Click()
    OptionForBirthDeath
    Call FillGrid
End Sub

Private Sub txtNoOfCertificate_LostFocus()
    Dim mRowCount As Integer
    Dim mFlag As Boolean
    Dim mTotalAmount  As Double
    mFlag = False
    If val(txtNoOfCertificate.Text) >= 1 Then
        txtCerrtificateCharge.Text = val(txtNoOfCertificate.Text) * vsGrid.TextMatrix(1, 7)
        txtCerrtificateCharge.Tag = (val(txtNoOfCertificate.Text) - 1) * vsGrid.TextMatrix(1, 7)
        For mRowCount = 2 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked Then
                mFlag = True
            End If
        Next
    Else
        txtCerrtificateCharge.Text = 0
    End If
    If mFlag = False Then
        'txtGrandTotal.Text = Val(txtCerrtificateCharge.Tag)
        txtGrandTotal.Text = val(txtCerrtificateCharge.Text)
        mTotalAmount = val(txtGrandTotal.Text)
    End If
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim mRowCount As Integer
    Dim mCheck As Boolean
    Dim mCount As Integer
    Dim mTempTotal As Double
    
    Dim mCertificateCharge As Variant
    
    If vsGrid.Cell(flexcpChecked, 1, 0) = vbChecked Then
        txtNoOfCertificate.Enabled = True
    Else
        txtNoOfCertificate.Enabled = False
    End If
    
    For mRowCount = 2 To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked Then
            txtNoOfCertificate.Enabled = False
        End If
    Next
    
    
    For mRowCount = 1 To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked Then
            mCheck = True
        End If
    Next
    
    If vsGrid.Cell(flexcpChecked, 1, 0) = vbChecked Then
        mCertificateCharge = IIf(val(txtNoOfCertificate.Text) = 0, 1, val(txtNoOfCertificate.Text)) * vsGrid.TextMatrix(1, 7)
        txtCerrtificateCharge.Text = mCertificateCharge
        txtCerrtificateCharge.Tag = mCertificateCharge - vsGrid.TextMatrix(1, 7)
    Else
        txtCerrtificateCharge.Text = ""
        txtCerrtificateCharge.Tag = 0
    End If
    
    If mCheck = True Then
        For mRowCount = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked Then
                mTempTotal = mTempTotal + val(vsGrid.Cell(flexcpText, mRowCount, 7))
            End If
        Next mRowCount
    Else
        txtCerrtificateCharge.Text = ""
        txtGrandTotal.Text = ""
        mTempTotal = 0
        txtCerrtificateCharge.Tag = 0
    End If
    txtGrandTotal.Text = Format("0.00", mTempTotal + val(txtCerrtificateCharge.Tag))
    
End Sub

Private Sub CheckForReceiptOrDemand()
    If mReceiptDemandFlagSelect = 1 Then
        Call FillTransType
    ElseIf mReceiptDemandFlagSelect = 2 Then
        Call FillTransType
    End If
End Sub

Private Sub OptionForBirthDeath()
    If optBirth.value = True Then
        mScheduleType = 13
        mScheduleSubID = 1
    ElseIf optDeath.value = True Then
        mScheduleType = 13
        mScheduleSubID = 2
    End If
End Sub

