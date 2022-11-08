VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPTaxMultipleBuilding 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmPTaxMultipleBuilding"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
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
      Height          =   660
      Left            =   1650
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5490
      Width           =   4695
   End
   Begin VB.TextBox txtCount 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   19
      Top             =   5070
      Width           =   870
   End
   Begin VB.Frame framParty 
      BackColor       =   &H00EAFFFF&
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
      Height          =   2310
      Left            =   0
      TabIndex        =   4
      Top             =   420
      Width           =   9960
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   30
         Top             =   1575
         Width           =   1650
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8415
         MaxLength       =   6
         TabIndex        =   29
         Top             =   1740
         Width           =   945
      End
      Begin VB.TextBox txtPost 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6510
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1740
         Width           =   1650
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   27
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtLocalPlace 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   26
         Top             =   1140
         Width           =   3375
      End
      Begin VB.TextBox txtStreet 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   25
         Top             =   810
         Width           =   3375
      End
      Begin VB.TextBox txtHouse 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   24
         Top             =   420
         Width           =   3375
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   23
         Top             =   150
         Width           =   3375
      End
      Begin VB.TextBox txtBuildingGr 
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
         Left            =   1650
         TabIndex        =   9
         Top             =   360
         Width           =   2940
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4620
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   465
      End
      Begin VB.TextBox txtWardNo 
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
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1095
         Width           =   870
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
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2580
      End
      Begin VB.ComboBox cmbZone 
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
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   3480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
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
         Height          =   240
         Left            =   900
         TabIndex        =   38
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
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
         Height          =   240
         Left            =   8190
         TabIndex        =   37
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
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
         Height          =   240
         Left            =   6165
         TabIndex        =   36
         Top             =   1785
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
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
         Height          =   240
         Left            =   5715
         TabIndex        =   35
         Top             =   1470
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Place"
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
         Height          =   240
         Left            =   5655
         TabIndex        =   34
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
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
         Height          =   240
         Left            =   5940
         TabIndex        =   33
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House/Office"
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
         Height          =   240
         Left            =   5400
         TabIndex        =   32
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   240
         Left            =   6000
         TabIndex        =   31
         Top             =   210
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building Group No"
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
         Height          =   240
         Left            =   60
         TabIndex        =   12
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   11
         Top             =   1155
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
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
         Height          =   240
         Left            =   1170
         TabIndex        =   10
         Top             =   765
         Width           =   360
      End
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
      Left            =   8595
      TabIndex        =   3
      Top             =   6180
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
      Left            =   6690
      TabIndex        =   2
      Top             =   6180
      Width           =   1875
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   8535
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1305
   End
   Begin VB.TextBox txtNoticeFee 
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
      Left            =   4530
      TabIndex        =   0
      Top             =   5070
      Width           =   1305
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3660
      Top             =   7035
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2265
      Left            =   -15
      TabIndex        =   13
      Top             =   2775
      Width           =   9960
      _cx             =   17568
      _cy             =   3995
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPTaxMultipleBuilding.frx":0000
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
      Caption         =   "Remarks"
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
      Left            =   840
      TabIndex        =   22
      Top             =   5550
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
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
      Left            =   1110
      TabIndex        =   20
      Top             =   5100
      Width           =   510
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property tax Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   3645
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
      Left            =   7215
      TabIndex        =   17
      Top             =   5055
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
      Left            =   8535
      TabIndex        =   16
      Top             =   5055
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
      Left            =   7365
      TabIndex        =   15
      Top             =   5370
      Width           =   1140
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notice Fee"
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
      Height          =   225
      Left            =   3390
      TabIndex        =   14
      Top             =   5070
      Width           =   1035
   End
End
Attribute VB_Name = "frmPTaxMultipleBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub FillZone()
        Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Where intLBID = " & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster)
End Sub
Private Sub FillWard()
    Dim mSql As String
    On Error Resume Next
    mSql = "SELECT chvWardNameEnglish, intWardNo, numWardID FROM GM_Ward"
    mSql = mSql + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID
    mSql = mSql + " AND numZoneID = " & cmbZone.ItemData(cmbZone.ListIndex)
    mSql = mSql + " AND intWardYear = " & cmbAssessmentYear.Text
    mSql = mSql + " Order By chvWardNameEnglish"
    PopulateList cmbWard, mSql, , , , True, enuSourceString.DBMaster
End Sub

Private Sub cmdCancel_Click()
    Dim objTranType As New clsTransactionType
    Unload Me
    If Not frmReceiptsCounter.InterruptEditMode Then
        objTranType.SetTransactionType (9999)
        frmReceiptsCounter.txtTransactionType.Text = objTranType.TransactionType
        frmReceiptsCounter.txtTransactionType.Tag = objTranType.TransactionTypeID
    End If
End Sub

Private Sub cmdCopy_Click()
    If txtBuildingGr.Tag <> 0 Then
            frmReceiptsCounter.SubLedgerID = numBuildingGrID
            frmReceiptsCounter.cmbZone.Text = cmbZone.Text
            frmReceiptsCounter.cmbDZone.Text = cmbZone.Text '   Added   '
            frmReceiptsCounter.txtWard.Text = cmbWard.Text
            frmReceiptsCounter.txtWardNo.Text = val(txtWardNo) 'cmbWard.ItemData(cmbWard.ListIndex) '   Added   '
            If cmbWard.ListIndex > -1 Then
                frmReceiptsCounter.txtWard.Tag = cmbWard.ItemData(cmbWard.ListIndex)
            End If
            frmReceiptsCounter.txtHouseNo1.Text = vchHouseName_4 'txtHouseNo1.Text
            frmReceiptsCounter.txtHouseNo2.Text = txtHouseNo2.Text
            frmReceiptsCounter.txtDoorNo1.Text = txtHouseNo1.Text   '   Added   '
            frmReceiptsCounter.txtDoorNo2.Text = txtHouseNo2.Text   '   Added   '
            frmReceiptsCounter.txtHouse.Text = vchHouseName_4
            frmReceiptsCounter.txtStreet.Text = vchStreetName_5
            frmReceiptsCounter.txtMainPlace.Text = vchMainPlace_6
            frmReceiptsCounter.txtPost.Text = vchPostOffice_7
            frmReceiptsCounter.SubLedgerID = mBuildingID
            frmReceiptsCounter.txtName.Text = vchName_3
            frmReceiptsCounter.txtRefNo.Text = vchRef_11
            frmReceiptsCounter.txtDescription.Text = vchNarration_10
            frmReceiptsCounter.txtBuildingNo.Text = mBuildingID
            frmReceiptsCounter.SubLedgerID = mBuildingID
            frmReceiptsCounter.DemandBasedFlag = True

            'frmReceiptsCounter.AssessmentYear = cmbAssessmentYear.ItemData(cmbAssessmentYear.ListIndex)
            '--------------------------------------------------------------------'
        End If
End Sub

Private Sub cmdsearch_Click()
     Dim arrInput        As Variant
        Dim Rec             As New ADODB.Recordset
        Dim objdb           As New clsDB
        Dim objAcc          As New clsAccounts
        Dim mCnn            As New ADODB.Connection
        Dim objPTax         As New clsPTax
        Dim mWardID         As Double
        Dim mSql            As String
        Dim mXmlStream      As New ADODB.Stream
        Dim RecID           As New ADODB.Recordset
        Dim mBuildingWeb    As String
        If cmbWard.ListIndex > -1 Then
            mWardID = cmbWard.ItemData(cmbWard.ListIndex)
        End If

        Dim numBuildingGrID   As Variant
'        Dim numZoneID       As Variant
'        Dim intAssessmentYear As Variant
'        Dim intWardNo       As Variant
'        Dim intDoorNo1      As Variant
'        Dim chvDoorNo2      As Variant
'        Dim chvName         As Variant
'        Dim chvResHName     As Variant
        Dim mUrl            As String
        Dim mArrIN          As String
        Dim mArrOut         As String
        Dim mArrAdd         As String
'        Dim mGetString()    As String
        Dim client          As New MSSOAPLib.SoapClient
        Dim objSOAP         As Variant
        Dim mSlNo           As Variant
         
            mvarBuildingID = -1
            Me.MousePointer = vbHourglass
            vsGrid.Clear 1, 0
            'txtAddress.Text = ""
            txtGrandTotal.Text = ""
            lblTotalArrear.Caption = ""
            lblTotalCurrent.Caption = ""
                
                          
           'Added On 28 Feb 2018 this will generate demand in Sanchaya

                numBuildingGrID = CStr(IIf(IsNull(txtBuildingGr.Text), 0, txtBuildingGr.Text))
                If numBuildingGrID > 1 Then
                    
                Else
                    MsgBox "Please Enter Building Group No", vbApplicationModal
                    Exit Sub
                End If
              
                mUrl = gbDefaultUrl
                
                Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                objSOAP.MSSoapInit mUrl + "?WSDL"
                On Error GoTo WebConnectionERROR:
               numBuildingGrID = 5074001013#
                   ' mArrOut = objSOAP.getGroupDemandSaankhyaXML(gbLBID, numBuildingGrID)
              mArrOut = objSOAP.getGroupDemandSaankhyaXML(gbLBID, numBuildingGrID)
                
            On Error GoTo ERROR_AfterWEBService:
                mXmlStream.Open

                mXmlStream.WriteText mArrOut
                mXmlStream.Position = 0
                Rec.Open mXmlStream
                mXmlStream.Close


                If Not (Rec.BOF And Rec.EOF) Then
                
                     vsGrid.Rows = 1
                    mRows = 1
                    vsGrid.MergeCells = flexMergeFree
                    While Not Rec.EOF

                        '--------------------------------------------------'
                        ' Beging of Block - Inserting Demands in Rows      '
                        '--------------------------------------------------'
                        vsGrid.Rows = vsGrid.Rows + 1
                        objAcc.SetAccountCode (Rec!chvSanHeadCode)
                        
                        vsGrid.Cell(flexcpText, mRows, 0) = Rec!chvSanHeadCode
                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead

                        vsGrid.MergeCol(12) = True
                        vsGrid.Cell(flexcpChecked, mRows, 12) = 1
                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                        vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
'                        vsGrid.Cell(flexcpText, mRows, 8) = Rec!tnyPeriodID
'                        vsGrid.Cell(flexcpText, mRows, 9) = Rec!ArrFlag
                        vsGrid.Cell(flexcpText, mRows, 10) = Rec!intGroupSlNo
                        vsGrid.Cell(flexcpText, mRows, 5) = Rec!fltAmount
                        'vsGrid.Cell(flexcpText, mRows, 11) = Rec!fltAmount
                        vsGrid.Cell(flexcpText, mRows, 13) = Rec!intGroupSlNo
                        mSlNo = Rec!intGroupSlNo
                        mRows = mRows + 1
                        Rec.MoveNext
          
                        Wend
                    Rec.Close
                    
                    mArrAdd = objSOAP.getGroupDetailsSaankhyaXML(gbLBID, numBuildingGrID, mSlNo)
                    
    On Error GoTo ERROR_AfterWEBService:
    
                    mXmlStream.Open
                    mXmlStream.WriteText mArrAdd
                    mXmlStream.Position = 0
                    Rec.Open mXmlStream
                    mXmlStream.Close
                    
                    txtCount.Text = Rec!BuildingCount
                    txtName.Text = Rec!chvownerEng
                    txtHouse.Text = Rec!chvHouseNameEng
                    txtStreet.Text = Rec!chvResStreetName
                    txtMainPlace.Text = Rec!chvMainPlace
                    txtLocalPlace.Text = Rec!chvLocalPlace
                    txtWardNo.Text = Rec!intWardNo
                    cmbWard.ListIndex = Rec!intWardNo
                    
                Else
                    MsgBox "No Record Found"
                    Exit Sub
                End If
               
                Rec.Close
                Set mCnn = Nothing
            
            Me.MousePointer = vbDefault
            Exit Sub
WebConnectionERROR:
        MsgBox "Connection to Web Service Failed :: " & Error, vbInformation
        Exit Sub
ERROR_AfterWEBService:
        MsgBox Error
        
        Me.MousePointer = vbDefault
End Sub

    Private Sub Form_Load()
        Call FillZone
        Call FillWard
    End Sub
    Private Sub txtBuildingGr_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then
            Call PressTabKey
         End If
    End Sub

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

