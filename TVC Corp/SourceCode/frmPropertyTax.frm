VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPropertyTax 
   BackColor       =   &H00EAFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10725
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
      Left            =   5055
      TabIndex        =   46
      Top             =   4590
      Width           =   1305
   End
   Begin VB.TextBox txtNetAmount 
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
      Left            =   8580
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5430
      Width           =   1305
   End
   Begin VB.TextBox txtFromYear 
      Height          =   285
      Left            =   1605
      TabIndex        =   40
      Top             =   5220
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   5055
      TabIndex        =   37
      Top             =   4905
      Width           =   1305
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
      Left            =   3135
      TabIndex        =   36
      Top             =   5895
      Width           =   420
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
      Left            =   1590
      TabIndex        =   35
      Top             =   5895
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
      Left            =   3135
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   5565
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
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   5565
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
      Left            =   3135
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   5145
      Width           =   1455
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
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   5145
      Width           =   1530
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
      Left            =   1590
      TabIndex        =   26
      Top             =   4845
      Width           =   1515
   End
   Begin VB.TextBox txtAdvance 
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
      Left            =   8565
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   5145
      Width           =   1305
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
      Left            =   8565
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4845
      Width           =   1305
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
      Left            =   7200
      TabIndex        =   13
      Top             =   5820
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
      Left            =   9105
      TabIndex        =   17
      Top             =   5820
      Width           =   1395
   End
   Begin VB.CheckBox chkFineWaiver 
      BackColor       =   &H80000018&
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
      Left            =   3180
      TabIndex        =   16
      Top             =   4920
      Width           =   1335
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3630
      Top             =   6555
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
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
      Height          =   1890
      Left            =   30
      TabIndex        =   15
      Top             =   -60
      Width           =   10710
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
         Left            =   5805
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   285
         Width           =   4695
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   3480
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
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   870
         Width           =   2580
      End
      Begin VB.ComboBox cmbAssessmentYear 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   3480
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
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   5
         Top             =   885
         Width           =   870
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "(Press F4 for Search)"
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
         Left            =   3810
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1470
         Visible         =   0   'False
         Width           =   1815
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
         Height          =   315
         Left            =   3300
         TabIndex        =   10
         Top             =   1185
         Width           =   480
      End
      Begin VB.TextBox txtBuildingNo 
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
         Left            =   1230
         TabIndex        =   12
         Top             =   1515
         Width           =   2550
      End
      Begin VB.TextBox txtHouseNo1 
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
         Left            =   1230
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1200
         Width           =   1140
      End
      Begin VB.TextBox txtHouseNo2 
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
         Left            =   2415
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   750
         TabIndex        =   2
         Top             =   585
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
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
         TabIndex        =   4
         Top             =   915
         Width           =   450
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
         Left            =   4995
         TabIndex        =   45
         Top             =   300
         Width           =   810
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   795
         TabIndex        =   0
         Top             =   255
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building No"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1545
         Width           =   960
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   1230
         Width           =   705
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2745
      Left            =   15
      TabIndex        =   14
      Top             =   1815
      Width           =   10710
      _cx             =   18891
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPropertyTax.frx":0000
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
   Begin VB.Label lblRRFlag 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   255
      Left            =   3420
      TabIndex        =   48
      Top             =   6270
      Visible         =   0   'False
      Width           =   3825
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
      Left            =   4140
      TabIndex        =   47
      Top             =   4590
      Width           =   885
   End
   Begin VB.Label lblGrandTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "------"
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
      Left            =   10215
      TabIndex        =   43
      Top             =   4905
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
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
      Left            =   7395
      TabIndex        =   42
      Top             =   5490
      Width           =   1140
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
      Left            =   4650
      TabIndex        =   38
      Top             =   4950
      Width           =   345
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
      Left            =   180
      TabIndex        =   34
      Top             =   5910
      Width           =   1380
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
      Left            =   750
      TabIndex        =   32
      Top             =   5625
      Width           =   810
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
      Left            =   525
      TabIndex        =   29
      Top             =   5280
      Width           =   1035
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
      Left            =   405
      TabIndex        =   27
      Top             =   4890
      Width           =   1155
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
      Left            =   6495
      TabIndex        =   25
      Top             =   5190
      Width           =   2040
   End
   Begin VB.Label lblFine 
      AutoSize        =   -1  'True
      Caption         =   "Fine(For Test)"
      Height          =   195
      Left            =   5040
      TabIndex        =   22
      Top             =   5940
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   7395
      TabIndex        =   21
      Top             =   4890
      Width           =   1140
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
      Left            =   8565
      TabIndex        =   20
      Top             =   4575
      Width           =   1305
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
      Left            =   7245
      TabIndex        =   19
      Top             =   4575
      Width           =   1305
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
      Left            =   75
      TabIndex        =   18
      Top             =   4575
      Width           =   1095
   End
End
Attribute VB_Name = "frmPropertyTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mDefaultTransactionTypeID   As Long
    Dim mDefaultAccountHeadCode     As String
    Dim mDefaultInstrumentID        As Long
    Dim mDefaultBankID              As Long
    Dim mDefaultBankHeadCode        As String
    Dim mDefaultZoneID              As Double
    Dim mFineAmt                    As Double   ' To Calculate the fine Amount '
    Dim vchName_3        As String
    
    Dim vchHouseName_4   As String
    Dim vchStreetName_5  As String
    Dim vchMainPlace_6   As String
    Dim vchPostOffice_7  As String
    Dim vchDistrict_8    As String
    Dim vchPinNumber_9   As String
    Dim vchNarration_10  As String
    Dim vchRef_11        As String
    
    Dim mBuildingID      As Double
    Dim mNumberOfSelections As Integer
    Dim mTransactionType            As Long
    Dim mDrAccountHeadID            As Long
    Dim mFineRate                   As Single
    Dim mAcHeadCodePTaxArrear       As String
    Dim mAcHeadCodePTaxNonResArrear As String
    Dim mAcHeadCodePTaxNonResCurrent As String
    Dim mAcHeadCodeFine             As String
    Dim mAcHeadCodeRoundOff         As String
    
    Private mPTaxArrearHeadCode     As String
    Private mPTaxCurrentHeadCode    As String
     
   
    
    
    Private mPTaxAdvanceCollected   As String
    Private mPTaxTransactionTypeID  As Long
    Private mPTaxLibraryCessCode    As String
    Private mvarBuildingID          As Double
    Public mvarDifferentZoneFlag   As Boolean
    Public mDemandWeb               As Boolean  ''P Tax Web
    
    
    Dim mFineWaiveFlag              As Boolean ' To Waive
    Dim mDescription                As String  ' To Waive
    Dim mSelectedAllFlag            As Boolean

    
    Dim mAdvAmt     As Double   ' Total Advance Amount
    Dim dtUptoDate  As Date     ' Fine Upto Date
    Dim dtFromDate  As Date
    Dim mAnyAdvanceFlag As Boolean ' Initial value Will be True.
    Dim mAdvCheckedRow  As Integer
    Dim mAdvanceExists As Boolean ' To flag if Advance came along with Demand
    
  
        
        
        
    
    Private Sub FillPTax()
        '--------------------------------------------------'
        ' Non-Demand PTax
        '--------------------------------------------------'
            If val(txtHalfYearTaxRate) <= 1 Then
                MsgBox "Please Enter Half year Tax rate", vbInformation
                txtHalfYearTaxRate.SetFocus
                Exit Sub
            End If
            
            Dim mLoop As Integer
            Dim mRow As Integer
            Dim mYearID As Long
            Dim mPeriod As Integer
            Dim objAcc As New clsAccounts
            Dim mAmtArrear As Double
            Dim mAmtCurrent As Double
            
            mYearID = cmbFromYear.ItemData(cmbFromYear.ListIndex)
            mPeriod = cmbFromPeriod.ItemData(cmbFromPeriod.ListIndex)
            vsGrid.Rows = 1
            For mLoop = 1 To val(txtNoOfHalfYears)
                '-------------------------------------------------------------------'
                ' Property Tax                                                      '
                '-------------------------------------------------------------------'
                vsGrid.Rows = vsGrid.Rows + 1
                mRow = vsGrid.Rows - 1
                If mYearID < gbFinancialYearID Then
                    objAcc.SetAccountCode mPTaxArrearHeadCode
                    vsGrid.Cell(flexcpText, mRow, 0) = mPTaxArrearHeadCode   'Rec!vchAccountHeadCode
                    vsGrid.Cell(flexcpText, mRow, 1) = objAcc.AccountHead
                Else
                    objAcc.SetAccountCode mPTaxCurrentHeadCode
                    vsGrid.Cell(flexcpText, mRow, 0) = mPTaxCurrentHeadCode   'Rec!vchAccountHeadCode
                    vsGrid.Cell(flexcpText, mRow, 1) = objAcc.AccountHead
                End If
                
                vsGrid.Cell(flexcpText, mRow, 2) = CStr(mYearID) & "-" & CStr(mYearID + 1)
                If mPeriod = 2 Then
                    vsGrid.Cell(flexcpText, mRow, 3) = "IInd Half"
                Else
                    vsGrid.Cell(flexcpText, mRow, 3) = "Ist Half"
                End If
                vsGrid.MergeCol(12) = True
                vsGrid.Cell(flexcpText, mRow, 12) = mLoop ' Rec!numDemandID
                vsGrid.Cell(flexcpChecked, mRow, 12) = 1
                vsGrid.Cell(flexcpText, mRow, 6) = objAcc.AccountHeadID ' Rec!intAccountHeadID
                vsGrid.Cell(flexcpText, mRow, 7) = mYearID ' Rec!intYearID
                vsGrid.Cell(flexcpText, mRow, 8) = mPeriod ' Rec!tnyPeriodID
                vsGrid.Cell(flexcpText, mRow, 9) = IIf(mYearID > gbFinancialYearID, 1, 0) 'Rec!tnyArrearFlag
                vsGrid.Cell(flexcpText, mRow, 10) = mLoop ' Rec!numDemandID
                vsGrid.Cell(flexcpText, mRow, 11) = txtHalfYearTaxRate  ' Rec!fltAmount
                If mYearID > Month(gbFinancialYearID) Then
                    mAmtArrear = mAmtArrear + val(txtHalfYearTaxRate.Text)
                    vsGrid.Cell(flexcpText, mRow, 4) = val(txtHalfYearTaxRate.Text)
                Else
                    mAmtCurrent = mAmtCurrent + val(txtHalfYearTaxRate.Text)
                    vsGrid.Cell(flexcpText, mRow, 5) = val(txtHalfYearTaxRate.Text)
                End If
                '-------------------------------------------------------------------'
                ' Library Cess                                                      '
                '-------------------------------------------------------------------'
                vsGrid.Rows = vsGrid.Rows + 1
                mRow = vsGrid.Rows - 1
                
                objAcc.SetAccountCode mPTaxLibraryCessCode
                vsGrid.Cell(flexcpText, mRow, 0) = objAcc.AccountCode    'Rec!vchAccountHeadCode
                vsGrid.Cell(flexcpText, mRow, 1) = objAcc.AccountHead
                
                vsGrid.Cell(flexcpText, mRow, 2) = CStr(mYearID) & "-" & CStr(mYearID + 1)
                If mPeriod = 2 Then
                    vsGrid.Cell(flexcpText, mRow, 3) = "IInd Half"
                Else
                    vsGrid.Cell(flexcpText, mRow, 3) = "Ist Half"
                End If
                vsGrid.MergeCol(12) = True
                vsGrid.Cell(flexcpText, mRow, 12) = mLoop 'Rec!numDemandID
                vsGrid.Cell(flexcpChecked, mRow, 12) = 1
                vsGrid.Cell(flexcpText, mRow, 6) = objAcc.AccountHeadID  'Rec!intAccountHeadID
                vsGrid.Cell(flexcpText, mRow, 7) = mYearID 'Rec!intYearID
                vsGrid.Cell(flexcpText, mRow, 8) = mPeriod 'Rec!tnyPeriodID
                vsGrid.Cell(flexcpText, mRow, 9) = IIf(mYearID > gbFinancialYearID, 1, 0) 'Rec!tnyArrearFlag
                vsGrid.Cell(flexcpText, mRow, 10) = mLoop 'Rec!numDemandID
                vsGrid.Cell(flexcpText, mRow, 11) = val(txtHalfYearTaxRate) * 5 / 100 'Rec!fltAmount
                If mYearID > Month(gbFinancialYearID) Then
                    mAmtArrear = mAmtArrear + val(txtHalfYearTaxRate.Text)
                    vsGrid.Cell(flexcpText, mRow, 4) = val(txtHalfYearTaxRate) * 5 / 100
                Else
                    mAmtCurrent = mAmtCurrent + val(txtHalfYearTaxRate.Text)
                    vsGrid.Cell(flexcpText, mRow, 5) = val(txtHalfYearTaxRate) * 5 / 100
                End If
                
                If mPeriod = 2 Then
                    mPeriod = 1
                    mYearID = mYearID + 1
                Else
                    mPeriod = 2
                End If
            Next
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
    
    Private Function Fine(mYearID As Integer, mPeriodID As Integer, mUptoDate As Date, mPTax As Double) As Double
        '==============================================================================='
        ' Modified By : Aiby                                                            '
        '             : For                                        '
        '==============================================================================='
        
        Dim dtFromDt As Variant
        Dim mNoOfMonths As Long
        Dim mAmount     As Double
        
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation Mode 1= Act and 2 = Circular                          '
        '-------------------------------------------------------------------------------'
        If gbFineCalculationMode = 1 Then
            If mPeriodID = 1 Then
                dtFromDt = DateSerial(mYearID, 10, 1)
            Else
                dtFromDt = DateSerial(mYearID + 1, 4, 1)
            End If
            
            If mYearID = gbFinancialYearID And mPeriodID = 2 Then
                Fine = 0
                Exit Function
            End If
            
            If mYearID < 2006 Then
                If mYearID = 2005 And mPeriodID = 2 Then
                    GoTo Skip
                End If
                If mUptoDate > DateSerial(2005, 9, 1) Then
                    'mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 4, 1), dtFromDt)) * 2 + 10
                    mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 9, 1), dtFromDt)) * 2
                    mNoOfMonths = mNoOfMonths + 1
                    dtFromDt = DateSerial(2005, 10, 1)
                    mYearID = 2005
                    mPeriodID = 2
                Else
                    mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2
                    dtFromDt = mUptoDate
                    mYearID = Year(dtFromDt)
                    If Month(dtFromDt) > 9 And Month(dtFromDt) < 4 Then
                        mPeriodID = 2
                    Else
                        mPeriodID = 1
                    End If
                End If
                
            
                'If Year(mUptoDate) = 2005 Then
'                If mYearID = 2005 Then
'                    If mPeriodID = 1 Then
'                        mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2 + 10
'                        If Month(mUptoDate) > 5 Then
'                            mNoOfMonths = mNoOfMonths - ((Month(mUptoDate) - 5) * 12)
'                        End If
'                    Else
'                        GoTo Skip:
'                    End If
                'End If
                'If Year(mUptoDate) < 2005 Then 'New Change For UptoDate
                'Else
                'If mYearID < 2005 Then 'New Change For UptoDate
                '    mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 5, 1), dtFromDt)) * 2 + 10
                '    dtFromDt = DateSerial(2005, 11, 1)
                'End If
                'Else
                 '   mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2 + 10
                'End If
                
                
                
            End If
Skip:
            If mUptoDate >= dtFromDt Then
                'mNoOfMonths = mNoOfMonths + (gbFinancialYearID - mYearID) * 12 'New Change For UptoDate
                mNoOfMonths = mNoOfMonths + 1 + Abs(DateDiff("M", mUptoDate, dtFromDt))  'New Change For UptoDate
            End If
            
            If mYearID = gbFinancialYearID And mPeriodID = 1 Then
                'mNoOfMonths = mNoOfMonths - 1
            End If
            'mNoOfMonths = mNoOfMonths + 1
            dtFromDate = DateAdd("m", 1, mUptoDate)
            'Debug.Print "No of Months (Fine) " & mNoOfMonths
            Fine = mPTax * mNoOfMonths / 100
            'If mNoOfMonths = 60 Then Stop
            Debug.Print "No of Months (Fine) " & mNoOfMonths & "    " & Fine
            Exit Function
        ElseIf gbFineCalculationMode = 2 Then
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation As Per Circular                                       '
        '-------------------------------------------------------------------------------'
           'mPTax = Format(mPTax * 2, "0.00")
            dtFromDt = DateSerial(mYearID, 11, 1)
            If mYearID = gbFinancialYearID Then
                Fine = 0
                Exit Function
            End If
            If mYearID < 2005 Then
                mNoOfMonths = Abs(DateDiff("m", DateSerial(2005, 8, 1), dtFromDt))
                dtFromDt = DateSerial(2005, 9, 1)
                mNoOfMonths = mNoOfMonths + Abs(DateDiff("m", gbTransactionDate, dtFromDt))
            End If
            mNoOfMonths = mNoOfMonths + Abs(DateDiff("m", gbTransactionDate, dtFromDt)) + 1
            Fine = mPTax * mNoOfMonths / 100
            Exit Function
        End If
    End Function
    
    Function FindUptoDate(Optional index As Integer = 1) As Date
        Dim mL As Integer
        Dim dtDate As Date
        For mL = index To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mL, 0) = gbAcHeadCodeAdvancePTax Then
                dtDate = vsGrid.TextMatrix(mL, 16)
                FindUptoDate = dtDate
                Return
            End If
        Next
        FindUptoDate = gbTransactionDate
        Return
    End Function
    
    Private Sub SetAdvanceAmt()
        'Can Remove this Subroutine
        Dim mLoop As Integer
        For mLoop = mAdvCheckedRow + 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeAdvancePTax Then
                'Note:- Find Fine Upto Which Date - Upto the Advance date
                mAdvCheckedRow = mLoop
                If IsDate(vsGrid.TextMatrix(mLoop, 16)) Then
                    dtUptoDate = DateSerial(Year(vsGrid.TextMatrix(mLoop, 16)), Month(vsGrid.TextMatrix(mLoop, 16)), 1)
                Else
                    If val(vsGrid.TextMatrix(mLoop, 8)) = 1 Then ' Checking with Period ID -> If First Period
                        dtUptoDate = DateSerial(val(vsGrid.TextMatrix(mLoop, 7)), 4, 1)
                    Else
                        dtUptoDate = DateSerial(val(vsGrid.TextMatrix(mLoop, 7)), 10, 1)
                    End If
                End If
                mAdvAmt = mAdvAmt + val(vsGrid.TextMatrix(mLoop, 11))
                Exit Sub
            End If
        Next
        mAnyAdvanceFlag = False
        If mLoop = vsGrid.Rows Then
            'Note:-No Advance Found, There for, Current Date will be set
            dtUptoDate = DateSerial(Year(gbTransactionDate), Month(gbTransactionDate), 1)
        End If
    End Sub
    
    Private Sub GetAdvanceAmt()
        Dim mLoop As Integer
        For mLoop = mAdvCheckedRow + 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeAdvancePTax Then
                'Note:- Find Fine Upto Which Date - Upto the Advance date
                mAdvCheckedRow = mLoop
                If vsGrid.Cell(flexcpChecked, mLoop, 12) = 2 Then
                    vsGrid.Cell(flexcpChecked, mLoop, 12) = 1
                End If
                'If IsDate(vsGrid.TextMatrix(mLoop, 16)) Then
                '    dtUptoDate = DateSerial(Year(vsGrid.TextMatrix(mLoop, 16)), Month(vsGrid.TextMatrix(mLoop, 16)), 1)
                'Else
                '    If Val(vsGrid.TextMatrix(mLoop, 8)) = 1 Then ' Checking with Period ID -> If First Period
                '        dtUptoDate = DateSerial(Val(vsGrid.TextMatrix(mLoop, 7)), 4, 1)
                '    Else
                '        dtUptoDate = DateSerial(Val(vsGrid.TextMatrix(mLoop, 7)), 10, 1)
                '    End If
                'End If
                If val(vsGrid.TextMatrix(mLoop, 8)) = 1 Then
                    dtUptoDate = DateSerial(val(vsGrid.TextMatrix(mLoop, 7)), 3, 1)
                Else
                    dtUptoDate = DateSerial(val(vsGrid.TextMatrix(mLoop, 7)), 10, 1)
                End If
                mAdvAmt = mAdvAmt + val(vsGrid.TextMatrix(mLoop, 11))
                Exit Sub
            End If
        Next
        mAnyAdvanceFlag = False
        If mLoop = vsGrid.Rows Then
            'Note:-No Advance Found, There for, Current Date will be set
            dtUptoDate = DateSerial(Year(gbTransactionDate), Month(gbTransactionDate), 1)
        End If
    End Sub
    Private Function calculateFine() As Double
        Dim mLoop As Integer
        Dim mLoopCrl As Integer ' Act as a Static Variable
        Dim mFineAmt As Double  ' Total Fine Amount
        Dim mPTax    As Double  ' Total Arrear Property Tax
        Dim mLC      As Double
        Dim mCess    As Double
        Dim mPartAmt As Double  ' Total Ptax+LC+Cess after adjusting Advance Amount
       'Dim dtUptoDate As Date  ' Fine Upto Date
        Dim dtDemandDate As Date
        Dim mFine    As Double
        
        mAdvAmt = 0
        dtUptoDate = gbTransactionDate
        mAnyAdvanceFlag = True
        mLoopCrl = 1
        mAdvCheckedRow = 0
        mPartAmt = 0
        mFineAmt = 0
        
        For mLoop = mLoopCrl To vsGrid.Rows - 1
               
               'Note:- Geting Advance if any :: Seting dtUptoDate
                
                If mAnyAdvanceFlag Then
                    If mAdvAmt <= 0 Then
                        Call GetAdvanceAmt
                    End If
                End If
                
               'Note: Not Arrear Property Tax Or Row is not selecte
               '      in both this case it skips the loop body
                If ((vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePropertyTaxArrear Or _
                     vsGrid.TextMatrix(mLoop, 0) = mAcHeadCodePTaxNonResArrear Or _
                     vsGrid.TextMatrix(mLoop, 0) = mAcHeadCodePTaxNonResCurrent Or _
                     vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePropertyTaxCurrent _
                     ) _
                     And vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked) = False Then  ' [IF-1]
                    GoTo GoNext:
                End If
                mPartAmt = 0
                'Note:- Finding Demand Date
                 If val(vsGrid.TextMatrix(mLoop, 8)) = 1 Then
                     dtDemandDate = DateSerial(val(vsGrid.TextMatrix(mLoop, 7)), 4, 1)
                 Else
                     dtDemandDate = DateSerial(val(vsGrid.TextMatrix(mLoop, 7)), 10, 1)
                 End If
                 
                 'If dtUptoDate >= dtDemandDate Then ' [1] dtUptoDate > dtDemandDate
                        
                        'Note:- Finding Property Tax/LC/Cess
                         mPTax = val(vsGrid.TextMatrix(mLoop, 11))
                         'Note:- In Next two rows expecting LC and Cess
                         If vsGrid.Rows - 1 >= mLoop + 1 Then
                         If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeLibraryCess Or vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodePoorHomeCess Then
                             If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeLibraryCess Then
                                 mLC = val(vsGrid.TextMatrix(mLoop + 1, 11))
                             Else
                                 mCess = val(vsGrid.TextMatrix(mLoop + 1, 11))
                             End If
                         End If
                         End If
                         
                         'Note:- Cess
                         If vsGrid.Rows - 1 >= mLoop + 2 Then
                         If vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodeLibraryCess Or vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodePoorHomeCess Then
                             If vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodeLibraryCess Then
                                 mLC = val(vsGrid.TextMatrix(mLoop + 2, 11))
                             Else
                                 mCess = val(vsGrid.TextMatrix(mLoop + 2, 11))
                             End If
                         End If
                         End If
                        'Note:- End of Block: Finding Property Tax/LC/Cess
FindNextAdvance:
                         'Note:- Find If any Advance And Set new dtUptoDate
                         If mAdvAmt <= 0 Then
                             If mAnyAdvanceFlag Then Call GetAdvanceAmt
                         End If
                         
                        If dtUptoDate >= dtDemandDate Then ' [1] dtUptoDate > dtDemandDate
                                'Note:- Calculate Fine
                                If mPartAmt <= 0 Then
                                    mFine = Format(Fine(val(vsGrid.TextMatrix(mLoop, 7)), val(vsGrid.TextMatrix(mLoop, 8)), dtUptoDate, mPTax), "0.00")
                                    mFineAmt = mFineAmt + mFine
                                    mFineAmt = Format(mFineAmt, "0")
                                Else
                                   If Month(dtFromDate) < 10 And Month(dtFromDate) > 3 Then
                                       mFine = Format(Fine(Year(dtFromDate), 1, dtUptoDate, mPTax), "0.00")
                                   Else
                                       mFine = Format(Fine(Year(dtFromDate), 2, dtUptoDate, mPTax), "0.00")
                                   End If
                                   mFineAmt = mFineAmt + mFine
                                End If
                        Else
                           dtFromDate = dtDemandDate
                           'MsgBox "End of [1] dtUptoDate > dtDemandDate"
                        End If ' End of [1] dtUptoDate > dtDemandDate
                 
                        'Note:- Setting of Advance
                         While (mFine > 0 And mAdvAmt > 0)
                             If mFine <= mAdvAmt Then
                                 mAdvAmt = mAdvAmt - mFine
                                 mFine = 0
                             Else
                                 mFine = mFine - mAdvAmt
                                 mAdvAmt = 0
                                 If mAnyAdvanceFlag Then Call GetAdvanceAmt '::: Gets Any other Advance Exists also sets UptoDate
                             End If
                         Wend
                         
                         If mAdvAmt > 0 Then ' [IF-3]
                             If (mPTax + mLC + mCess) <= mAdvAmt Then
                                 'No need to calculate Fine
                                 mAdvAmt = mAdvAmt - (mPTax + mLC + mCess)
                                 mPTax = 0
                                 mLC = 0
                                 mCess = 0
                             Else '(mPTax + mLC + mCess) > mAdvAmt
                                 mPartAmt = 0
                                 If mCess > 0 Then
                                     'NOTE:- Not completed!!! - Aiby/Dated:16-Sep-2009
                                     'This part should be changed further to seperate to find Cess and LC
                                     mPartAmt = (mPTax + mLC) - mAdvAmt
                                     mPTax = Format(mPartAmt * 100 / 105, "0.00")
                                     mLC = mPartAmt - mPTax
                                 Else
                                     mPartAmt = (mPTax + mLC) - mAdvAmt
                                     mPTax = Format(mPartAmt * 100 / 105, "0.00")
                                     mLC = mPartAmt - mPTax
                                 End If
                                 mAdvAmt = 0
                                 If mPartAmt > 0 Then
                                    'dtFromDate = dtUptoDate
                                    GoTo FindNextAdvance:
                                End If
                             End If
                         End If ' [IF-3]
                 
GoNext:
        Next
        calculateFine = mFineAmt
    End Function
    
    Private Function CalculateFine_ToFixError() As Double
        Dim mLoop As Integer
        Dim mLoopCrl As Integer ' Act as a Static Variable
        Dim mFineAmt As Double  ' Total Fine Amount
        Dim mPTax    As Double  ' Total Arrear Property Tax
        Dim mLC      As Double
        Dim mCess    As Double
        Dim mPartAmt As Double  ' Total Ptax+LC+Cess after adjusting Advance Amount
       'Dim dtUptoDate As Date  ' Fine Upto Date
        
        mAdvAmt = 0
        dtUptoDate = gbTransactionDate
        mAnyAdvanceFlag = True
        mLoopCrl = 1
        mAdvCheckedRow = 0
        
        For mLoop = mLoopCrl To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, mLoop, 12) = 2 Then
                    GoTo NextLoop:
                End If
                
                If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeAdvancePTax Then
                    If mAnyAdvanceFlag Then Call SetAdvanceAmt
                End If
                
                If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePropertyTaxArrear Then ' [IF-1]
                        mPTax = val(vsGrid.TextMatrix(mLoop, 11))
                        'Note:- In Next two rows expecting LC and Cess
                        If vsGrid.Rows - 1 >= mLoop + 1 Then
                        If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeLibraryCess Or vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodePoorHomeCess Then
                            If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeLibraryCess Then
                                mLC = val(vsGrid.TextMatrix(mLoop + 1, 11))
                            Else
                                mCess = val(vsGrid.TextMatrix(mLoop + 1, 11))
                            End If
                        End If
                        End If
                        
                        'Note:- Cess
                        If vsGrid.Rows - 1 >= mLoop + 2 Then
                        If vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodeLibraryCess Or vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodePoorHomeCess Then
                            If vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodeLibraryCess Then
                                mLC = val(vsGrid.TextMatrix(mLoop + 2, 11))
                            Else
                                mCess = val(vsGrid.TextMatrix(mLoop + 2, 11))
                            End If
                        End If
                        End If
SetOffAdvance:
                        'Note:- Check Where there is Any Fine Amount
                        If mAdvAmt > 0 Then ' [IF-2]
calculateFine:
                            mFineAmt = mFineAmt + Format(Fine(val(vsGrid.TextMatrix(mLoop, 7)), val(vsGrid.TextMatrix(mLoop, 8)), dtUptoDate, mPTax), "0.00")
                            While (mFineAmt > 0 And mAdvAmt > 0)
                                If mFineAmt <= mAdvAmt Then
                                    mAdvAmt = mAdvAmt - mFineAmt
                                    mFineAmt = 0
                                Else
                                    mFineAmt = mFineAmt - mAdvAmt
                                    mAdvAmt = 0
                                    Call SetAdvanceAmt
                                End If
                            Wend
                            
                            If mAdvAmt > 0 Then ' [IF-3]
                                If (mPTax + mLC + mCess) <= mAdvAmt Then
                                    'No need to calculate Fine
                                    mAdvAmt = mAdvAmt - (mPTax + mLC + mCess)
                                    mPTax = 0
                                    mLC = 0
                                    mCess = 0
                                Else '(mPTax + mLC + mCess) > mAdvAmt
                                    If mCess > 0 Then
                                        'NOTE:- Not completed!!! - Aiby/Dated:16-Sep-2009
                                        'This part should be changed further to seperate to find Cess and LC
                                        mPartAmt = (mPTax + mLC) - mAdvAmt
                                        mPTax = Format(mPartAmt * 100 / 105, "0.00")
                                        mLC = mPartAmt - mPTax
                                    Else
                                        mPartAmt = (mPTax + mLC) - mAdvAmt
                                        mPTax = Format(mPartAmt * 100 / 105, "0.00")
                                        mLC = mPartAmt - mPTax
                                    End If
                                    mAdvAmt = 0
                                    Call SetAdvanceAmt
                                    'GoTo CalculateFine: ':::: Changed for Fix ReCalculation of Fine after adjustment.
                                End If
                            End If ' [IF-3]
                        Else ' [IF-2 -ELSE] '
                            If mAnyAdvanceFlag Then
                                Call SetAdvanceAmt
                                If mAdvAmt > 0 Then
                                    GoTo SetOffAdvance:
                                Else
                                    'Note:- Find Fine
                                    If mPTax > 0 Then
                                        mFineAmt = mFineAmt + Format(Fine(val(vsGrid.TextMatrix(mLoop, 7)), val(vsGrid.TextMatrix(mLoop, 8)), dtUptoDate, mPTax), "0.00")
                                    End If
                                End If
                            Else
                                'Note:- Find Fine
                                If mPTax > 0 Then
                                    mFineAmt = mFineAmt + Format(Fine(val(vsGrid.TextMatrix(mLoop, 7)), val(vsGrid.TextMatrix(mLoop, 8)), dtUptoDate, mPTax), "0.00")
                                End If
                            End If
                        End If ' End Of [IF-2]
                End If ' End Of [IF-1] ::If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePropertyTaxArrear Then ' [IF-1]
NextLoop:
        Next
        CalculateFine_ToFixError = mFineAmt
    End Function
    
    
    
    
    
    
    Private Function CalculatePTaxFine(numBuildingID As Double, mYearID As Long, mPeriodID As Long) As Double
'        'CalculatePTaxFine(numBuildingID As Double, numDemandID As Double) As Double
'        Dim objDb As New clsDB
'        Dim mCnn As New ADODB.Connection
'        Dim RecIDemand As New ADODB.Recordset
'        Dim RecAdv As New ADODB.Recordset
'        Dim mSQL As String
'
'        Dim mAdvAmt As Double
'        Dim mFineAmt As Double
'        Dim mTotalFine As Double
'        Dim mPTAmt As Double
'        Dim mPTRate As Single
'        Dim mFromDate As Date
'        Dim mToDate As Date
'        Dim mNote As String
'
'        mAdvAmt = 0
'        mFineAmt = 0
'        mTotalFine = 0
'        mPTAmt = 0
'        mPTRate = 1
'        objDb.SetConnection mCnn
'
'        mSQL = ""
'        mSQL = mSQL + " Select faIDemandChild.numDemandID, faIDemandChild.dtOnDate, faIDemandChild.fltAmount, numSubLedgerID"
'        mSQL = mSQL + " ,faIDemandTbl.intYearID, faIDemandTbl.tnyPeriodID"
'        mSQL = mSQL + " From faIDemandChild Inner Join"
'        mSQL = mSQL + " faIDemandTbl On faIDemandTbl.numDemandID = faIDemandChild.numDemandID"
'        mSQL = mSQL + " Where faIDemandTbl.tnyStatus = 0 And faIDemandTbl.intTransactionTypeID = " & mPTaxTransactionTypeID
'        mSQL = mSQL + " And faIDemandChild.vchAccountHeadCode = '" & mPTaxArrearHeadCode & "'"
'        mSQL = mSQL + " And faIDemandTbl.numSubLedgerID = " & numBuildingID
'        mSQL = mSQL + " And ( faIDemandTbl.intYearID < " & mYearID
'        mSQL = mSQL + " Or ( faIDemandTbl.intYearID = " & mYearID & " AND faIDemandTbl.tnyPeriodID = " & mPeriodID & " ) )"
'
'        'mSQL = mSQL + " And faIDemandTbl.numDemandID <= " & numDemandID
'        RecIDemand.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'        If Not (RecIDemand.EOF And RecIDemand.BOF) Then
'            mSQL = ""
'            mSQL = mSQL + " Select faIDemandChild.numDemandID, faIDemandChild.dtOnDate, faIDemandChild.fltAmount"
'            mSQL = mSQL + " From faIDemandChild Inner Join"
'            mSQL = mSQL + " faIDemandTbl On faIDemandTbl.numDemandID = faIDemandChild.numDemandID "
'            mSQL = mSQL + " Where faIDemandTbl.tnyStatus = 0 And faIDemandTbl.intTransactionTypeID = " & mPTaxTransactionTypeID
'            mSQL = mSQL + " And faIDemandChild.vchAccountHeadCode = '" & mPTaxAdvanceCollected & "' And faIDemandTbl.numSubLedgerID = " & RecIDemand!numSubLedgerID
'            RecAdv.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'        Else
'            CalculatePTaxFine = 0
'            Exit Function
'        End If
'        While Not RecIDemand.EOF
'            mPTAmt = RecIDemand!fltAmount
'            If IsDate(RecIDemand!dtOnDate) Then
'                mFromDate = RecIDemand!dtOnDate
'            Else
'                If RecIDemand!tnyPeriodID = 1 Then
'                    mFromDate = DateSerial(RecIDemand!intYearID, 4, 1)
'                Else
'                    mFromDate = DateSerial(RecIDemand!intYearID, 10, 1)
'                End If
'            End If
'            '->
'            mNote = mNote + DdMmmYy(mFromDate) + "  PTax : " + Format(mPTAmt, "0.00") + vbCrLf
'            While Not RecAdv.EOF
'                If mAdvAmt <= 0 Then
'                    mAdvAmt = RecAdv!fltAmount
'                    mToDate = RecAdv!dtOnDate
'                    '->
'                    mNote = mNote + DdMmmYy(mToDate) + "   Adv : " + Format(mPTAmt, "0.00") + vbCrLf
'                    GoTo CalculatFine:
'                Else
'CalculatFine:
'                    mFineAmt = CalculateFine(mFromDate, mToDate, mPTAmt, mPTRate)
'                    '->
'                    mNote = mNote + str(mFineAmt) & DdMmmYy(mFromDate) & "  " & DdMmmYy(mToDate) & str(mPTAmt) & str(mPTRate)
'                    mTotalFine = mTotalFine + mFineAmt
'                    If mAdvAmt >= mFineAmt Then
'                        mAdvAmt = mAdvAmt - mFineAmt
'                        mFineAmt = 0
'                    Else
'                        mFineAmt = mFineAmt - mAdvAmt
'                        mAdvAmt = 0
'                    End If
'                    If mAdvAmt >= mPTAmt Then
'                        mAdvAmt = mAdvAmt - mPTAmt
'                        mPTAmt = 0
'                    Else
'                        mPTAmt = mPTAmt - mAdvAmt
'                        mAdvAmt = 0
'                    End If
'                    If mAdvAmt > 0 Then
'                        GoTo ReadNextDemand:
'                    End If
'                    If mPTAmt > 0 Then
'                        mFromDate = mToDate
'                    End If
'                    RecAdv.MoveNext
'                End If
'            Wend
'            If mPTAmt > 0 Then
'                mToDate = gbTransactionDate
'                mFineAmt = CalculateFine(mFromDate, mToDate, mPTAmt, mPTRate)
'                mTotalFine = mTotalFine + mFineAmt
'            End If
'
'ReadNextDemand:
'            RecIDemand.MoveNext
'        Wend
'        RecIDemand.Close
'        Set RecIDemand = Nothing
'        CalculatePTaxFine = mTotalFine
    End Function
    
    Private Sub SetDefaultSettings()
        Dim objTranType As New clsTransactionType
        Dim objAc As New clsAccounts
        Dim objInstruments As New clsInstruments
        Dim objBank As New clsBank
        Dim mLoopCount As Long
        
        mFineRate = 1  ' Fine Rate = 1.%'
        If gbLBType = 3 Or gbLBType = 4 Then
            mAcHeadCodePTaxArrear = "431100200"
            mAcHeadCodeFine = "140200200"
            mAcHeadCodeRoundOff = "00000"
        Else 'PANCHAYATs
            mAcHeadCodePTaxArrear = "431100102"
            mAcHeadCodePTaxNonResArrear = "431100104"
            mAcHeadCodePTaxNonResCurrent = "431100103"
            mAcHeadCodeFine = "140200101"
            mAcHeadCodeRoundOff = "00000"
        End If
        
        mDefaultTransactionTypeID = gbDefaultTransactionTypeID 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultTransactionTypeID"))
        mDefaultAccountHeadCode = gbAcHeadCodeCash 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultAccountHeadCode")
        mDefaultInstrumentID = gbInstrumentCash 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultInstumentID"))
        mDefaultBankID = gbDefaultBankID 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultBankID"))
        mDefaultZoneID = gbnumZonalID 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultZone"))
        
        If mDefaultZoneID > 0 Then
            For mLoopCount = 0 To cmbZone.ListCount - 1
                If cmbZone.ItemData(mLoopCount) = mDefaultZoneID Then
                    cmbZone.ListIndex = mLoopCount
                    Exit For
                End If
            Next
        End If
        
        objTranType.SetTransactionType (mDefaultTransactionTypeID)
        objAc.SetAccountCode (mDefaultAccountHeadCode)
        objInstruments.SetInstrumentType (mDefaultInstrumentID)
        objBank.SetBankInfo (mDefaultBankID)
        mDefaultBankHeadCode = objBank.BankAccountHeadCode
    End Sub
    
    Private Sub DisplayGrid()
        '------------------------------------------------------------------------'
        ' Aiby :                                                                 '
        '       Subroutine will fetch property tax demands and calculate fine    '
        '       And evalute it with total amount in hand and select the demands  '
        '------------------------------------------------------------------------'
        Dim Rec         As New ADODB.Recordset
        Dim arrInput    As Variant
        Dim arrOutPut   As Variant
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objAcc      As New clsAccounts
        Dim mArrearFlag As Boolean
        Dim mAmtArrear  As Double
        Dim mAmtCurrent As Double
        Dim mRows       As Long
        Dim mFine       As Double
        Dim mTotalFine  As Double
        Dim mTotalAmt   As Double
        Dim mTotalCash  As Double
        Dim mDemandTotal As Double
        Dim mDemandID   As Double
        
        vsGrid.Rows = 1
        mRows = 1
        arrInput = Array(mBuildingID)
        mTotalCash = val(txtGrandTotal.Text)
        
        If objdb.SetConnection(mCnn) Then
            Rec.CursorLocation = adUseClient
            Set Rec = objdb.ExecuteSP("spGetPropertyTaxDemands", arrInput, , , mCnn, adCmdStoredProc)
            While Not Rec.EOF
                If mDemandID <> Rec!numDemandID Then
                    mDemandID = Rec!numDemandID
                    arrInput = Array(Rec!numDemandID)
                    Call objdb.ExecuteSP("spGetTotalAmountOfAnyDemand", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
                    If IsArray(arrOutPut) Then
                        mDemandTotal = arrOutPut(0, 0)
                    End If
                End If
                If Rec!vchAccountHeadCode = mPTaxArrearHeadCode Then
                    'mFine = CalculatePTaxFineDemandWise(mBuildingID, Rec!numDemandID)
                    'mFine = CalculateFineforPTax(
                    If mTotalCash <= mTotalAmt + mDemandTotal + mFine Then
                        GoTo InsertFineRow
                    End If
                    mTotalFine = mTotalFine + mFine
                End If
                vsGrid.Rows = vsGrid.Rows + 1
                objAcc.SetAccountID (Rec!intAccountHeadID)
                vsGrid.Cell(flexcpText, mRows, 0) = Rec!vchAccountHeadCode
                vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                vsGrid.Cell(flexcpText, mRows, 2) = str(Rec!intYearID) & " - " & str(Rec!intYearID + 1)
                Select Case Rec!tnyPeriodID
                    Case Is = 1: vsGrid.Cell(flexcpText, mRows, 3) = "Ist Half"
                    Case Is = 2: vsGrid.Cell(flexcpText, mRows, 3) = "IInd Half"
                    Case Is = 3: vsGrid.Cell(flexcpText, mRows, 3) = "Full Year"
                End Select
                vsGrid.Cell(flexcpText, mRows, 6) = Rec!intAccountHeadID
                vsGrid.Cell(flexcpText, mRows, 7) = Rec!intYearID
                vsGrid.Cell(flexcpText, mRows, 8) = Rec!tnyPeriodID
                vsGrid.Cell(flexcpText, mRows, 9) = Rec!tnyArrearFlag
                vsGrid.Cell(flexcpText, mRows, 10) = Rec!numDemandID
                vsGrid.Cell(flexcpText, mRows, 11) = Rec!fltAmount
                vsGrid.MergeCol(12) = True
                vsGrid.Cell(flexcpText, mRows, 12) = Rec!numDemandID
                vsGrid.Cell(flexcpChecked, mRows, 12) = 1
                mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), 0, Rec!tnyArrearFlag)
                If mArrearFlag Then
                    mAmtArrear = mAmtArrear + Rec!fltAmount
                    vsGrid.Cell(flexcpText, mRows, 4) = Rec!fltAmount
                Else
                    mAmtCurrent = mAmtCurrent + Rec!fltAmount
                    vsGrid.Cell(flexcpText, mRows, 5) = Rec!fltAmount
                End If
                mTotalAmt = mTotalAmt + Rec!fltAmount
                mRows = mRows + 1
                Rec.MoveNext
            Wend

InsertFineRow:
            If mTotalFine > 0 Then
                vsGrid.Rows = vsGrid.Rows + 1
                objAcc.SetAccountCode (mAcHeadCodeFine)
                vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
                vsGrid.Cell(flexcpText, mRows, 9) = 0
                vsGrid.Cell(flexcpText, mRows, 11) = mTotalFine
                mAmtCurrent = mAmtCurrent + mTotalFine
                vsGrid.Cell(flexcpText, mRows, 5) = mTotalFine
                mRows = mRows + 1
            End If
            
            lblTotalArrear = Format(mAmtArrear, "0.00")
            lblTotalCurrent.Caption = Format(mAmtCurrent, "0.00")
            txtAdvance.Text = Format(val(txtGrandTotal) - Format(mAmtArrear + mAmtCurrent, "0.00"), "0.00")
            If val(txtAdvance) > 0 Then
                txtAdvance.Visible = True
                lblAdvance.Visible = True
            Else
                txtAdvance.Visible = False
                lblAdvance.Visible = False
            End If
            txtGrandTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
            
        End If
    End Sub
    
    Private Sub Settings()
        mAcHeadCodePTaxArrear = "431100100"
        mAcHeadCodeFine = "140200000"
        mAcHeadCodeRoundOff = ""
        mFineRate = 1
    End Sub
    Private Function FindBalanceAfterFine(mUptoRow As Integer, mYearID As Integer, mPeriodID As Integer, mAdvAmt As Double) As Double
        Dim mLoop As Integer
        Dim mAmt As Double
        For mLoop = 1 To mUptoRow
            'Note:- Check Property Tax Head (Arrear or Current)
            '       And whether the Year and Period is same as Advance collected
            If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePropertyTaxCurrent Or vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePropertyTaxArrear Then
            If vsGrid.Cell(flexcpValue, mLoop, 7) = mYearID And vsGrid.Cell(flexcpValue, mLoop, 8) = mPeriodID Then
                If val(vsGrid.Cell(flexcpText, mLoop, 4)) > 0 Then
                    mAmt = val(vsGrid.Cell(flexcpText, mLoop, 4))
                Else
                    mAmt = val(vsGrid.Cell(flexcpText, mLoop, 5))
                End If
                mAmt = mAmt - mAdvAmt
                If mAmt < 0 Then mAmt = 0
                FindBalanceAfterFine = mAmt
                Exit For
            End If
            End If
        Next
    End Function
    
 Private Sub Calculate()
        Dim mLoop As Long
        Dim mArrearAmt As Double
        Dim mCurrentAmt As Double
        Dim mFine As Double
        Dim mNoOfMonths As Integer
        Dim dtFromDt As Date
        Dim objAcc As New clsAccounts
        Dim mAdvAmt As Variant
        Dim mAmtStr As String
        Dim mNoticeFee As Double
        Dim mLoopCnl As Long
        Dim mPI As Boolean
        Dim mNF As Boolean
        
        mPI = False
        mNF = False
        If vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = mAcHeadCodeFine Then
            vsGrid.RemoveItem (vsGrid.Rows - 1)
        End If
        If val(txtNoticeFee.Text) > 0 Then
            mNoticeFee = val(txtNoticeFee.Text)
            txtNoticeFee.Text = ""
        End If
        
        
        If chkFineWaiver.Value = 0 Then
            txtFine.Text = ""
            mFine = calculateFine
        Else
            If val(txtFine.Text) > 0 Then
                mFine = val(txtFine.Text)
            End If
        End If
        
        ''Added On 03 Nov 2016 For checking Penal interrest /Notice fee exiists
        For mLoopCnl = 1 To vsGrid.Rows - 1
            
            If vsGrid.Cell(flexcpChecked, mLoopCnl, 12) = vbChecked And vsGrid.Cell(flexcpText, mLoopCnl, 0) = mAcHeadCodeFine Then
                mPI = True
            End If
            If vsGrid.Cell(flexcpChecked, mLoopCnl, 12) = vbChecked And vsGrid.Cell(flexcpText, mLoopCnl, 0) = gbAcHeadCodeNoticeFee Then
                mNF = True
            End If
        Next
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked And vsGrid.Cell(flexcpText, mLoop, 0) = gbAcHeadCodeAdvancePTax Then
                ' Advance Amount Adjusted '
                If val(vsGrid.Cell(flexcpText, mLoop, 4)) > 0 Then
                    '---------------------------------------------------------------'
                    'To Sort out the Round off issues                               '
                    '---------------------------------------------------------------'
                    mAmtStr = Format(val(vsGrid.Cell(flexcpText, mLoop, 4)), "0.00")
                    vsGrid.Cell(flexcpText, mLoop, 4) = mAmtStr
                    mAdvAmt = mAdvAmt + val(vsGrid.Cell(flexcpText, mLoop, 4))
                Else
                    mAmtStr = Format(val(vsGrid.Cell(flexcpText, mLoop, 5)), "0.00")
                    vsGrid.Cell(flexcpText, mLoop, 5) = mAmtStr
                    mAdvAmt = mAdvAmt + val(vsGrid.Cell(flexcpText, mLoop, 5))
                End If
                vsGrid.TextMatrix(mLoop, 15) = Format(FindBalanceAfterFine(mLoop - 1, val(vsGrid.Cell(flexcpText, mLoop, 7)), val(vsGrid.Cell(flexcpText, mLoop, 8)), val(mAmtStr)), "0.00")
                GoTo NextRow
            End If
            If val(vsGrid.Cell(flexcpText, mLoop, 4)) > 0 Then
                If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
                    '---------------------------------------------------------------'
                    'To Sort out the Round off issues                               '
                    '---------------------------------------------------------------'
                    mAmtStr = Format(val(vsGrid.Cell(flexcpText, mLoop, 4)), "0.00")
                    vsGrid.Cell(flexcpText, mLoop, 4) = mAmtStr
                    mArrearAmt = mArrearAmt + val(vsGrid.Cell(flexcpText, mLoop, 4))
                    If vsGrid.Cell(flexcpText, mLoop, 0) = mAcHeadCodePTaxArrear Then
                        'mFine = mFine + CalculateFineforPTax(Val(vsGrid.Cell(flexcpText, mLoop, 7)), Val(vsGrid.Cell(flexcpText, mLoop, 8)), Val(vsGrid.Cell(flexcpText, mLoop, 4)))
                    End If
                End If
            Else
                If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
                    '---------------------------------------------------------------'
                    'To Sort out the Round off issues                               '
                    '---------------------------------------------------------------'
                    mAmtStr = Format(val(vsGrid.Cell(flexcpText, mLoop, 5)), "0.00")
                    vsGrid.Cell(flexcpText, mLoop, 5) = mAmtStr
                    mCurrentAmt = mCurrentAmt + val(vsGrid.Cell(flexcpText, mLoop, 5))
                End If
            End If
NextRow:
        Next mLoop
        
        If mFine > 0 Then
            If mPI = False Then
                vsGrid.Rows = vsGrid.Rows + 1
                objAcc.SetAccountCode (mAcHeadCodeFine)
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 0) = objAcc.AccountCode
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1) = objAcc.AccountHead
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 6) = objAcc.AccountHeadID
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 7) = gbFinancialYearID
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 9) = 0
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 11) = Format(mFine, "#0")
                vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 12) = vbChecked
            
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 5) = Format(mFine, "#0")
                txtFine.Text = Format(mFine, "#0")
                txtFine.Enabled = False
                If gbTransactionDate <= Format("15/Mar/2011") Then
                    If mFineWaiveFlag Then
                        'mDescription = "Penal Interest waived to an extend 50% of  total : " & Format(mFine, "0.00")
                        mDescription = "Penal Interest waived for Rs : " & Format(mFine, "#0")
                        mFine = 0
                        txtFine.Text = 0 'Format(mFine / 2, "0.00")
                        vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 5) = 0 'Format(mFine / 2, "#.00")
                    End If
                End If
                mCurrentAmt = mCurrentAmt + Format(mFine, "0.00")
            End If
        End If
        
        If mNoticeFee > 0 Then
            If mNF = False Then
                vsGrid.Rows = vsGrid.Rows + 1
                objAcc.SetAccountCode (gbAcHeadCodeNoticeFee)
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 0) = objAcc.AccountCode
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1) = objAcc.AccountHead
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 6) = objAcc.AccountHeadID
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 7) = gbFinancialYearID
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 9) = 0
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 11) = Format(mNoticeFee, "0.00")
                vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 12) = vbChecked
                vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 5) = Format(mNoticeFee, "0.00")
                
                mCurrentAmt = mCurrentAmt + Format(mNoticeFee, "0.00")
                mNoticeFee = 0
            End If
        End If
        
        lblTotalArrear.Caption = Format(mArrearAmt, "0.00")
        lblTotalCurrent.Caption = Format(mCurrentAmt, "0.00")
        lblGrandTotal.Caption = Format(mArrearAmt + mCurrentAmt, "0.00")
        If val(txtNoticeFee.Text) > 0 Then
            lblGrandTotal.Caption = Format(lblGrandTotal.Caption + val(txtNoticeFee.Text), "0.00")
        End If
        If mAdvAmt > 0 Then
            txtAdvance.Visible = True
            txtAdvance.Text = Format(mAdvAmt, "0.00")
        Else
            txtAdvance.Text = ""
        End If
        txtNetAmount.Text = Format(val(lblGrandTotal.Caption) - val(txtAdvance), "0.00")
        If val(txtNetAmount) <= 0 Then
            cmdCopy.Enabled = False
        Else
            cmdCopy.Enabled = True
        End If
    End Sub
    
    
    Private Sub DisplayBuildingTaxDemands(mBuildingID As Double)
        Dim arrInput        As Variant
        Dim Rec             As New ADODB.Recordset
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim mRows           As Long
        Dim objAcc          As New clsAccounts
        Dim mArrearFlag     As Integer
        Dim mAmtArrear      As Double
        Dim mAmtCurrent     As Double
        Dim mFineFromDate   As Date
        Dim mNoOfMonths     As Integer
        'Dim mFineAmt        As Double
        Dim mFineFlag       As Boolean
        Dim mFine4PT        As Double
        Dim mAmt            As Double
        Dim mDemandID       As Double
        Dim mLastYearID     As Long
        Dim mPeriodID       As Long
        Dim mLoop           As Long
        Dim objPTax         As New clsPTax
        Dim mFromYear       As Variant
        Dim mNextYear       As Variant
        Dim mAdvanceAmt     As Variant
        Dim mStringIn       As String
        Dim mStringOut      As String
        Dim mSplitCol()     As String
        Dim mSplitRow()     As String
        Dim mUrl            As String
        Dim mRCnt           As Integer
        Dim mSql            As String
        Dim client1         As New MSSOAPLib.SoapClient
        Dim objSOAP         As Variant
        Dim mBuildingWeb    As String
        Dim mArrOut         As String
        Dim mXmlStream      As New ADODB.Stream
        Dim mNoticeFee      As Double    '''31/Oct/2016
        
        mAdvCheckedRow = 0
        vsGrid.Rows = 1
        mNumberOfSelections = 0
        arrInput = Array(mBuildingID, gbFinancialYearID)
        '-----------------------------
        mStringIn = CStr(mBuildingID) & "~" & CStr(gbFinancialYearID)
        '------------------------------
        '------Anju
        '------------------------------------------
        'For Cochin Corporation
        '------------------------------------------
'''''''        If gbLocalBodyID = 169 Then  ''05-12-2015 by Anju for cochin
'''''''
'''''''                Dim xmlHttp As Object
'''''''                Set xmlHttp = CreateObject("MSXML2.xmlHttp")
'''''''                Dim mXmlString   As Variant
'''''''                Dim oRs As ADODB.Recordset
'''''''                Dim oNode As Object 'MSXML2.IXMLDOMNode
'''''''                Dim oSubNodes As Object 'MSXML2.IXMLDOMSelection
'''''''                Dim oDoc As Object
'''''''                Dim params As String
'''''''
'''''''                If mBuildingID Then
'''''''                    mBuildingID = CStr(IIf(IsNull(mBuildingID), "NA", mBuildingID))
'''''''                    mUrl = gbDefaultUrl + "/getDemandDtlsUTF16/" & mBuildingID
'''''''                    'xmlHttp.Open "POST", "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService/searchAssesmentDetails?searchParam=" & params, False
'''''''                    xmlHttp.Open "POST", mUrl, False
'''''''                    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
'''''''                    xmlHttp.send
'''''''
'''''''                    mXmlString = xmlHttp.responseText
'''''''                   ' mXmlString = Replace(mXmlString, "UTF-8", "UTF-16")
'''''''
'''''''                    Set oDoc = CreateObject("MSXML2.DOMDocument")
'''''''                    oDoc.async = False
'''''''                    oDoc.validateOnParse = False
'''''''                    If Not oDoc.LoadXml(mXmlString) Then
'''''''                        MsgBox "Error Loading"
'''''''                        Exit Sub
'''''''                    Else
'''''''                        'MsgBox "Sucess"
'''''''                    End If
'''''''
'''''''                    Set oRs = New ADODB.Recordset
'''''''                    Set oRs.ActiveConnection = Nothing
'''''''                    oRs.CursorLocation = adUseClient
'''''''                    oRs.LockType = adLockBatchOptimistic
'''''''
'''''''                    With oRs.Fields
'''''''                        .Append "uniqId", adInteger
'''''''                        .Append "headCode", adVarChar, 20
'''''''                        .Append "head", adVarChar, 50
'''''''                        .Append "yearId", adInteger
'''''''                        .Append "periodId", adVarChar, 20
'''''''                        .Append "amount", adInteger
'''''''                        .Append "arriarFlag", adInteger
'''''''                    End With
'''''''
'''''''                    oRs.Open
'''''''                    Dim mCnt As Integer
'''''''                    mCnt = 1
'''''''                    vsGrid.Rows = 2
'''''''                    For Each oNode In oDoc.selectNodes("/PropertyTaxVo/demandRegisters")
'''''''                     'For Each oNode In oDoc.selectNodes("/PropertyTaxVo/demandRegisters")
'''''''                        oRs.ADDNEW
'''''''                        vsGrid.Rows = vsGrid.Rows + 1
'''''''                        oRs.Fields("uniqId").value = oNode.selectSingleNode("uniqId").Text
'''''''                        oRs.Fields("headCode").value = oNode.selectSingleNode("headCode").Text
'''''''                        objAcc.SetAccountCode (oNode.selectSingleNode("headCode").Text)
'''''''                        If oNode.selectSingleNode("headCode").Text = gbAcHeadCodeAdvancePTax Then
'''''''                            mAdvanceAmt = mAdvanceAmt + Rec!fltAmount
'''''''                            vsGrid.Cell(flexcpText, mCnt, 14) = 1
'''''''                            'GoTo NextRecord:
'''''''                            mAdvanceExists = True
'''''''                        End If
'''''''                        'oNode.selectSingleNode("amount").Text = 100 'test
'''''''                        vsGrid.Cell(flexcpText, mCnt, 0) = oNode.selectSingleNode("headCode").Text
'''''''                        vsGrid.Cell(flexcpText, mCnt, 1) = objAcc.AccountHead
'''''''                        vsGrid.Cell(flexcpText, mCnt, 2) = str(oNode.selectSingleNode("yearId").Text) & " - " & str(oNode.selectSingleNode("yearId").Text + 1)
'''''''                        Select Case oNode.selectSingleNode("periodId").Text
'''''''                            Case Is = 1: vsGrid.Cell(flexcpText, mCnt, 3) = "Ist Half"
'''''''                            Case Is = 2: vsGrid.Cell(flexcpText, mCnt, 3) = "IInd Half"
'''''''                            Case Is = 3: vsGrid.Cell(flexcpText, mCnt, 3) = "Full Year"
'''''''                        End Select
'''''''                        vsGrid.MergeCol(12) = True
'''''''                        vsGrid.Cell(flexcpText, mCnt, 12) = str(oNode.selectSingleNode("yearId").Text) & "-" & str(oNode.selectSingleNode("periodId").Text) 'Rec!intKeyID 'Rec!numDemandID
'''''''                        vsGrid.Cell(flexcpChecked, mCnt, 12) = 1
'''''''                        vsGrid.Cell(flexcpText, mCnt, 6) = objAcc.AccountHeadID
'''''''                        vsGrid.Cell(flexcpText, mCnt, 7) = oNode.selectSingleNode("yearId").Text
'''''''                        vsGrid.Cell(flexcpText, mCnt, 8) = oNode.selectSingleNode("periodId").Text
'''''''                        vsGrid.Cell(flexcpText, mCnt, 9) = oNode.selectSingleNode("arriarFlag").Text
'''''''                        vsGrid.Cell(flexcpText, mCnt, 10) = oNode.selectSingleNode("uniqId").Text
'''''''                       ' gbDemandID = oNode.selectSingleNode("uniqId").Text
'''''''                        vsGrid.Cell(flexcpText, mCnt, 11) = oNode.selectSingleNode("amount").Text
'''''''                        vsGrid.Cell(flexcpText, mCnt, 13) = "" 'Rec!numBatchID
'''''''                        'vsGrid.Cell(flexcpText, mCnt, 16) = DdMmmYy(Rec!dtDemandDate)
'''''''                        mArrearFlag = IIf(IsNull(oNode.selectSingleNode("arriarFlag").Text), 0, oNode.selectSingleNode("arriarFlag").Text)
'''''''                        If mArrearFlag Then
'''''''                            '-------------------------------'
'''''''                            '       To Calculate the Fine   '
'''''''                            '-------------------------------'
'''''''                            'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
'''''''                            '-------------------------------'
'''''''                            mAmtArrear = mAmtArrear + oNode.selectSingleNode("amount").Text
'''''''                            vsGrid.Cell(flexcpText, mCnt, 4) = oNode.selectSingleNode("amount").Text
'''''''                            If objAcc.AccountCode = mAcHeadCodePTaxArrear Or objAcc.AccountCode = mAcHeadCodePTaxNonResArrear Then
'''''''                                mFineFlag = True
'''''''                            End If
'''''''                        Else
'''''''                            mAmtCurrent = mAmtCurrent + oNode.selectSingleNode("amount").Text
'''''''                            vsGrid.Cell(flexcpText, mCnt, 5) = oNode.selectSingleNode("amount").Text
'''''''                        End If
'''''''
'''''''                        If objAcc.AccountCode = gbAcHeadCodePropertyTaxArrear Or objAcc.AccountCode = gbAcHeadCodePropertyTaxCurrent Then
'''''''                            '-------------------------------'
'''''''                            '       To Calculate the Fine   '
'''''''                            '-------------------------------'
'''''''                             If chkFineWaiver.value = 0 Then
'''''''                             'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
'''''''                             End If
'''''''                            '-------------------------------'
'''''''                            mFine4PT = oNode.selectSingleNode("amount").Text
'''''''                        End If
'''''''
'''''''                        ''''''
'''''''                        oRs.Fields("head").value = oNode.selectSingleNode("head").Text
'''''''                        'vsGrid.Cell(flexcpText, mCnt, 3) = oNode.selectSingleNode("head").Text
'''''''                        oRs.Fields("yearId").value = oNode.selectSingleNode("yearId").Text
'''''''                        oRs.Fields("periodId").value = oNode.selectSingleNode("periodId").Text
'''''''                        oRs.Fields("amount").value = oNode.selectSingleNode("amount").Text
'''''''                        oRs.Fields("arriarFlag").value = oNode.selectSingleNode("arriarFlag").Text
'''''''                        mCnt = mCnt + 1
'''''''                    Next
'''''''
'''''''                    '------------------------------
'''''''
'''''''                     '-------------------------------'
'''''''                    '       To Calculate the Fine   '
'''''''                    '-------------------------------'
'''''''                    If chkFineWaiver.value = 0 Then
'''''''                        mFineAmt = calculateFine
'''''''                        'Debug.Print "Property Tax Fine : " & str(mFineAmt)
'''''''                    End If
'''''''                    If mFineAmt > 0 Then
'''''''                        vsGrid.Rows = vsGrid.Rows + 1
'''''''                        objAcc.SetAccountCode (mAcHeadCodeFine)
'''''''                        vsGrid.Cell(flexcpText, mCnt, 0) = objAcc.AccountCode
'''''''                        vsGrid.Cell(flexcpText, mCnt, 1) = objAcc.AccountHead
'''''''                        vsGrid.Cell(flexcpText, mCnt, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
'''''''                        vsGrid.Cell(flexcpText, mCnt, 6) = objAcc.AccountHeadID
'''''''                        vsGrid.Cell(flexcpText, mCnt, 7) = gbFinancialYearID
'''''''                        vsGrid.Cell(flexcpText, mCnt, 9) = 0
'''''''                        vsGrid.Cell(flexcpText, mCnt, 11) = Format(mFineAmt, "#0")
'''''''                        vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
'''''''                        mAmtCurrent = mAmtCurrent + Format(mFineAmt, "#0")
'''''''                        vsGrid.Cell(flexcpText, mCnt, 5) = Format(mFineAmt, "#0")
'''''''                        mCnt = mCnt + 1
'''''''                    End If
'''''''
'''''''                    lblTotalArrear = Format(mAmtArrear, "0.00")
'''''''                    lblTotalCurrent.Caption = Format(mAmtCurrent, "0.00")
'''''''                    lblGrandTotal.Caption = Format(mAmtArrear + mAmtCurrent, "0.00")
'''''''                    If mAdvanceAmt > 0 Then
'''''''                        txtAdvance.Text = Format(mAdvanceAmt, "0.00")
'''''''                    Else
'''''''                        txtAdvance.Text = ""
'''''''                    End If
'''''''
'''''''                    '--------------------------------------------------------------------'
'''''''                    ' Fine waving
'''''''                    '--------------------------------------------------------------------'
'''''''                    'If mLastYearID = gbFinancialYearID And mAmtArrear > 0 Then
'''''''                    If mLastYearID = gbFinancialYearID And mPeriodID = gbCurrentPeriodID And mAmtCurrent > 0 Then
'''''''                        mFineWaiveFlag = True
'''''''                    Else
'''''''                        mFineWaiveFlag = False
'''''''                    End If
'''''''                Else
'''''''                    MsgBox "No Demand found for this Building!", vbInformation
'''''''                End If
'''''''                On Error Resume Next
'''''''                txtFromYear.Text = vsGrid.TextMatrix(1, 2)
'''''''        Exit Sub
'-----Anju
       'ElseIf mDemandWeb = True Then
                
              If mDemandWeb = True Then
                mUrl = gbDefaultUrl
                
                Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                objSOAP.MSSoapInit mUrl + "?WSDL"
                
                mBuildingWeb = CStr(mBuildingID)
                mArrOut = objSOAP.getBuildingDemandSaankhyaXML(gbLBID, mBuildingWeb, gbLocationID)
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
'                        If Rec!intKeyID <> mDemandID Then
'                            '--------------------------------------------------'
'                            ' Clearing the Fine Flag and Demands               '
'                            '--------------------------------------------------'
'                             mFineFlag = False
'                             mFine4PT = 0
'                             mDemandID = Rec!intKeyID
'                             mLastYearID = Rec!intYearID
'                             mPeriodID = Rec!tnyPeriodID
'                        End If
                        '--------------------------------------------------'
                        ' Beging of Block - Inserting Demands in Rows      '
                        '--------------------------------------------------'
                        vsGrid.Rows = vsGrid.Rows + 1
                        'objAcc.SetAccountID (Rec!intAccountHeadID)
                        objAcc.SetAccountCode (Rec!HeadCode)
                        
                        If Rec!HeadCode = gbAcHeadCodeAdvancePTax Then
                            mAdvanceAmt = mAdvanceAmt + Rec!fltAmount
                            vsGrid.Cell(flexcpText, mRows, 14) = 1
                            'GoTo NextRecord:
                            mAdvanceExists = True
                        End If
                        
                        vsGrid.Cell(flexcpText, mRows, 0) = Rec!HeadCode
                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                        vsGrid.Cell(flexcpText, mRows, 2) = str(Rec!intYearID) & " - " & str(Rec!intYearID + 1)
                        Select Case Rec!tnyPeriodID
                            Case Is = 1: vsGrid.Cell(flexcpText, mRows, 3) = "Ist Half"
                            Case Is = 2: vsGrid.Cell(flexcpText, mRows, 3) = "IInd Half"
                            Case Is = 3: vsGrid.Cell(flexcpText, mRows, 3) = "Full Year"
                        End Select
                        vsGrid.MergeCol(12) = True
                        vsGrid.Cell(flexcpText, mRows, 12) = str(Rec!intYearID) & "-" & str(Rec!tnyPeriodID) 'Rec!intKeyID 'Rec!numDemandID
                        vsGrid.Cell(flexcpChecked, mRows, 12) = 1
                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                        vsGrid.Cell(flexcpText, mRows, 7) = Rec!intYearID
                        vsGrid.Cell(flexcpText, mRows, 8) = Rec!tnyPeriodID
                        vsGrid.Cell(flexcpText, mRows, 9) = Rec!ArrFlag
                        vsGrid.Cell(flexcpText, mRows, 10) = Rec!intKeyID
                        vsGrid.Cell(flexcpText, mRows, 11) = Rec!fltAmount
                        vsGrid.Cell(flexcpText, mRows, 13) = "" 'Rec!numBatchID
                        vsGrid.Cell(flexcpText, mRows, 16) = DdMmmYy(Rec!dtDemandDate)
                        mArrearFlag = IIf(IsNull(Rec!ArrFlag), 0, Rec!ArrFlag)
                        If mArrearFlag Then
                            '-------------------------------'
                            '       To Calculate the Fine   '
                            '-------------------------------'
                            'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                            '-------------------------------'
                            mAmtArrear = mAmtArrear + Rec!fltAmount
                            vsGrid.Cell(flexcpText, mRows, 4) = Rec!fltAmount
                            If objAcc.AccountCode = mAcHeadCodePTaxArrear Or objAcc.AccountCode = mAcHeadCodePTaxNonResArrear Then
                                mFineFlag = True
                            End If
                        Else
                            mAmtCurrent = mAmtCurrent + Rec!fltAmount
                            vsGrid.Cell(flexcpText, mRows, 5) = Rec!fltAmount
                        End If
                        
                        mRows = mRows + 1
                        If objAcc.AccountCode = gbAcHeadCodePropertyTaxArrear Or objAcc.AccountCode = gbAcHeadCodePropertyTaxCurrent Then
                            '-------------------------------'
                            '       To Calculate the Fine   '
                            '-------------------------------'
                             If chkFineWaiver.Value = 0 Then
                             'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                             End If
                            '-------------------------------'
                            mFine4PT = Rec!fltAmount
                        End If
'NextRecord:
                        Rec.MoveNext
                    Wend
                    '-------------------------------'
                    '       To Calculate the Fine   '
                    '-------------------------------'
                    If chkFineWaiver.Value = 0 Then
                        mFineAmt = calculateFine
                        'Debug.Print "Property Tax Fine : " & str(mFineAmt)
                    End If
                    If mFineAmt > 0 Then
                        vsGrid.Rows = vsGrid.Rows + 1
                        objAcc.SetAccountCode (mAcHeadCodeFine)
                        vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                        vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                        vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
                        vsGrid.Cell(flexcpText, mRows, 9) = 0
                        vsGrid.Cell(flexcpText, mRows, 11) = Format(mFineAmt, "#0")
                        vsGrid.Cell(flexcpChecked, mRows, 12) = vbChecked
                        mAmtCurrent = mAmtCurrent + Format(mFineAmt, "#0")
                        vsGrid.Cell(flexcpText, mRows, 5) = Format(mFineAmt, "#0")
                        mRows = mRows + 1
                    End If
                    
                    lblTotalArrear = Format(mAmtArrear, "0.00")
                    lblTotalCurrent.Caption = Format(mAmtCurrent, "0.00")
                    
                    If val(txtNoticeFee.Text) > 0 Then
                        mNoticeFee = val(txtNoticeFee.Text)
                    Else
                        mNoticeFee = 0
                    End If
                    
                    lblGrandTotal.Caption = Format(mAmtArrear + mAmtCurrent, "0.00")
                    lblGrandTotal.Caption = Format(val(lblGrandTotal.Caption) + mNoticeFee, "0.00")
                    If mAdvanceAmt > 0 Then
                        txtAdvance.Text = Format(mAdvanceAmt, "0.00")
                    Else
                        txtAdvance.Text = ""
                    End If
                    
                    '--------------------------------------------------------------------'
                    ' Fine waving
                    '--------------------------------------------------------------------'
                    'If mLastYearID = gbFinancialYearID And mAmtArrear > 0 Then
                    If mLastYearID = gbFinancialYearID And mPeriodID = gbCurrentPeriodID And mAmtCurrent > 0 Then
                        mFineWaiveFlag = True
                    Else
                        mFineWaiveFlag = False
                    End If
                Else
                    MsgBox "No Demand found for this Building!", vbInformation
                End If
                On Error Resume Next
                txtFromYear.Text = vsGrid.TextMatrix(1, 2)
                Rec.Close
                Set mCnn = Nothing
        '----------------------------------------------------------------------------------------'
        ' N o t e s:-                                                                            '
        '   If property DifferentZoneFlag is True then Connection Redirected to Main Office DB   '
        '   Other wise Normal Connection to the Sanchya Local Database                           '
        '----------------------------------------------------------------------------------------'
        
        ElseIf Right(gbLocationID, 2) = 1 Then ' --> [ If Main Office ]
                If mvarDifferentZoneFlag Then
                    If objdb.CreateNewConnection(mCnn, SanchayaHO) = False Then
                        mSql = "Didn't able to connect to the Main office Server"
                        MsgBox mSql, vbInformation
                        Exit Sub
                    End If
                Else
                    If objdb.CreateNewConnection(mCnn, SanchayaLite) = False Then
                        mSql = "Didn't able to connect to the Sanchaya Server"
                        MsgBox mSql, vbInformation
                        Exit Sub
                    End If
                End If
                '------------'
                ' Common Op  '
                '------------'
                Rec.CursorLocation = adUseClient
                'Set Rec = objdb.ExecuteSP("spSanGiveDemandToSaankhya", arrInput, , , mCnn, adCmdStoredProc)
                
                arrInput = Array(mBuildingID, gbLocationID)
                Set Rec = objdb.ExecuteSP("spSn_AdvancePenalCalcLB_S", arrInput, , , mCnn, adCmdStoredProc)
                If Not (Rec.BOF And Rec.EOF) Then
                    vsGrid.Rows = 1
                    mRows = 1
                    vsGrid.MergeCells = flexMergeFree
                    While Not Rec.EOF
'                        If Rec!intKeyID <> mDemandID Then
'                            '--------------------------------------------------'
'                            ' Clearing the Fine Flag and Demands               '
'                            '--------------------------------------------------'
'                             mFineFlag = False
'                             mFine4PT = 0
'                             mDemandID = Rec!intKeyID
'                             mLastYearID = Rec!intYearID
'                             mPeriodID = Rec!tnyPeriodID
'                        End If
                        '--------------------------------------------------'
                        ' Beging of Block - Inserting Demands in Rows      '
                        '--------------------------------------------------'
                        vsGrid.Rows = vsGrid.Rows + 1
                        'objAcc.SetAccountID (Rec!intAccountHeadID)
                        objAcc.SetAccountCode (Rec!HeadCode)
                        
                        If Rec!HeadCode = gbAcHeadCodeAdvancePTax Then
                            mAdvanceAmt = mAdvanceAmt + Rec!fltAmount
                            vsGrid.Cell(flexcpText, mRows, 14) = 1
                            'GoTo NextRecord:
                            mAdvanceExists = True
                        End If
                        
                        vsGrid.Cell(flexcpText, mRows, 0) = Rec!HeadCode
                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                        vsGrid.Cell(flexcpText, mRows, 2) = str(Rec!intYearID) & " - " & str(Rec!intYearID + 1)
                        Select Case Rec!tnyPeriodID
                            Case Is = 1: vsGrid.Cell(flexcpText, mRows, 3) = "Ist Half"
                            Case Is = 2: vsGrid.Cell(flexcpText, mRows, 3) = "IInd Half"
                            Case Is = 3: vsGrid.Cell(flexcpText, mRows, 3) = "Full Year"
                        End Select
                        vsGrid.MergeCol(12) = True
                        vsGrid.Cell(flexcpText, mRows, 12) = str(Rec!intYearID) & "-" & str(Rec!tnyPeriodID) 'Rec!intKeyID 'Rec!numDemandID
                        vsGrid.Cell(flexcpChecked, mRows, 12) = 1
                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                        vsGrid.Cell(flexcpText, mRows, 7) = Rec!intYearID
                        vsGrid.Cell(flexcpText, mRows, 8) = Rec!tnyPeriodID
                        vsGrid.Cell(flexcpText, mRows, 9) = Rec!ArrFlag
                        vsGrid.Cell(flexcpText, mRows, 10) = "" 'Rec!intKeyID
                        vsGrid.Cell(flexcpText, mRows, 11) = Rec!fltAmount
                        vsGrid.Cell(flexcpText, mRows, 13) = "" 'Rec!numBatchID
                        vsGrid.Cell(flexcpText, mRows, 16) = DdMmmYy(Rec!dtDemandDate)
                        mArrearFlag = IIf(IsNull(Rec!ArrFlag), 0, Rec!ArrFlag)
                        If mArrearFlag Then
                            '-------------------------------'
                            '       To Calculate the Fine   '
                            '-------------------------------'
                            'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                            '-------------------------------'
                            mAmtArrear = mAmtArrear + Rec!fltAmount
                            vsGrid.Cell(flexcpText, mRows, 4) = Rec!fltAmount
                            If objAcc.AccountCode = mAcHeadCodePTaxArrear Or objAcc.AccountCode = mAcHeadCodePTaxNonResArrear Then
                                mFineFlag = True
                            End If
                        Else
                            mAmtCurrent = mAmtCurrent + Rec!fltAmount
                            vsGrid.Cell(flexcpText, mRows, 5) = Rec!fltAmount
                        End If
                        
                        mRows = mRows + 1
                        If objAcc.AccountCode = gbAcHeadCodePropertyTaxArrear Or objAcc.AccountCode = gbAcHeadCodePropertyTaxCurrent Then
                            '-------------------------------'
                            '       To Calculate the Fine   '
                            '-------------------------------'
                             If chkFineWaiver.Value = 0 Then
                             'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                             End If
                            '-------------------------------'
                            mFine4PT = Rec!fltAmount
                        End If
'NextRecord:
                        Rec.MoveNext
                    Wend
                    '-------------------------------'
                    '       To Calculate the Fine   '
                    '-------------------------------'
                    If chkFineWaiver.Value = 0 Then
                        mFineAmt = calculateFine
                        'Debug.Print "Property Tax Fine : " & str(mFineAmt)
                    End If
                    If mFineAmt > 0 Then
                        vsGrid.Rows = vsGrid.Rows + 1
                        objAcc.SetAccountCode (mAcHeadCodeFine)
                        vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                        vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                        vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
                        vsGrid.Cell(flexcpText, mRows, 9) = 0
                        vsGrid.Cell(flexcpText, mRows, 11) = Format(mFineAmt, "#0")
                        vsGrid.Cell(flexcpChecked, mRows, 12) = vbChecked
                        mAmtCurrent = mAmtCurrent + Format(mFineAmt, "#0")
                        vsGrid.Cell(flexcpText, mRows, 5) = Format(mFineAmt, "#0")
                        mRows = mRows + 1
                    End If
                    
                    lblTotalArrear = Format(mAmtArrear, "0.00")
                    lblTotalCurrent.Caption = Format(mAmtCurrent, "0.00")
                    lblGrandTotal.Caption = Format(mAmtArrear + mAmtCurrent, "0.00")
                    If mAdvanceAmt > 0 Then
                        txtAdvance.Text = Format(mAdvanceAmt, "0.00")
                    Else
                        txtAdvance.Text = ""
                    End If
                    
                    '--------------------------------------------------------------------'
                    ' Fine waving
                    '--------------------------------------------------------------------'
                    'If mLastYearID = gbFinancialYearID And mAmtArrear > 0 Then
                    If mLastYearID = gbFinancialYearID And mPeriodID = gbCurrentPeriodID And mAmtCurrent > 0 Then
                        mFineWaiveFlag = True
                    Else
                        mFineWaiveFlag = False
                    End If
                Else
                    MsgBox "No Demand found for this Building!", vbInformation
                End If
                On Error Resume Next
                txtFromYear.Text = vsGrid.TextMatrix(1, 2)
                Rec.Close
                Set mCnn = Nothing
        Else  ' [ If Not Main Office Then ]
                If mvarDifferentZoneFlag Then ' [Y] BEGINING => Other Zonal Collection From ZONAL Office
                            '----------------------------------------------
                            'For Zonal Integration calling WebService for Connectivity
                            'Added on 16.08.09
                            '----------------------------------------------
                            'mUrl = ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultUrl")
                            'client1.mssoapinit mUrl + "?WSDL"
                            'mStringOut = (client1.GetDemand(mStringIn))
                            
                            
                            '----------------------------------------------------------------------'
                            'Changed By Aiby : To Support in WINDOWS 2000 Server
                            '----------------------------------------------------------------------'
                            mUrl = gbDefaultUrl 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultUrl")
                            Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                            objSOAP.MSSoapInit mUrl + "?WSDL"
                            mStringOut = (objSOAP.GetDemand(mStringIn))
                            '----------------------------------------------------------------------'
                            
                            
                            vsGrid.Rows = 1
                            mRows = 1
                            vsGrid.MergeCells = flexMergeFree
                            mSplitCol = Split(mStringOut, "#")
                            If Not IsMissing(mSplitCol) Then
                                For mRCnt = 0 To UBound(mSplitCol) - 1
                                        mSplitRow = Split(mSplitCol(mRCnt), "~")
                                        If mSplitRow(1) <> mDemandID Then
                                            '--------------------------------------------------'
                                            ' Clearing the Fine Flag and Demands               '
                                            '--------------------------------------------------'
                                            mFineFlag = False
                                            mFine4PT = 0
                                            mDemandID = mSplitRow(1)
                                            mLastYearID = mSplitRow(4)
                                            mPeriodID = mSplitRow(5)
                                        End If
                                        '--------------------------------------------------'
                                        ' Beging of Block - Inserting Demands in Rows      '
                                        '--------------------------------------------------'
                                        vsGrid.Rows = vsGrid.Rows + 1
                                        'objAcc.SetAccountID (Rec!intAccountHeadID)
                                        objAcc.SetAccountCode mSplitRow(10)
                                        If mSplitRow(10) = gbAcHeadCodeAdvancePTax Then
                                            mAdvanceAmt = mAdvanceAmt + mSplitRow(11)
                                            vsGrid.Cell(flexcpText, mRows, 14) = 1
                                            'GoTo NextRec:
                                        End If
                                        vsGrid.Cell(flexcpText, mRows, 0) = mSplitRow(10)
                                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                                        vsGrid.Cell(flexcpText, mRows, 2) = mSplitRow(4) & " - " & str(mSplitRow(4) + 1)
                                        Select Case mSplitRow(5)
                                            Case Is = 1: vsGrid.Cell(flexcpText, mRows, 3) = "Ist Half"
                                            Case Is = 2: vsGrid.Cell(flexcpText, mRows, 3) = "IInd Half"
                                            Case Is = 3: vsGrid.Cell(flexcpText, mRows, 3) = "Full Year"
                                        End Select
                                                
                                        vsGrid.MergeCol(12) = True
                                        vsGrid.Cell(flexcpText, mRows, 12) = mSplitRow(1)
                                        vsGrid.Cell(flexcpChecked, mRows, 12) = 1
                                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                                        vsGrid.Cell(flexcpText, mRows, 7) = mSplitRow(4)
                                        vsGrid.Cell(flexcpText, mRows, 8) = mSplitRow(5)
                                        vsGrid.Cell(flexcpText, mRows, 9) = mSplitRow(9)
                                        vsGrid.Cell(flexcpText, mRows, 10) = mSplitRow(1)
                                        vsGrid.Cell(flexcpText, mRows, 11) = val(mSplitRow(11))
                                        vsGrid.Cell(flexcpText, mRows, 13) = mSplitRow(0)
                                        mArrearFlag = IIf(IsNull(mSplitRow(9)), 0, mSplitRow(9))
                                        If mArrearFlag Then
                                            '-------------------------------'
                                            '       To Calculate the Fine   '
                                            '-------------------------------'
                                            'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                                            '-------------------------------'
                                            mAmtArrear = mAmtArrear + val(mSplitRow(11))
                                            vsGrid.Cell(flexcpText, mRows, 4) = val(mSplitRow(11))
                                            If objAcc.AccountCode = mAcHeadCodePTaxArrear Then mFineFlag = True
                                        Else
                                            mAmtCurrent = mAmtCurrent + val(mSplitRow(11))
                                            vsGrid.Cell(flexcpText, mRows, 5) = val(mSplitRow(11))
                                        End If
                                        mRows = mRows + 1
                                        If objAcc.AccountCode = mAcHeadCodePTaxArrear Then
                                            '-------------------------------'
                                            '       To Calculate the Fine   '
                                            '-------------------------------'
                                             mFineAmt = mFineAmt + CalculateFineforPTax(val(mSplitRow(4)), val(mSplitRow(5)), val(mSplitRow(11)))
                                            '-------------------------------'
                                            mFine4PT = val(mSplitRow(11))
                                        End If
                                Next mRCnt
                                Debug.Print "Property Tax Fine : " & str(mFineAmt)
                                mFineAmt = calculateFine
                                If mFineAmt > 0 Then
                                    vsGrid.Rows = vsGrid.Rows + 1
                                    objAcc.SetAccountCode (mAcHeadCodeFine)
                                    vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                                    vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                                    vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                                    vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                                    vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
                                    vsGrid.Cell(flexcpText, mRows, 9) = 0
                                    vsGrid.Cell(flexcpText, mRows, 11) = Format(mFineAmt, "#0")
                                    vsGrid.Cell(flexcpChecked, mRows, 12) = vbChecked
                                    mAmtCurrent = mAmtCurrent + Format(mFineAmt, "#0")
                                    vsGrid.Cell(flexcpText, mRows, 5) = Format(mFineAmt, "#0")
                                    mRows = mRows + 1
                                End If
                
                                lblTotalArrear = Format(mAmtArrear, "0.00")
                                lblTotalCurrent.Caption = Format(mAmtCurrent, "0.00")
                                lblGrandTotal.Caption = Format(mAmtArrear + mAmtCurrent, "0.00")
                                If mAdvanceAmt > 0 Then
                                    txtAdvance.Text = Format(mAdvanceAmt, "0.00")
                                Else
                                    txtAdvance.Text = ""
                                End If
                                '--------------------------------------------------------------------'
                                ' Fine waving
                                '--------------------------------------------------------------------'
'                                If mLastYearID = gbFinancialYearID And mAmtArrear > 0 Then
                                If mLastYearID = gbFinancialYearID And mPeriodID = gbCurrentPeriodID And mAmtCurrent > 0 Then
                                    mFineWaiveFlag = True
                                Else
                                    mFineWaiveFlag = False
                                End If
                            Else
                                MsgBox "No Demand found for this Building!", vbInformation
                                txtFromYear.Text = vsGrid.TextMatrix(1, 2)
                            End If 'NOTE:- If Not IsMissing(mSplitCol) Then
                
                     ' Note:- ELSE PART OF BLOCK [Y]
                Else ' Note:- Same Zone Collection [Y] Else Part Of Block [Y}
                            If objdb.CreateNewConnection(mCnn, SanchayaLite) = False Then
                                mSql = "Didn't able to connect to the Sanchaya Server"
                                MsgBox mSql, vbInformation
                                Exit Sub
                            End If
                            
                            Rec.CursorLocation = adUseClient
                            arrInput = Array(mBuildingID, gbLocationID)
                            Set Rec = objdb.ExecuteSP("spSn_AdvancePenalCalcLB_S", arrInput, , , mCnn, adCmdStoredProc)
                            
                            'Set Rec = objdb.ExecuteSP("spSanGiveDemandToSaankhya", arrInput, , , mCnn, adCmdStoredProc)
                            If Not (Rec.BOF And Rec.EOF) Then ' [Z] BEGINING
                                vsGrid.Rows = 1
                                mRows = 1
                                vsGrid.MergeCells = flexMergeFree
                                While Not Rec.EOF
'                                        If Rec!intKeyID <> mDemandID Then
'                                            '--------------------------------------------------'
'                                            ' Clearing the Fine Flag and Demands               '
'                                            '--------------------------------------------------'
'                                             mFineFlag = False
'                                             mFine4PT = 0
'                                             mDemandID = Rec!intKeyID
'                                             mLastYearID = Rec!intYearID
'                                             mPeriodID = Rec!tnyPeriodID
'                                        End If
                                        '--------------------------------------------------'
                                        ' Beging of Block - Inserting Demands in Rows      '
                                        '--------------------------------------------------'
                                        vsGrid.Rows = vsGrid.Rows + 1
                                        objAcc.SetAccountID (Rec!intAccountHeadID)
                                        objAcc.SetAccountCode (Rec!HeadCode)
                                        
                                        If Rec!HeadCode = gbAcHeadCodeAdvancePTax Then
                                            mAdvanceAmt = mAdvanceAmt + Rec!fltAmount
                                            vsGrid.Cell(flexcpText, mRows, 14) = 1
                                            'GoTo NextRecord:
                                        End If
                                        
                                        vsGrid.Cell(flexcpText, mRows, 0) = Rec!HeadCode
                                        vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                                        vsGrid.Cell(flexcpText, mRows, 2) = str(Rec!intYearID) & " - " & str(Rec!intYearID + 1)
                                        Select Case Rec!tnyPeriodID
                                            Case Is = 1: vsGrid.Cell(flexcpText, mRows, 3) = "Ist Half"
                                            Case Is = 2: vsGrid.Cell(flexcpText, mRows, 3) = "IInd Half"
                                            Case Is = 3: vsGrid.Cell(flexcpText, mRows, 3) = "Full Year"
                                        End Select
                                        
                                        vsGrid.MergeCol(12) = True
                                        vsGrid.Cell(flexcpText, mRows, 12) = str(Rec!intYearID) & "-" & str(Rec!tnyPeriodID) 'Rec!intKeyID 'Rec!numDemandID
                                        vsGrid.Cell(flexcpChecked, mRows, 12) = 1
                                        vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                                        vsGrid.Cell(flexcpText, mRows, 7) = Rec!intYearID
                                        vsGrid.Cell(flexcpText, mRows, 8) = Rec!tnyPeriodID
                                        vsGrid.Cell(flexcpText, mRows, 9) = Rec!ArrFlag
                                        vsGrid.Cell(flexcpText, mRows, 10) = "" 'Rec!intKeyID
                                        vsGrid.Cell(flexcpText, mRows, 11) = Rec!fltAmount
                                        vsGrid.Cell(flexcpText, mRows, 13) = "" 'Rec!numBatchID
                                        vsGrid.Cell(flexcpText, mRows, 16) = DdMmmYy(Rec!dtDemandDate)
                                        mArrearFlag = IIf(IsNull(Rec!ArrFlag), 0, Rec!ArrFlag)
                                        
                                        If mArrearFlag Then
                                            '-------------------------------'
                                            '       To Calculate the Fine   '
                                            '-------------------------------'
                                            'mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                                            '-------------------------------'
                                            mAmtArrear = mAmtArrear + Rec!fltAmount
                                            vsGrid.Cell(flexcpText, mRows, 4) = Rec!fltAmount
                                            If objAcc.AccountCode = mAcHeadCodePTaxArrear Then mFineFlag = True
                                        Else
                                            mAmtCurrent = mAmtCurrent + Rec!fltAmount
                                            vsGrid.Cell(flexcpText, mRows, 5) = Rec!fltAmount
                                        End If
                                        mRows = mRows + 1
                                        If objAcc.AccountCode = mAcHeadCodePTaxArrear Then
                                            '-------------------------------'
                                            '       To Calculate the Fine   '
                                            '-------------------------------'
                                             mFineAmt = mFineAmt + CalculateFineforPTax(Rec!intYearID, Rec!tnyPeriodID, Rec!fltAmount)
                                            '-------------------------------'
                                            mFine4PT = Rec!fltAmount
                                        End If
                                        Rec.MoveNext
                                Wend
                                Debug.Print "Property Tax Fine : " & str(mFineAmt)
                                mFineAmt = calculateFine
                                If mFineAmt > 0 Then
                                    vsGrid.Rows = vsGrid.Rows + 1
                                    objAcc.SetAccountCode (mAcHeadCodeFine)
                                    vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                                    vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                                    vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                                    vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                                    vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
                                    vsGrid.Cell(flexcpText, mRows, 9) = 0
                                    vsGrid.Cell(flexcpText, mRows, 11) = Format(mFineAmt, "#0")
                                    vsGrid.Cell(flexcpChecked, mRows, 12) = vbChecked
                                    mAmtCurrent = mAmtCurrent + Format(mFineAmt, "#0")
                                    vsGrid.Cell(flexcpText, mRows, 5) = Format(mFineAmt, "#0")
                                    mRows = mRows + 1
                                End If
                                
                                lblTotalArrear = Format(mAmtArrear, "0.00")
                                lblTotalCurrent.Caption = Format(mAmtCurrent, "0.00")
                                lblGrandTotal.Caption = Format(mAmtArrear + mAmtCurrent, "0.00")
                                If mAdvanceAmt > 0 Then
                                    txtAdvance.Text = Format(mAdvanceAmt, "0.00")
                                Else
                                    txtAdvance.Text = ""
                                End If
                                
                                '--------------------------------------------------------------------'
                                ' Fine waving
                                '--------------------------------------------------------------------'
'                                If mLastYearID = gbFinancialYearID And mAmtArrear > 0 Then
                                If mLastYearID = gbFinancialYearID And mPeriodID = gbCurrentPeriodID And mAmtCurrent > 0 Then
                                    mFineWaiveFlag = True
                                Else
                                    mFineWaiveFlag = False
                                End If
                            Else
                                MsgBox "No Demand found for this Building!", vbInformation
                            End If ' [Z] ENDING
                            On Error Resume Next
                            txtFromYear.Text = vsGrid.TextMatrix(1, 2)
                            Rec.Close
                            Set mCnn = Nothing
                End If '[Y} ENDING BLOCK [Y]
        End If
        mFineAmt = 0
    End Sub

    Private Sub FillAssessmentYear()
        Dim mSql As String
        On Error Resume Next
        mSql = "SELECT DISTINCT intWardYear,intWardYear as ID From GM_Ward Where tnyWardType = 1 AND intLBID = " & gbLocalBodyID & " ORDER BY intWardYear DESC"
        Call PopulateList(cmbAssessmentYear, mSql, , , , True, DBMaster)
        cmbAssessmentYear.ListIndex = 0
    End Sub
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
    
    
    Private Sub FormInitialize()
        mAdvanceExists = False
        Dim mCrl As Control
        
        
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.Value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
        lblRRFlag.Caption = ""
        'If gbLocalBodyID = 171 Then
         txtHouseNo1.MaxLength = 5
        'End If
        Call Settings
        If Month(Date) < 4 Then
            cmbToYear.Text = CStr(Year(Date) - 1) & "-" & CStr(Year(Date))
        Else
            cmbToYear.Text = CStr(Year(Date)) & "-" & CStr(Year(Date) + 1)
        End If
        
        cmbFromPeriod.Text = "First Half"
        If Month(Date) < 10 And Month(Date) > 3 Then
            cmbToPeriod.Text = "First Half"
        Else
            cmbToPeriod.Text = "Second Half"
        End If
        
        vsGrid.Rows = 1
        vsGrid.Rows = 10
        
        '---------------------------------------------------------'
        ' Initialize Property Tax related account heads and       '
        ' Transaction Types                                       '
        '---------------------------------------------------------'
        Dim objTranType As New clsTransactionType
                mPTaxTransactionTypeID = objTranType.GetTransactionTypeID("Property Tax")
        mPTaxArrearHeadCode = "431100200"
        mPTaxCurrentHeadCode = "431100100"
        mPTaxAdvanceCollected = "350410101"
        mPTaxLibraryCessCode = "350300100"
        
        Set objTranType = Nothing
        txtFromYear.Text = ""           '   Added Newly '
        txtNoOfHalfYears.Text = ""      '   Added Newly '
        txtWardNo.Text = "" '   Added Newly '
        mvarDifferentZoneFlag = False
        
        mSelectedAllFlag = False ' This Flag to Identify Whether Complete Demand Selected or not
        
        On Error Resume Next
        cmbZone.Text = gbLocation
        On Error GoTo 0
    End Sub
    
    Private Sub chkFineWaiver_Click()
        If chkFineWaiver.Value = 1 Then
            frmFineWaiver.Mode = 1
            frmFineWaiver.Show vbModal, frmPropertyTax
            Calculate
        Else
        
        End If
    End Sub

    Private Sub cmbAssessmentYear_Click()
        FillWard
    End Sub

    Private Sub cmbAssessmentYear_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            PressTabKey
        End If
    End Sub

    Private Sub cmbFromPeriod_Click()
        Call FindNumberOfHalfYears
    End Sub
    Private Sub cmbFromYear_Click()
        Call FindNumberOfHalfYears
    End Sub
    Private Sub cmbToPeriod_Click()
        Call FindNumberOfHalfYears
    End Sub
    Private Sub cmbToYear_Click()
        If cmbToYear.ListIndex > -1 Then Call FindNumberOfHalfYears
    End Sub

    

    Private Sub cmbWard_Click()
        If cmbWard.ListIndex > -1 Then
            txtWardNo.Text = cmbWard.ItemData(cmbWard.ListIndex)
        End If
    End Sub

    Private Sub cmbWard_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            PressTabKey
        End If
    End Sub

    Private Sub cmbZone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            PressTabKey
        End If
    End Sub
    Private Sub cmbZone_LostFocus()
        
        
        ' -----------------------------------------------------------------------'
        ' Added by Aiby on 08-April-2009                                         '
        ' In Case Zonal Id is different than Curret Location                     '
        ' Set Prorty varibale (mvarDifferentZoneFlag) to True                    '
        ' -----------------------------------------------------------------------'
        If cmbZone.ListIndex > -1 Then
            If cmbZone.ItemData(cmbZone.ListIndex) <> gbLocationID Then
                mvarDifferentZoneFlag = True
            Else
                mvarDifferentZoneFlag = False
            End If
        End If
        
        '-------------------------------------------------------------------------'
        ' F i l l   W a r d                                                       '
        '-------------------------------------------------------------------------'
        Call FillWard
        
        
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
    Private Sub cmdFind_Click()
        Dim arrInput        As Variant
        Dim Rec             As New ADODB.Recordset
        Dim objdb           As New clsDB
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
        '-------------------------------------------------------------'
        ' Changed to Change the Stored Procedure spGetBuildingDetails '
        ' New Stored Procedure is spSanGetSearchBuildingList          '
        '-------------------------------------------------------------'
        'arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
        '     mWardID, _
        '     Val(txtHouseNo1), _
        '     Trim(txtHouseNo2))
        '
        
        '--------------------------------------'
        ' Additions                            '
        '--------------------------------------'
        Dim numBuildingID   As Variant
        Dim numZoneID       As Variant
        Dim intAssessmentYear As Variant
        Dim intWardNo       As Variant
        Dim intDoorNo1      As Variant
        Dim chvDoorNo2      As Variant
        Dim chvName         As Variant
        Dim chvResHName     As Variant
        Dim mUrl            As String
        Dim mArrIn          As String
        Dim mArrOut         As String
        Dim mGetString()    As String
        Dim client          As New MSSOAPLib.SoapClient
        Dim objSOAP         As Variant
        'If Trim(txtBuildingNo) <> "" Then numBuildingID = Val(txtBuildingNo) Else numBuildingID = Null
        numBuildingID = Null
        If cmbZone.ListIndex > -1 Then numZoneID = cmbZone.ItemData(cmbZone.ListIndex) Else numZoneID = Null
        If cmbAssessmentYear.ListIndex > -1 Then intAssessmentYear = cmbAssessmentYear.ItemData(cmbAssessmentYear.ListIndex) Else intAssessmentYear = Null
        If cmbWard.ListIndex > -1 Then intWardNo = cmbWard.ItemData(cmbWard.ListIndex) Else intWardNo = Null
        If Trim(txtHouseNo1) <> "" Then intDoorNo1 = val(txtHouseNo1) Else intDoorNo1 = Null
        If Trim(txtHouseNo2) <> "" Then chvDoorNo2 = Trim(txtHouseNo2) Else chvDoorNo2 = Null
        chvName = Null
        chvResHName = Null
        
        arrInput = Array(numBuildingID, _
            numZoneID, _
            intAssessmentYear, _
            intWardNo, _
            intDoorNo1, _
            chvDoorNo2, _
            chvName, _
            chvResHName)
            
        mArrIn = CStr(IIf(IsNull(numBuildingID), "", numBuildingID)) & "~" & _
                 CStr(IIf(IsNull(numZoneID), "", numZoneID)) & "~" & _
                 CStr(IIf(IsNull(intAssessmentYear), "", intAssessmentYear)) & "~" & _
                 CStr(IIf(IsNull(intWardNo), "", intWardNo)) & "~" & _
                 CStr(IIf(IsNull(intDoorNo1), "", intDoorNo1)) & "~" & _
                 CStr(IIf(IsNull(chvDoorNo2), "", chvDoorNo2)) & "~" & _
                 CStr(IIf(IsNull(chvName), "", chvName)) & "~" & _
                 CStr(IIf(IsNull(chvResHName), "", chvResHName))
                        
 
            
            
            mvarBuildingID = -1
            Me.MousePointer = vbHourglass
            vsGrid.Clear 1, 0
            txtBuildingNo.Text = ""
            txtAddress.Text = ""
            txtGrandTotal.Text = ""
            lblGrandTotal.Caption = ""
            txtAdvance.Text = ""
            txtNetAmount.Text = ""
            lblTotalArrear.Caption = ""
            lblTotalCurrent.Caption = ""
            txtAdvance.Visible = False
             
            '-------------------------------------------------------
            '----------Added On 18 Sep 2015 Building Search From Tcs WebService
            ' Web Service  '
            ' To fetch data from Tcs Ptax
            'For Cochin Corporation
            '-------------------------------------------------------
''''''            If gbLocalBodyID = 169 Then     ''05-12-2015  Anju for cochin
''''''
''''''                Dim xmlHttp As Object
''''''                Set xmlHttp = CreateObject("MSXML2.xmlHttp")
''''''                Dim mXmlString   As Variant
''''''                Dim oRs As ADODB.Recordset
''''''                Dim oNode As Object 'MSXML2.IXMLDOMNode
''''''                Dim oSubNodes As Object 'MSXML2.IXMLDOMSelection
''''''                Dim oDoc As Object
''''''                Dim params As String
''''''                Dim mSqlWeb As String
''''''                Dim mRecWeb As New ADODB.Recordset
''''''                Dim mZoneNo As Integer
''''''                mSqlWeb = "SELECT tnyZoneNo FROM GM_Zone WHERE numZoneID=" & numZoneID
''''''                If objDB.CreateNewConnection(mCnn, enuSourceString.DBMaster) = True Then
''''''                    mRecWeb.Open mSqlWeb, mCnn, adOpenDynamic
''''''                    If Not (mRecWeb.BOF And mRecWeb.EOF) Then
''''''                       mZoneNo = IIf(IsNull(mRecWeb!tnyZoneNo), "", mRecWeb!tnyZoneNo)
''''''                    End If
''''''                    mRecWeb.Close
''''''                    mCnn.Close
''''''                End If
''''''                '----test
''''''                'intWardNo = 25
''''''                'intDoorNo1 = 1
''''''                '----test
''''''                'If cmbZone.ListIndex > -1 Then numZoneID = cmbZone.ItemData(cmbZone.ListIndex) Else numZoneID = Null
''''''
''''''                params = ""
''''''                params = params + CStr(IIf(IsNull(mZoneNo), "NA", mZoneNo)) + "~"        ' 1.Zone
''''''                params = params + CStr(IIf(IsNull(intWardNo), "NA", intWardNo)) + "~"    ' 2.WardID
''''''                params = params + CStr(IIf(IsNull(intDoorNo1), "NA", intDoorNo1)) + "~"  ' 3.DoorNo
''''''                params = params + CStr(IIf(IsNull(chvDoorNo2), "NA", chvDoorNo2)) + "~"  ' 4.Door No2
''''''                params = params + CStr(IIf(IsNull(chvName), "NA", chvName)) + "~"        ' 5.Owner Name
''''''                params = params + "NA~"   ' 6.Address
''''''                params = params + "NA~"   ' 7.Pin
''''''                params = params + "NA~"   ' 8.Phone
''''''                params = params + "NA~"   ' 9.Application No
''''''                params = params + "NA~"   ' 10.AssessNo
''''''                params = params + "NA"   ' 11.Application Stuatsu
''''''
''''''                mUrl = gbDefaultUrl + "/searchAssesmentDetailsUTF16?searchParam=" + params
''''''                'xmlHttp.Open "POST", "http://117.239.77.103:9081/RestFulWSTest/RestFulWSTest/SaankhyaIntegrationService/searchAssesmentDetails?searchParam=" & params, False
''''''                xmlHttp.Open "POST", mUrl, False
''''''                xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
''''''                xmlHttp.send
''''''
''''''                mXmlString = xmlHttp.responseText
''''''                'mXmlString = Replace(mXmlString, "UTF-8", "UTF-16")
''''''
''''''                Set oDoc = CreateObject("MSXML2.DOMDocument")
''''''                oDoc.async = False
''''''                oDoc.validateOnParse = False
''''''                If Not oDoc.LoadXml(mXmlString) Then
''''''                    MsgBox "Error Loading"
''''''                    Exit Sub
''''''                Else
''''''                    'MsgBox "Sucess"
''''''                End If
''''''
''''''                Set oRs = New ADODB.Recordset
''''''                Set oRs.ActiveConnection = Nothing
''''''                oRs.CursorLocation = adUseClient
''''''                oRs.LockType = adLockBatchOptimistic
''''''
''''''                With oRs.Fields
''''''                    .Append "buildlingID", adInteger
''''''                    .Append "assementYear", adVarChar, 20
''''''                    .Append "zoneID", adInteger
''''''                    .Append "wardNo", adInteger
''''''                    .Append "wardName", adVarChar, 20
''''''                    .Append "doorNo1", adInteger
''''''                    .Append "doorNo2", adVarChar, 10
''''''                    .Append "ownerName", adVarChar, 50
''''''                    .Append "houseBuildingName", adVarChar, 50
''''''                    .Append "street", adVarChar, 50
''''''                    .Append "localplace", adVarChar, 50
''''''                    .Append "mainplace", adVarChar, 50
''''''                    .Append "post", adVarChar, 10
''''''                    .Append "district", adVarChar, 50
''''''                    .Append "pin", adVarChar, 20
''''''                    .Append "phone", adVarChar, 20
''''''                    .Append "ownerMobileNo", adVarChar, 20
''''''
''''''                End With
''''''
''''''                oRs.Open
''''''                 For Each oNode In oDoc.selectNodes("/PropertyTaxVOss/propertyTaxVO")
''''''                 'For Each oNode In oDoc.selectNodes("/PropertyTaxVo/demandRegisters")
''''''                    oRs.ADDNEW
''''''                    oRs.Fields("buildlingID").value = oNode.selectSingleNode("buildlingID").Text
''''''                    oRs.Fields("assementYear").value = oNode.selectSingleNode("assementYear").Text
''''''                    oRs.Fields("zoneID").value = oNode.selectSingleNode("zoneID").Text
''''''                    oRs.Fields("wardNo").value = oNode.selectSingleNode("wardNo").Text
''''''                    oRs.Fields("wardName").value = oNode.selectSingleNode("wardName").Text
''''''                    oRs.Fields("doorNo1").value = oNode.selectSingleNode("doorNo1").Text
''''''                    oRs.Fields("doorNo2").value = oNode.selectSingleNode("doorNo2").Text
''''''                    oRs.Fields("ownerName").value = oNode.selectSingleNode("ownerName").Text
''''''                    oRs.Fields("houseBuildingName").value = oNode.selectSingleNode("houseBuildingName").Text
''''''                    oRs.Fields("street").value = oNode.selectSingleNode("street").Text
''''''                    oRs.Fields("localplace").value = oNode.selectSingleNode("localplace").Text
''''''                    oRs.Fields("mainplace").value = oNode.selectSingleNode("mainplace").Text
''''''                    oRs.Fields("post").value = oNode.selectSingleNode("post").Text
''''''                    oRs.Fields("district").value = oNode.selectSingleNode("district").Text
''''''                    oRs.Fields("pin").value = oNode.selectSingleNode("pin").Text
''''''                    oRs.Fields("phone").value = oNode.selectSingleNode("phone").Text
''''''                    oRs.Fields("ownerMobileNo").value = oNode.selectSingleNode("ownerMobileNo").Text
''''''                Next
''''''                If oRs.RecordCount > 0 Then
''''''                    mBuildingID = oRs.Fields("buildlingID").value
''''''                    txtBuildingNo.Text = oRs.Fields("buildlingID").value
''''''                    txtAddress.Text = oRs.Fields("ownerName").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("houseBuildingName").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("street").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("localplace").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("mainplace").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("post").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("district").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("pin").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("phone").value
''''''                    txtAddress.Text = txtAddress.Text & vbCrLf & oRs.Fields("ownerMobileNo").value
''''''
''''''                    vchName_3 = IIf(IsNull(oRs.Fields("ownerName").value), "", oRs.Fields("ownerName").value)
''''''                    vchHouseName_4 = IIf(IsNull(oRs.Fields("houseBuildingName").value), "", oRs.Fields("houseBuildingName").value)
''''''                    vchMainPlace_6 = IIf(IsNull(oRs.Fields("mainplace").value), "", oRs.Fields("mainplace").value)
''''''                    vchStreetName_5 = IIf(IsNull(oRs.Fields("street").value), "", oRs.Fields("street").value)
''''''
''''''                    'vchNarration_10 = IIf(IsNull(Rec!chvRemarks), "", Rec!chvRemarks)
''''''                    'vchRef_11 = IIf(IsNull(Rec!chvRefNo), "", Rec!chvRefNo)
''''''
''''''                    Call DisplayBuildingTaxDemands(mBuildingID)
''''''                Else
''''''                    txtBuildingNo.Text = ""
''''''                    txtAddress.Text = ""
''''''                End If
''''''                Me.MousePointer = vbDefault
''''''                Exit Sub
''''
''''
''''            '-------------------------------------------------------'
''''            '----------Added On 21 Jul 2015 Building Search From Web
''''            ' Web Service  '
''''            ' To fetch data from Sanchaya Web
''''            '-------------------------------------------------------'
''''''            ElseIf mDemandWeb = True Then
                
            If mDemandWeb = True Then
                           
           'Added On 21 Jul 2015 By Anisha For Fetching building details
        
            'If mDemandWeb = True Then
                mBuildingWeb = CStr(IIf(IsNull(val(txtBuildingNo)), 0, val(txtBuildingNo)))
                numZoneID = IIf(IsNull(numZoneID), 0, numZoneID)
                intAssessmentYear = IIf(IsNull(intAssessmentYear), 0, intAssessmentYear)
                intWardNo = IIf(IsNull(intWardNo), 0, intWardNo)
                intDoorNo1 = IIf(IsNull(intDoorNo1), 0, intDoorNo1)
                chvDoorNo2 = CStr(IIf(IsNull(chvDoorNo2), 0, chvDoorNo2))
            'End If
            
                mUrl = gbDefaultUrl
                
                Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                objSOAP.MSSoapInit mUrl + "?WSDL"
            On Error GoTo WebConnectionERROR:
                If mBuildingWeb = "0" Then
                    mArrOut = objSOAP.getBuildingDetSaankhyaXML(gbLBID, mBuildingWeb, numZoneID, intAssessmentYear, intWardNo, intDoorNo1, chvDoorNo2)
                Else
                    mArrOut = objSOAP.getBuildingDetSaankhyaXML(gbLBID, mBuildingWeb, numZoneID, 0, 0, 0, "0")
                End If
                
            On Error GoTo ERROR_AfterWEBService:
                mXmlStream.Open

                mXmlStream.WriteText mArrOut
                mXmlStream.Position = 0
                Rec.Open mXmlStream
                mXmlStream.Close


                If Not (Rec.BOF And Rec.EOF) Then
                    
                    txtBuildingNo.Text = Rec!numBuildingID
                    txtWardNo.Tag = Rec!intWardNo
                    txtWardNo.Enabled = False
                    txtHouseNo1.Text = Rec!intDoorNo
                    txtHouseNo1.Tag = Rec!intDoorNo
                    txtHouseNo1.Enabled = False
                    txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNOSub), "", Rec!chvDoorNOSub)
                    txtHouseNo2.Enabled = False
                    cmbWard.Enabled = False
                    
                    vchName_3 = IIf(IsNull(Rec!chvownerEng), "", Rec!chvownerEng)
                    vchHouseName_4 = IIf(IsNull(Rec!chvHouseNameEng), "", Rec!chvHouseNameEng)
                    vchMainPlace_6 = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                    vchStreetName_5 = IIf(IsNull(Rec!chvResStreetName), "", Rec!chvResStreetName)
                    
                    vchNarration_10 = IIf(IsNull(Rec!chvRemarks), "", Rec!chvRemarks)
                    vchRef_11 = IIf(IsNull(Rec!chvRefNo), "", Rec!chvRefNo)
                    
                    txtAddress.Text = vchName_3
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName_4
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
                    txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice_7
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
                    txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
                    
                    txtHalfYearTaxRate.Text = Format(Rec!fltPTax2, "0.00")
                    mBuildingID = val(txtBuildingNo.Text)
                  '  RRStayFlag--Returns : 0 - Not in Revenue Recovery/Prosecution/Stay, 1 - The building is under Revenue Recovery/Prosecution/Stay
                    If (IIf(IsNull(Rec!RRStayFlag), 0, Rec!RRStayFlag)) Then
                        lblRRFlag.Tag = Format(Rec!RRStayFlag, "0")
                    Else
                        lblRRFlag.Tag = Format(Rec!RRStayFlag, "0")
                        lblRRFlag.Caption = "The building is under Revenue Recovery/Prosecution/Stay"
                    End If
                    Call DisplayBuildingTaxDemands(mBuildingID)

                Else
                    MsgBox "No Record Found"
                End If
               
            'Note:- MAIN OFFICE
            ElseIf Right(gbLocationID, 2) = 1 Then
                If mvarDifferentZoneFlag Then
                    If objdb.CreateNewConnection(mCnn, SanchayaHO) = False Then
                        mSql = "Didn't able to connect to the Main office Server"
                        MsgBox mSql, vbInformation
                        Exit Sub
                    End If
                Else
                    If objdb.CreateNewConnection(mCnn, SanchayaLite) = False Then
                        mSql = "Didn't able to connect to the Sanchaya Server"
                        MsgBox mSql, vbInformation
                        Exit Sub
                    End If
                End If
                Set Rec = objdb.ExecuteSP("spSanGetSearchBuildingList", arrInput, , , mCnn, adCmdStoredProc)
                If Not (Rec.BOF And Rec.EOF) Then
                    mBuildingID = Rec!numBuildingID
                    txtBuildingNo.Text = Rec!numBuildingID
                    'txtWard.Text = Rec!intWardNO
                    txtHouseNo1.Text = Rec!intDoorNo1
                    txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
                    
                    vchName_3 = IIf(IsNull(Rec!chvOwners), "", Rec!chvOwners)
                    'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial1), "", "." & Rec!chvInitial1)
                    'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial2), "", "." & Rec!chvInitial2)
                    'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial3), "", "." & Rec!chvInitial3)
                    'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial4), "", "." & Rec!chvInitial4)
                    vchHouseName_4 = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
                    vchMainPlace_6 = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                    vchStreetName_5 = IIf(IsNull(Rec!chvResStreetName), "", Rec!chvResStreetName)
                    'vchMainPlace_6 = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                    'vchPostOffice_7 = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
                    'vchDistrict_8 = IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
                    'vchPinNumber_9 = IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
                    
                    vchNarration_10 = IIf(IsNull(Rec!chvRemarks), "", Rec!chvRemarks)
                    vchRef_11 = IIf(IsNull(Rec!chvRefNo), "", Rec!chvRefNo)
                    
                    txtAddress.Text = vchName_3
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName_4
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
                    txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice_7
                    txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
                    txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
                    
                    txtHalfYearTaxRate.Text = Format(Rec!fltPTax2, "0.00")
                    
                    Call DisplayBuildingTaxDemands(mBuildingID)
                    Call Calculate
                Else
                    MsgBox "This Door No Does Not Exists", vbInformation
                    txtHouseNo1.Text = ""
                    txtHouseNo2.Text = ""
                    txtHouseNo1.SetFocus
                End If
                Rec.Close
                Set mCnn = Nothing
            Else 'Note:- Location Other than Main Office
                If mvarDifferentZoneFlag Then
                    '--------------'
                    ' Web Service  '
                    '--------------'
                    '----------------------------------------------------------------------
                    'Coded on 16.08.09 to get Demand From SanchayaHO for Zonal Integration
                    'Added "Microsoft Soap Type Library" Reference
                    '----------------------------------------------------------------------
                    mUrl = gbDefaultUrl 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultUrl")
                    'client.mssoapinit mUrl + "?WSDL"
                    'mArrOut = (client.GetDemandList(mArrIn))
                    
                    '----------------------------------------------------------------------'
                    'Changed By Aiby : To Support in WINDOWS 2000 Server
                    '----------------------------------------------------------------------'
                    Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                    objSOAP.MSSoapInit mUrl + "?WSDL"
                    mArrOut = (objSOAP.GetDemandList(mArrIn))
                    '----------------------------------------------------------------------'
                    mGetString = Split(mArrOut, "~")
                    If Not IsMissing(mGetString) Then
                         mBuildingID = val(mGetString(7))
                        txtBuildingNo.Text = mGetString(7)
                        txtWardNo.Text = val(mGetString(13))
                        txtHouseNo1.Text = mGetString(11)
                        txtHouseNo2.Text = mGetString(12)
                        vchName_3 = mGetString(0)
                        vchHouseName_4 = mGetString(1)
                        txtAddress.Text = vchName_3
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName_4
    '                    txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
    '                    txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
    '                    txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice_7
    '                    txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
    '                    txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
                        
                        txtHalfYearTaxRate.Text = Format(mGetString(15), "0.00")
                        
                        Call DisplayBuildingTaxDemands(mBuildingID)
                        Call Calculate
                    Else
                          MsgBox "This Door No Does Not Exists", vbInformation
                        txtHouseNo1.Text = ""
                        txtHouseNo2.Text = ""
                        txtHouseNo1.SetFocus
                    End If
                Else
                    '----------'
                    ' Common Op'
                    '----------'
                    If objdb.CreateNewConnection(mCnn, SanchayaLite) = False Then
                        mSql = "Didn't able to connect to the Sanchaya Server"
                        MsgBox mSql, vbInformation
                        Exit Sub
                    End If
                    
                    
                    Set Rec = objdb.ExecuteSP("spSanGetSearchBuildingList", arrInput, , , mCnn, adCmdStoredProc)
                    If Not (Rec.BOF And Rec.EOF) Then
                        mBuildingID = Rec!numBuildingID
                        txtBuildingNo.Text = Rec!numBuildingID
                        'txtWard.Text = Rec!intWardNO
                        txtHouseNo1.Text = Rec!intDoorNo1
                        txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
                        
                        vchName_3 = IIf(IsNull(Rec!chvOwners), "", Rec!chvOwners)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial1), "", "." & Rec!chvInitial1)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial2), "", "." & Rec!chvInitial2)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial3), "", "." & Rec!chvInitial3)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial4), "", "." & Rec!chvInitial4)
                        vchHouseName_4 = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
                        vchStreetName_5 = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
                        'vchStreetName_5 = IIf(IsNull(Rec!chvResStreetName), "", Rec!chvResStreetName)
                        'vchMainPlace_6 = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                        'vchPostOffice_7 = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
                        'vchDistrict_8 = IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
                        'vchPinNumber_9 = IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
                        
                        txtAddress.Text = vchName_3
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName_4
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
                        txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice_7
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
                        txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
                        
                        txtHalfYearTaxRate.Text = Format(Rec!fltPTax2, "0.00")
                        
                        
                        Call DisplayBuildingTaxDemands(mBuildingID)
                        Call Calculate
                    Else
                        MsgBox "This Door No Does Not Exists", vbInformation
                        txtHouseNo1.Text = ""
                        txtHouseNo2.Text = ""
                        txtHouseNo1.SetFocus
                    End If
                    Rec.Close
                    Set mCnn = Nothing
                End If
            End If
            mSelectedAllFlag = True 'Setting as Complete Demand Selected
            If mAdvanceExists Then
                cmdCopy.Enabled = False
                Dim mStr As String
                mStr = "There is unadjusted demand against this building " & vbCrLf
                mStr = mStr + " which shall be adjusted in Sanchaya"
                MsgBox mStr, vbInformation
            End If
            Me.MousePointer = vbDefault
            Exit Sub
WebConnectionERROR:
        MsgBox "Connection to Web Service Failed :: " & Error, vbInformation
        Exit Sub
ERROR_AfterWEBService:
        MsgBox Error
        
        Me.MousePointer = vbDefault
    End Sub
    Private Sub cmdCopy_Click()
        
        Dim mLoop As Long
        Dim mLoopChild As Long
        Dim mCount As Long
        Dim mFineWaveDate As String
        Dim objAcc As New clsAccounts
        
        If lblRRFlag.Tag = 1 Then
            If (MsgBox("This building is under Revenue Recovery/Prosecution/Stay,Are you sure to do Receipt", vbYesNo)) = vbNo Then
                Exit Sub
            End If
        End If
        
        '--------------------------------------------------------------------------'
        ' To Decide Whether to Fine Waive Or NOT - Based on Circular on Jan, 2010  '
        '
         mFineWaveDate = Format("31/Mar/2010", "dd/mmm/yy")
         For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 12) = 2 Then
                mSelectedAllFlag = False
                Exit For
            End If
         Next
         If mLoop = vsGrid.Rows Then
            mSelectedAllFlag = True
         End If
        '--------------------------------------------------------------------------'
'        If txtWardNo.Text <> txtWardNo.Tag Then
'            MsgBox "Ward No. and Building No. are not Maching", vbApplicationModal
'            Exit Sub
'        End If
'        If txtHouseNo1.Text <> txtHouseNo1.Tag Then
'            MsgBox "Door No. and Building No. are not Maching", vbApplicationModal
'            Exit Sub
'        End If
        '''Modified On 31/Dec/2016

        If mBuildingID > 0 Then
            frmReceiptsCounter.SubLedgerID = mBuildingID
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
            Call frmReceiptsCounter.DisplayBuildingDetails
            frmReceiptsCounter.DemandBasedFlag = True
            ' modified on 12-aug-2009 by cijith for Sanchya Zonal Connectivity  '
            frmReceiptsCounter.AssessmentYear = cmbAssessmentYear.ItemData(cmbAssessmentYear.ListIndex)
            '--------------------------------------------------------------------'
        End If
        
        'Note:- If the demand is from Different Location then change
        '       the default location to the selected location
        If cmbZone.ItemData(cmbZone.ListIndex) <> gbLocationID Then
            frmReceiptsCounter.lblZone.Visible = True
            frmReceiptsCounter.txtZone.Visible = True
            frmReceiptsCounter.txtZone.Tag = cmbZone.ItemData(cmbZone.ListIndex)
            frmReceiptsCounter.txtZone.Text = cmbZone.Text
            'frmReceiptsCounter.cmbDZoned.Text = cmbZone.Text
        End If
        
        
        frmReceiptsCounter.vsGrid.Rows = 1
        frmReceiptsCounter.vsGrid.MergeCells = flexMergeFree
        mCount = 0
        For mLoop = 1 To vsGrid.Rows - 1
            If val(vsGrid.TextMatrix(mLoop, 4)) > 0 Or val(vsGrid.TextMatrix(mLoop, 5)) > 0 Then
                frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 1
                If vsGrid.Cell(flexcpChecked, mLoop, 12) = 1 And val(vsGrid.Cell(flexcpText, mLoop, 11)) > 0 Then
CopyToGrid:
                   
                   If vsGrid.Cell(flexcpText, mLoop, 0) = gbAcHeadCodePoorHomeCess Then
                        frmReceiptsCounter.PoorHomeCess = True
                   End If
                   
                   '------For Fine wave
                    If vsGrid.TextMatrix(mLoop, 0) = 140200200 Or vsGrid.TextMatrix(mLoop, 0) = 140200101 Then
                        If mSelectedAllFlag And gbTransactionDate <= mFineWaveDate Then
                            Call FineWave
                            vsGrid.TextMatrix(mLoop, 5) = 0
                            frmReceiptsCounter.mFinewave = True
                        End If
                    End If
                    mCount = mCount + 1
                    frmReceiptsCounter.vsGrid.Row = mCount
                    For mLoopChild = 0 To vsGrid.Cols - 1
                        If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
                            If mLoopChild = 2 Then
                                frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, mLoopChild) = vsGrid.Cell(flexcpText, mLoop, mLoopChild + 5)
                                frmReceiptsCounter.vsGrid.Cell(flexcpChecked, mCount, 12) = 1
                            ElseIf mLoopChild = 3 Then
                                frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, mLoopChild) = vsGrid.Cell(flexcpText, mLoop, mLoopChild + 5)
                                frmReceiptsCounter.vsGrid.Cell(flexcpChecked, mCount, 12) = 1
                            Else
                                frmReceiptsCounter.vsGrid.Cell(flexcpText, mCount, mLoopChild) = vsGrid.Cell(flexcpText, mLoop, mLoopChild)
                                frmReceiptsCounter.vsGrid.Cell(flexcpChecked, mCount, 12) = 1
                            End If
                            
                        End If
                    Next
                Else ' Check for case: Fine
                     
                End If
            End If
        Next
            ''-----------Adding New Row to vsGrid for Advance------------''
            If val(txtGrandTotal.Text) > val(txtNetAmount.Text) Then
                With frmReceiptsCounter.vsGrid
                    .Rows = .Rows + 1         ' One Row Added
                    objAcc.SetAccountCode ("350410101")             '' Setting Accounts with Head Code
                    .Cell(flexcpText, mCount + 1, 0) = "350410101"             'AccountHead
                    .Cell(flexcpText, mCount + 1, 1) = objAcc.AccountHead      'AccountHead
                    .Cell(flexcpText, mCount + 1, 2) = gbFinancialYearID       'YearID
                    .Cell(flexcpText, mCount + 1, 3) = gbCurrentPeriodID       'Period ID
                    .Cell(flexcpText, mCount + 1, 5) = val(txtGrandTotal.Text) - val(txtNetAmount.Text) 'Current Amount
                    .Cell(flexcpText, mCount + 1, 6) = objAcc.AccountHeadID
                    .Cell(flexcpText, mCount + 1, 7) = gbFinancialYearID
                    .Cell(flexcpText, mCount + 1, 8) = 1
                    .Cell(flexcpText, mCount + 1, 10) = ""                     'Demand ID What??
                    .Cell(flexcpText, mCount + 1, 11) = val(txtGrandTotal.Text) - val(txtNetAmount.Text) 'Current Amount                       'Amount Paid
                    .Cell(flexcpChecked, mCount + 1, 12) = 1                  'Checked
                    frmReceiptsCounter.txtBuildingNo.Text = val(txtBuildingNo.Text)
                End With
            End If
            ''-----------------------------------------------------------''
             ''-----------Adding New Row to vsGrid for NoticeFee------------''
            If val(txtNoticeFee.Text) > 0 Then
   
                With frmReceiptsCounter.vsGrid
                    .Rows = .Rows + 1         ' One Row Added
                    objAcc.SetAccountCode (gbAcHeadCodeNoticeFee)             '' Setting Accounts with Head Code
                    .Cell(flexcpText, mCount + 1, 0) = gbAcHeadCodeNoticeFee             'AccountHead
                    .Cell(flexcpText, mCount + 1, 1) = objAcc.AccountHead      'AccountHead
                    .Cell(flexcpText, mCount + 1, 2) = gbFinancialYearID       'YearID
                    .Cell(flexcpText, mCount + 1, 3) = gbCurrentPeriodID       'Period ID
                    .Cell(flexcpText, mCount + 1, 5) = val(txtNoticeFee.Text) 'Current Amount
                    .Cell(flexcpText, mCount + 1, 6) = objAcc.AccountHeadID
                    .Cell(flexcpText, mCount + 1, 7) = gbFinancialYearID
                    .Cell(flexcpText, mCount + 1, 8) = 1
                    .Cell(flexcpText, mCount + 1, 10) = ""                     'Demand ID What??
                    .Cell(flexcpText, mCount + 1, 11) = val(txtNoticeFee.Text) 'Current Amount                       'Amount Paid
                    .Cell(flexcpChecked, mCount + 1, 12) = 1                  'Checked
                End With
            End If
            frmReceiptsCounter.vsGrid.Editable = flexEDNone
            frmReceiptsCounter.txtAdvance = txtAdvance.Text
            frmReceiptsCounter.Calculate
            
            frmReceiptsCounter.txtWardNo.Enabled = False
            frmReceiptsCounter.txtDoorNo1.Enabled = False
            frmReceiptsCounter.txtDoorNo2.Enabled = False
            frmReceiptsCounter.txtName.Enabled = False
            frmReceiptsCounter.txtInit1.Enabled = False
            frmReceiptsCounter.txtInit2.Enabled = False
            frmReceiptsCounter.txtInit3.Enabled = False
            frmReceiptsCounter.txtInit4.Enabled = False
            frmReceiptsCounter.txtHouse.Enabled = False
            frmReceiptsCounter.txtStreet.Enabled = False
            frmReceiptsCounter.txtLocalPlace.Enabled = False
            frmReceiptsCounter.txtMainPlace.Enabled = False
            frmReceiptsCounter.txtPost.Enabled = False
            frmReceiptsCounter.txtPin.Enabled = False
            frmReceiptsCounter.txtPhone.Enabled = False
            
            Unload Me
        
    End Sub

    Private Sub cmdListDemand_Click()
        Call FillPTax
        Call Calculate
        
    End Sub

    Private Sub cmdsearch_Click()
        Unload Me
        frmSearchBuildingDetails.Visible = True
        frmSearchBuildingDetails.ZOrder (0)
    End Sub
    Private Sub Form_Activate()
        Me.Left = (frmMenu.Width - Me.Width) / 2
        Me.Top = 290
        
        If mvarBuildingID > -1 Then
            Dim arrInput As Variant
            Dim Rec As New ADODB.Recordset
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim objPTax As New clsPTax
            Dim mLoop As Integer
               
                '-------------------------------------------------------------'
                ' Changed to Change the Stored Procedure spGetBuildingDetails '
                ' New Stored Procedure is spSanGetSearchBuildingList          '
                '-------------------------------------------------------------'
                'arrInput = Array(Null, _
                'Null, _
                'Null, _
                'Null, _
                'mvarBuildingID)
                'mvarBuildingID = -1
                '--------------------------------------'
                ' Additions                            '
                '--------------------------------------'
                Dim numBuildingID   As Variant
                Dim numZoneID       As Variant
                Dim intAssessmentYear As Variant
                Dim intWardNo       As Variant
                Dim intDoorNo1      As Variant
                Dim chvDoorNo2      As Variant
                Dim chvName         As Variant
                Dim chvResHName     As Variant
                
                
                If Trim(txtBuildingNo) <> "" Then numBuildingID = val(txtBuildingNo) Else numBuildingID = Null
                If cmbZone.ListIndex > -1 Then numZoneID = cmbZone.ItemData(cmbZone.ListIndex) Else numZoneID = Null
                If cmbAssessmentYear.ListIndex > -1 Then intAssessmentYear = cmbAssessmentYear.ItemData(cmbAssessmentYear.ListIndex) Else intAssessmentYear = Null
                If cmbWard.ListIndex > -1 Then intWardNo = cmbWard.ItemData(cmbWard.ListIndex) Else intWardNo = Null
                If Trim(txtHouseNo1) <> "" Then intDoorNo1 = val(txtHouseNo1) Else intDoorNo1 = Null
                If Trim(txtHouseNo2) <> "" Then chvDoorNo2 = Trim(txtHouseNo2) Else chvDoorNo2 = Null
                chvName = Null
                chvResHName = Null
                
                arrInput = Array(numBuildingID, _
                    numZoneID, _
                    intAssessmentYear, _
                    intWardNo, _
                    intDoorNo1, _
                    chvDoorNo2, _
                    chvName, _
                    chvResHName)
                mvarBuildingID = -1
                If objdb.CreateNewConnection(mCnn, SanchayaLite) Then
                    Set Rec = objdb.ExecuteSP("spSanGetSearchBuildingList", arrInput, , , mCnn, adCmdStoredProc)
                    If Not (Rec.BOF And Rec.EOF) Then
                        mBuildingID = Rec!numBuildingID
                        txtBuildingNo.Text = Rec!numBuildingID
                        For mLoop = 0 To cmbWard.ListCount - 1
                            If cmbWard.ItemData(mLoop) = Rec!intWardNo Then
                                cmbWard.ListIndex = mLoop
                                Exit For
                            End If
                        Next
                        txtHouseNo1.Text = Rec!intDoorNo1
                        txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
                        vchName_3 = IIf(IsNull(Rec!chvOwners), "", Rec!chvOwners)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial1), "", "." & Rec!chvInitial1)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial2), "", "." & Rec!chvInitial2)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial3), "", "." & Rec!chvInitial3)
                        'vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial4), "", "." & Rec!chvInitial4)
                        vchHouseName_4 = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
                        vchStreetName_5 = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
                        vchMainPlace_6 = "" ' IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                        vchPostOffice_7 = "" 'IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
                        vchDistrict_8 = "" 'IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
                        vchPinNumber_9 = "" 'IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
                        
                        txtAddress.Text = vchName_3
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName_4
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
                        txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice_7
                        txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
                        'txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
                        Call cmdFind_Click
                    End If
                    Rec.Close
                End If
                Set objdb = Nothing
                'txtWardNo.SetFocus
        End If
    End Sub
    Private Sub FineWave()
        'Finewaving for Receipt in (as per Circular on Jan, 2010)
        Dim objdb           As New clsDB
        Dim Rec             As New Recordset
        Dim mCnn            As New ADODB.Connection
        Dim mArrIn          As Variant
        Dim mArrOut         As Variant
        Dim mVoucherNo        As Variant
        Dim mRemarks        As String
        Dim mSql            As String
        Dim mFineWaveID     As Variant
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
             mArrIn = Array(gbCounterID, _
                           val(frmReceiptsCounter.txtInstrument.Tag), _
                           gbFinancialYearID)
             objdb.ExecuteSP "spGetNextReceiptNo", mArrIn, mArrOut, , mCnn, adCmdStoredProc
            If IsArray(mArrOut) Then
                mVoucherNo = mArrOut(0, 0)
                mSql = "Select * from faFineWaiver Where intVoucherNo=" & mVoucherNo
                Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
                If Rec.RecordCount > 0 Then
                    mFineWaveID = Rec!intID
                Else
                    mFineWaveID = Null
                End If
            End If
            mRemarks = ""
            mArrIn = Array(gbTransactionDate, _
                            gbTransactionTypePTax, _
                            gbUserID, _
                            gbSeatID, _
                            gbCounterID, _
                            mVoucherNo, _
                            val(txtFine.Text), _
                            0, _
                            mRemarks, _
                            mFineWaveID, _
                            gbLocalBodyID, _
                            mBuildingID, _
                            gbLocationID)
            
             objdb.ExecuteSP "spSaveFineWaiver", mArrIn, , , mCnn, adCmdStoredProc
        Else
            MsgBox "Connection To Saankhya Doesn't Exists"
       
        End If
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            Unload Me
            frmSearchBuildingDetails.Visible = True
            frmSearchBuildingDetails.ZOrder (0)
        End If
    End Sub

    Private Sub Form_Load()
        WindowsXPC.InitSubClassing
        mvarBuildingID = -1
        If gbLinkWithPropertyTax Then
            cmdListDemand.Enabled = False
        End If
        Call FillYear
        Call FillAssessmentYear
        Call FormInitialize
        Call FillZone
        Call SetDefaultSettings
        Call FillWard
        lblGrandTotal.Caption = ""
    End Sub
    Private Sub FindNumberOfHalfYears()
        
        Dim mCount As Long
        Dim mNoOfyears As Long
        On Error Resume Next
        mNoOfyears = cmbToYear.ItemData(cmbToYear.ListIndex) - cmbFromYear.ItemData(cmbFromYear.ListIndex)
        mCount = mNoOfyears * 2 + 2
        If cmbFromPeriod.ItemData(cmbFromPeriod.ListIndex) = 2 Then
            mCount = mCount - 1
        End If
        If cmbToPeriod.ItemData(cmbToPeriod.ListIndex) = 1 Then
            mCount = mCount - 1
        End If
        txtNoOfHalfYears.Text = mCount
        
    End Sub



'''    Private Sub lblFine_Click()
'''        Call DisplayGrid
'''    End Sub
    Private Sub txtBuildingNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            PressTabKey
        End If
    End Sub





    Private Sub txtGrandTotal_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, ".")
    End Sub

    Private Sub txtGrandTotal_LostFocus()
'       Call DisplayGrid
        Call AutoCheckDemand
        cmdCopy.Enabled = True
        If val(txtNetAmount.Text) < 1 Then
            cmdCopy.Enabled = False
        End If
'        If val(txtGrandTotal.Text) > val(lblGrandTotal.Caption) And val(txtGrandTotal.Text) > 0 Then
'            If val(txtNetAmount.Text) > -1 Then
'                cmdCopy.Enabled = True
'            End If
'        Else
'            If val(txtGrandTotal.Text) > 0 Then
'                cmdCopy.Enabled = False
'            End If
'        End If

    
    End Sub
    Private Sub txtHalfYearTaxRate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    Private Sub txtHalfYearTaxRate_LostFocus()
        txtHalfYearTaxRate.Text = Format(val(txtHalfYearTaxRate), "0.00")
    End Sub
    Private Sub txtHouseNo1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            PressTabKey
        End If
    End Sub
    Private Sub txtHouseNo2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            PressTabKey
        End If
    End Sub
    Private Sub txtWard_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub

  




    Private Sub txtNoticeFee_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, ".")
        
    End Sub

    Private Sub txtNoticeFee_LostFocus()
        Call Calculate
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
    Private Sub vsGrid_Click()
        If vsGrid.Col = 12 Then
            If vsGrid.TextMatrix(vsGrid.Row, 0) <> mAcHeadCodeFine Then
                Call Calculate
            End If
        End If
    End Sub
    Private Sub vsGrid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        Dim mLoop As Long
        If Row > 0 Then
            If vsGrid.Cell(flexcpChecked, Row, Col) = 2 Then
                If Row = 1 Or vsGrid.Cell(flexcpChecked, Row - 1, Col) = vbChecked Then
                    vsGrid.Cell(flexcpChecked, Row, Col) = vbChecked
                    mNumberOfSelections = mNumberOfSelections + 1 'IIf(Row Mod 2 = 0, 1, 0)
                Else
                    Cancel = True
                End If
            Else ' Already  Checked
                If vsGrid.Cell(flexcpChecked, Row - 1, Col) = 1 Then
                    For mLoop = 1 To vsGrid.Rows - 1
                        If vsGrid.TextMatrix(Row, 10) <> vsGrid.TextMatrix(mLoop, 10) Then
                            vsGrid.Cell(flexcpChecked, mLoop, 12) = 2
                            'If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
                            'Cancel = True
                            'End If
                            mNumberOfSelections = mNumberOfSelections - 1
                            'Exit For
                        End If
                    Next mLoop
                    mFineWaiveFlag = False
                Else
                    Cancel = True
                End If
            End If
        End If
    End Sub
    
    Private Sub AutoCheckDemand()       ' Sinoj
        If txtGrandTotal.Text <> "" Then
            Dim mAdv    As Double
            Dim mFine   As Double
            Dim mCnt    As Integer
            
            Dim mYearID As Integer
            Dim mPeriod As Integer
            
            Dim mLoop As Integer
            
            If vsGrid.Rows > 1 Then
                vsGrid.Cell(flexcpChecked, 1, 12, vsGrid.Rows - 1, 12) = 2
            End If
            
            
            lblTotalArrear.Caption = ""
            lblTotalCurrent.Caption = ""
            lblGrandTotal.Caption = ""
            txtFine.Text = ""
            txtAdvance.Text = ""
            txtNetAmount.Text = ""

            For mCnt = 1 To vsGrid.Rows - 1
                
                '--------------Checking the Demand IDs are Equal-------------------'
                
                
                'mintYearID_6 = val(vsGrid.Cell(flexcpText, mLoopCount, 7))
                'mtnyPeriodID_7 = val(vsGrid.Cell(flexcpText, mLoopCount, 8))
                If val(vsGrid.Cell(flexcpText, mCnt, 7)) <> mYearID Then
                    mYearID = val(vsGrid.Cell(flexcpText, mCnt, 7))
                End If
                If val(vsGrid.Cell(flexcpText, mCnt, 8)) <> mPeriod Then
                    mPeriod = val(vsGrid.Cell(flexcpText, mCnt, 8))
                End If
                
                For mLoop = mCnt To vsGrid.Rows - 1
                    If val(vsGrid.Cell(flexcpText, mCnt, 7)) = mYearID And _
                        val(vsGrid.Cell(flexcpText, mCnt, 8)) = mPeriod Then
                        
                        vsGrid.Cell(flexcpChecked, mCnt, 12) = 1
                        mCnt = mCnt + 1
                    Else
                        'mYearID = val(vsGrid.Cell(flexcpText, mCnt, 7))
                        'mPeriod = val(vsGrid.Cell(flexcpText, mCnt, 8))
                        mCnt = mCnt - 1
                        Exit For
                    End If
                Next mLoop
                Call Calculate
                
''                If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
''                    For mLoop = mCnt To 1 Step -1
''                    If mLoop <= vsGrid.Rows - 1 Then
''                        If val(vsGrid.Cell(flexcpText, mCnt, 7)) = mYearID And _
''                            val(vsGrid.Cell(flexcpText, mCnt, 8)) = mPeriod Then
''
''                            vsGrid.Cell(flexcpChecked, mCnt, 12) = 2
''                            mCnt = mCnt - 1
''                        Else
''                            mYearID = val(vsGrid.Cell(flexcpText, mCnt, 7))
''                            mPeriod = val(vsGrid.Cell(flexcpText, mCnt, 8))
''                            Exit For
''                        End If
''                    End If
''                    Next
''                    Call Calculate
''                    Exit For
''                Else
''                    'Exit For
''                End If
                ''''
                 If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
                    For mLoop = mCnt To 1 Step -1
                    If mLoop <= vsGrid.Rows - 1 Then
                        If val(vsGrid.Cell(flexcpText, mLoop, 7)) = mYearID And _
                            val(vsGrid.Cell(flexcpText, mLoop, 8)) = mPeriod Then
                            
                            vsGrid.Cell(flexcpChecked, mLoop, 12) = 2
                            mCnt = mCnt - 1
                        Else
                            mYearID = val(vsGrid.Cell(flexcpText, mLoop, 7))
                            mPeriod = val(vsGrid.Cell(flexcpText, mLoop, 8))
                            Exit For
                        End If
                    End If
                    Next
                    Call Calculate
                    Exit For
                Else
                    'Exit For
                End If
                                                                                                    '--1
            Next
            
        End If
    End Sub
    
    
    Private Sub AutoCheckDemand1()       ' Sinoj
        If txtGrandTotal.Text <> "" Then
            Dim mAdv    As Double
            Dim mFine   As Double
            Dim mCnt    As Integer
            
            If vsGrid.Rows > 1 Then
                vsGrid.Cell(flexcpChecked, 1, 12, vsGrid.Rows - 1, 12) = 2
            End If
            
            lblTotalArrear.Caption = ""
            lblTotalCurrent.Caption = ""
            lblGrandTotal.Caption = ""
            txtFine.Text = ""
            txtAdvance.Text = ""
            txtNetAmount.Text = ""
'
            For mCnt = 1 To vsGrid.Rows - 1
                '--------------Checking the Demand IDs are Equal-------------------'
                If mCnt + 1 < vsGrid.Rows Then                                                                  '--1
                    '--------------Checking the Next Row is Equal
                    
                    'If vsGrid.TextMatrix(mCnt + 1, 0) = gbAcHeadCodeLibraryCess Then
                    If vsGrid.TextMatrix(mCnt, 10) = vsGrid.TextMatrix(mCnt + 1, 10) Then                       '--2
                        If mCnt + 2 < vsGrid.Rows Then                                                          '--3
                            
                            If vsGrid.TextMatrix(mCnt, 10) = vsGrid.TextMatrix(mCnt + 2, 10) Then               '--4
                                vsGrid.Cell(flexcpChecked, mCnt, 12) = 1            ' Checking the Row
                                vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = 1        ' Checking the Row
                                'vsGrid.Cell(flexcpChecked, mCnt + 2, 12) = 1        ' Checking the Row
                                mCnt = mCnt + 2                                     ' Incrementing Row
                                If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then                           '--5.1
                                    Call Calculate                                      ' Calculating the Amount
                                End If                                                                          '
                                If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then 'Amount Checking      '--5.2
                                    vsGrid.Cell(flexcpChecked, mCnt, 12) = 2            ' UnChecking the Row
                                    vsGrid.Cell(flexcpChecked, mCnt - 1, 12) = 2        ' UnChecking the Row
                                    vsGrid.Cell(flexcpChecked, mCnt - 2, 12) = 2        ' UnChecking the Row
                                    Call Calculate                                      ' Again Calculating
                                    Exit For
                                End If
                            Else                                                                                '--4
                                vsGrid.Cell(flexcpChecked, mCnt, 12) = 1            ' Checking the Row
                                vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = 1        ' Checking the Row
                                mCnt = mCnt + 1                                     ' Incrementing Row
                                If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then                           '--4.1
                                    Call Calculate                                      ' Calculating the Amount
                                End If
                                If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then 'Amount Checking      '--4.2
                                    vsGrid.Cell(flexcpChecked, mCnt, 12) = 2            ' UnChecking the Row
                                    vsGrid.Cell(flexcpChecked, mCnt - 1, 12) = 2        ' UnChecking the Row
                                    Call Calculate                                      ' Again Calculating
                                    Exit For
                                End If
                            End If                                                                              '--4
                        Else                                                                                    '--3
                            vsGrid.Cell(flexcpChecked, mCnt, 12) = 1            ' Checking the Row
                            vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = 1        ' Checking the Row
                            mCnt = mCnt + 1                                     ' Incrementing Row
                            If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then                               '--3.1
                                Call Calculate                                      ' Calculating the Amount
                            End If
                            If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then 'Amount Checking          '--3.2
                                vsGrid.Cell(flexcpChecked, mCnt, 12) = 2            ' UnChecking the Row
                                vsGrid.Cell(flexcpChecked, mCnt - 1, 12) = 2        ' UnChecking the Row
                                Call Calculate                                      ' Again Calculating
                                Exit For
                            End If
                        End If
                    Else                                                                                        '--2
                        vsGrid.Cell(flexcpChecked, mCnt, 12) = 1            ' Checking the Row
                        If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then                                   '--2.1
                            'Call Calculate                                      ' Calculating the Amount
                        End If
                        If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then 'Amount Checking              '--2.2
                            vsGrid.Cell(flexcpChecked, mCnt, 12) = 2            ' UnChecking the Row
                            'Call Calculate                                      ' Again Calculating
                            Exit For
                        End If
                    End If
                Else                                                                                            '--1
                    vsGrid.Cell(flexcpChecked, mCnt, 12) = 1            ' Checking the Row
                    If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then                                       '--1.1
                        'Call Calculate                                      ' Calculating the Amount
                    End If
                    If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then 'Amount Checking                  '--1.2
                        vsGrid.Cell(flexcpChecked, mCnt, 12) = 2            ' UnChecking the Row
                        'Call Calculate                                      ' Again Calculating
                        Exit For
                    End If
                End If                                                                                      '--1
            Next
            Call Calculate
        End If
    End Sub
    Public Property Let BuildingID(mBuildingID As Double)
        mvarBuildingID = mBuildingID
    End Property
    Property Let DifferentZoneFlag(mFlag As Boolean)
        mvarDifferentZoneFlag = mFlag
    End Property
   
