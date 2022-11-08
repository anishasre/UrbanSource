VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmRentOnLandBuildings 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Rent on Land & Building Search"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeTotal 
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   45
      TabIndex        =   25
      Top             =   4950
      Width           =   11805
      Begin VB.CheckBox chkGrantTotalAdj 
         BackColor       =   &H80000016&
         Caption         =   "Grand Total Adjust"
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
         Left            =   6210
         TabIndex        =   40
         Top             =   360
         Width           =   1950
      End
      Begin VB.CommandButton cmdCopyToReceipt 
         Caption         =   "Copy to Receipt"
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
         Height          =   435
         Left            =   3780
         TabIndex        =   39
         Top             =   810
         Width           =   1545
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5400
         TabIndex        =   38
         Top             =   810
         Width           =   1545
      End
      Begin VB.CheckBox chkFineWaiver 
         BackColor       =   &H80000016&
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
         Height          =   285
         Left            =   6210
         TabIndex        =   36
         Top             =   75
         Width           =   1365
      End
      Begin VB.TextBox txtFine 
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
         Left            =   4875
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   90
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   315
         Width           =   1260
      End
      Begin VB.TextBox txtAdvance 
         Alignment       =   1  'Right Justify
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   615
         Width           =   1260
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label lblGrandTotal 
         Height          =   240
         Left            =   11040
         TabIndex        =   41
         Top             =   360
         Width           =   780
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
         Left            =   4500
         TabIndex        =   37
         Top             =   120
         Width           =   345
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
         Left            =   8415
         TabIndex        =   34
         Top             =   45
         Width           =   1260
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
         Left            =   9720
         TabIndex        =   33
         Top             =   45
         Width           =   1260
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
         Left            =   8550
         TabIndex        =   32
         Top             =   375
         Width           =   1140
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
         Left            =   7650
         TabIndex        =   31
         Top             =   615
         Width           =   2040
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   7920
         TabIndex        =   30
         Top             =   90
         Width           =   375
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
         Left            =   8550
         TabIndex        =   29
         Top             =   945
         Width           =   1140
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsMaster 
      Height          =   2745
      Left            =   1440
      TabIndex        =   8
      Top             =   1755
      Visible         =   0   'False
      Width           =   3690
      _cx             =   6509
      _cy             =   4842
      Appearance      =   1
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
      BackColor       =   -2147483626
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483633
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
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRentOnLandBuildings.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   10
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
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3570
      Top             =   5250
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame framParty 
      BackColor       =   &H80000016&
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
      Height          =   1800
      Left            =   30
      TabIndex        =   20
      Top             =   45
      Width           =   11760
      Begin VB.ComboBox cmbCategory 
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
         Height          =   330
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   585
         Width           =   3210
      End
      Begin VB.CommandButton cmdMaster 
         Caption         =   "...."
         Height          =   345
         Left            =   4725
         TabIndex        =   7
         Top             =   960
         Width           =   345
      End
      Begin VB.TextBox txtMaster 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1455
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   945
         Width           =   3180
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4665
         TabIndex        =   11
         Top             =   1350
         Width           =   405
      End
      Begin VB.ComboBox cmbZone 
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
         Height          =   330
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   3210
      End
      Begin VB.ComboBox cmbWard 
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
         Height          =   330
         Left            =   6495
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2595
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
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   3
         Top             =   255
         Width           =   600
      End
      Begin VB.TextBox txtLesseName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6465
         TabIndex        =   15
         Top             =   1410
         Width           =   5085
      End
      Begin VB.TextBox txtSubItemName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6465
         TabIndex        =   17
         Top             =   1065
         Width           =   5085
      End
      Begin VB.TextBox txtDeedRegNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1455
         TabIndex        =   10
         Top             =   1335
         Width           =   3180
      End
      Begin VB.TextBox txtShopName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6465
         TabIndex        =   13
         Top             =   705
         Width           =   5085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   525
         TabIndex        =   24
         Top             =   660
         Width           =   885
      End
      Begin VB.Label lblDemandname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building Name"
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
         Left            =   0
         TabIndex        =   5
         Top             =   990
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
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
         Left            =   930
         TabIndex        =   0
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label3 
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
         Left            =   5340
         TabIndex        =   2
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Lessee"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   5280
         TabIndex        =   14
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Item Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   5340
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deed Reg. No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   330
         TabIndex        =   9
         Top             =   1365
         Width           =   1080
      End
      Begin VB.Label lblShopName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shop Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   5640
         TabIndex        =   12
         Top             =   735
         Width           =   795
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2745
      Left            =   270
      TabIndex        =   19
      Top             =   2205
      Width           =   11310
      _cx             =   19950
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
      BackColor       =   -2147483626
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483628
      BackColorAlternate=   -2147483626
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRentOnLandBuildings.frx":004C
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Check1"
         Height          =   225
         Left            =   10755
         TabIndex        =   21
         Top             =   30
         Width           =   255
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsDeedDetails 
      Height          =   2745
      Left            =   45
      TabIndex        =   18
      Top             =   2190
      Visible         =   0   'False
      Width           =   11760
      _cx             =   20743
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
      BackColor       =   -2147483626
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483633
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRentOnLandBuildings.frx":0250
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label lblCaption 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Label2"
      Height          =   285
      Left            =   45
      TabIndex        =   22
      Top             =   1890
      Visible         =   0   'False
      Width           =   11715
   End
End
Attribute VB_Name = "frmRentOnLandBuildings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ************************************************************************************************** '
' Module    : Rent on Land and Buildings Interface to Integrate Sanchaya
' Modified  : 15-Jun-2009
' Database  : DB_Sanchaya and DB_Finance
' Stored Pro: spGetDeedDetails, GetDemandOfARentedBuilding
' ************************************************************************************************** '

Option Explicit
    Public mCategory        As Integer
    
    Dim mDeedRegNo          As Variant
    Dim mNumberOfSelections As Integer
    Dim mDueDay             As Integer
    Dim mAcHeadCodeFine     As String
    Dim mAcHeadCodeRoundOff As String
    Public Property Let Category(mMode As Integer)
        mCategory = mMode
    End Property


    Private Sub chkFineWaiver_Click()
'        If chkFineWaiver.value = vbUnchecked Then
'            Call CalculateFineforRLB
'            Call Calculate
'        Else
'            txtFine.Locked = False
'        End If
        If chkFineWaiver.Value = vbChecked Then
            frmFineWaiver.Mode = 3
            frmFineWaiver.Show vbModal, frmPropertyTax
        Else
        
        End If
    End Sub
    Private Sub chkGrantTotalAdj_Click()
        If chkGrantTotalAdj.Value = vbChecked Then
            If (MsgBox("Do you want to Enter Grand Total", vbYesNo)) = vbYes Then
                chkGrantTotalAdj.Value = vbChecked
                MsgBox "Please Enter the Amount in Grand Total", vbApplicationModal
                vsGrid.Editable = False
                txtGrandTotal.Enabled = True
                txtGrandTotal.Locked = False
                txtGrandTotal.SetFocus
            Else
                chkGrantTotalAdj.Value = vbUnchecked
                txtGrandTotal.Enabled = False
            End If
        Else
            chkSelectAll.Value = vbChecked
            txtGrandTotal.Locked = True
        End If
    End Sub

    Private Sub chkSelectAll_Click()
        If chkGrantTotalAdj.Value = vbChecked Then
            chkSelectAll.Enabled = False
        Else
            If chkSelectAll.Value = vbChecked Then
                If vsGrid.Rows > 1 Then
                    vsGrid.Cell(flexcpChecked, 1, 12, vsGrid.Rows - 1, 12) = True
                    Call Calculate
                End If
            ElseIf chkSelectAll.Value = vbUnchecked Then
                If vsGrid.Rows > 1 Then
                    vsGrid.Cell(flexcpChecked, 1, 12, vsGrid.Rows - 1, 12) = False
                    txtFine.Text = ""
                    txtGrandTotal.Text = ""
                    lblTotalCurrent.Caption = ""
                    lblTotalArrear.Caption = ""
                End If
            End If
            Call CalculateFineforRLB
            Call Calculate
        End If
    End Sub



    Private Sub cmbWard_Click()
        If cmbWard.ListIndex > -1 Then
            txtWardNo.Text = cmbWard.ItemData(cmbWard.ListIndex)
        End If
    End Sub

    Private Sub cmdClose_Click()
        Dim objTranType As New clsTransactionType
        Unload Me
        objTranType.SetTransactionType (9999)
        frmReceiptsCounter.txtTransactionType.Text = objTranType.TransactionType
        frmReceiptsCounter.txtTransactionType.Tag = objTranType.TransactionTypeID
    End Sub

    Private Sub cmdCopyToReceipt_Click()
         Dim objAcc  As New clsAccounts
        Dim mLoop As Integer
        Dim mCount As Integer
        Dim mLoopChild As Integer
        If mDeedRegNo > 0 Then
            frmReceiptsCounter.SubLedgerID = mDeedRegNo
            frmReceiptsCounter.cmbZone.Text = cmbZone.Text
            frmReceiptsCounter.cmbZone.Locked = True
            frmReceiptsCounter.txtWard.Text = cmbWard.Text
            frmReceiptsCounter.txtWard.Locked = True
            frmReceiptsCounter.txtWardNo.Text = val(txtWardNo)
            frmReceiptsCounter.txtBuildingNo = mDeedRegNo
            If cmbWard.ListIndex > -1 Then
                frmReceiptsCounter.txtWard.Tag = cmbWard.ItemData(cmbWard.ListIndex)
            End If
            'frmReceiptsCounter.txtHouseNo.Text = Trim(txtShopName.Text)
            frmReceiptsCounter.txtHouse.Text = Trim(txtMaster.Text)
            frmReceiptsCounter.txtName.Text = txtLesseName.Text
            frmReceiptsCounter.DemandBasedFlag = True
        End If
''         For mLoop = 1 To vsGrid.Rows - 1
''            If vsGrid.Cell(flexcpChecked, mLoop, 12) = 1 Then
''                If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeAdvanceLand Or vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeAdvanceBuilding Then
''                    MsgBox "Advance Posting of Rent on Land/Building is not Implemented", vbApplicationModal
''                    fmeTotal.Visible = True
''                    cmdCopyToReceipt.Enabled = False
''                    Exit Sub
''                End If
''            End If
''        Next

        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 12) = 1 Then
                If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeCivicAmenitiesArrear Or vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeCivicAmenitiesCurrent _
                Or vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeRentLandArrear Or vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeRentLandCurrent Then
                    If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeServiceTax Then
                    
                    ElseIf vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeCGST Then
                        If mLoop + 1 < vsGrid.Rows Then
                            'If vsGrid.TextMatrix(mLoop + 1, 0) <> gbAcHeadCodeServiceTax Then ''' commented on '7 10 2017
'                            If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeServiceTax Then
'                                MsgBox "Please update GST heads in Sanchaya"
'                                Exit Sub
'                            End If
                            If vsGrid.TextMatrix(mLoop + 1, 0) <> gbAcHeadCodeCGST Then
                                MsgBox "Demands is not in Proper Order", vbApplicationModal
                                fmeTotal.Visible = True
                                cmdCopyToReceipt.Enabled = False
                                Exit Sub
                            End If
                        End If
                        If vsGrid.TextMatrix(mLoop, 7) >= 2019 And vsGrid.TextMatrix(mLoop, 8) >= 18 Then
                            If mLoop + 2 < vsGrid.Rows Then
                                'If vsGrid.TextMatrix(mLoop + 1, 0) <> gbAcHeadCodeServiceTax Then ''' commented on '7 10 2017
                                'If vsGrid.TextMatrix(mLoop + 2, 0) <> gbAcHeadCodeSGST Then
                                If vsGrid.TextMatrix(mLoop + 2, 0) <> gbAcHeadCodeFloodCess Then
                                    MsgBox "Demands is not in Proper Order", vbApplicationModal
                                    fmeTotal.Visible = True
                                    cmdCopyToReceipt.Enabled = False
                                    Exit Sub
                                End If
                            End If
                        '''' After Aug 19
                        ''If vsGrid.TextMatrix(mLoop, 7) >= 2019 And vsGrid.TextMatrix(mLoop, 8) >= 18 Then
                            If mLoop + 3 < vsGrid.Rows Then
                                'If vsGrid.TextMatrix(mLoop + 1, 0) <> gbAcHeadCodeServiceTax Then ''' commented on '7 10 2017
                                'If vsGrid.TextMatrix(mLoop + 3, 0) <> gbAcHeadCodeFloodCess Then
                                If vsGrid.TextMatrix(mLoop + 3, 0) <> gbAcHeadCodeSGST Then
                                    MsgBox "Demands is not in Proper Order", vbApplicationModal
                                    fmeTotal.Visible = True
                                    cmdCopyToReceipt.Enabled = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        For mLoop = 1 To vsGrid.Rows - 1
            frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 1
            If vsGrid.Cell(flexcpChecked, mLoop, 12) = 1 Then
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
            End If
        Next
        If val(txtGrandTotal.Text) > val(txtNetAmount.Text) Then
            
            With frmReceiptsCounter.vsGrid
                
                .Rows = .Rows + 1         ' One Row Added
                If mCategory = 1 Then
                    objAcc.SetAccountCode (gbAcHeadCodeAdvanceBuilding)              '' Setting Accounts with Head Code
                Else
                    objAcc.SetAccountCode (gbAcHeadCodeAdvanceLand)
                End If
                .Cell(flexcpText, mCount + 1, 0) = objAcc.AccountCode      'AccountHead
                .Cell(flexcpText, mCount + 1, 1) = objAcc.AccountHead      'AccountHead
                .Cell(flexcpText, mCount + 1, 2) = gbFinancialYearID       'YearID
                .Cell(flexcpText, mCount + 1, 3) = gbCurrentPeriodID       'Period ID
                .Cell(flexcpText, mCount + 1, 5) = val(txtGrandTotal.Text) - val(txtNetAmount.Text) 'Current Amount
                .Cell(flexcpText, mCount + 1, 6) = objAcc.AccountHeadID
                .Cell(flexcpText, mCount + 1, 7) = gbFinancialYearID
                .Cell(flexcpText, mCount + 1, 8) = 1
                .Cell(flexcpText, mCount + 1, 10) = ""
                .Cell(flexcpText, mCount + 1, 11) = val(txtGrandTotal.Text) - val(txtNetAmount.Text) 'Current Amount                       'Amount Paid
                .Cell(flexcpChecked, mCount + 1, 12) = 1
                .Cell(flexcpChecked, mCount + 1, 14) = 0
            End With
        End If
        frmReceiptsCounter.Calculate
        frmReceiptsCounter.txtBuildingNo.Text = txtDeedRegNo.Text
        frmReceiptsCounter.txtDoorNo2.MaxLength = 15
        frmReceiptsCounter.txtDoorNo2.Locked = True
        frmReceiptsCounter.txtStreet.Text = txtSubItemName.Text
        frmReceiptsCounter.txtDoorNo2.Tag = txtSubItemName.Tag
        frmReceiptsCounter.vsGrid.Editable = flexEDNone
        frmReceiptsCounter.txtName.Enabled = False
       
        Unload Me
    End Sub

    Private Sub cmdFind_Click()
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mCnt        As Integer
        Dim arrInput    As Variant
        Dim numMasterID As Variant
        Dim numZoneID   As Variant
        Dim intWardNo   As Variant
        Dim chvShopName As String
        Dim chvLicenceeName As String
        Dim chvSubItemName As String
        Dim numDeedRegNo As Variant

        If cmbZone.ListIndex > -1 Then
            numZoneID = cmbZone.ItemData(cmbZone.ListIndex)
        Else
            MsgBox "Please select Zone", vbInformation
            Exit Sub
        End If
        
        
        If txtMaster.Text <> "" Then
            numMasterID = val(txtMaster.Tag)
        Else
            MsgBox "Please Select Building"
            txtMaster.SetFocus
            Exit Sub
        End If
        If cmbWard.ListIndex > -1 Then
            intWardNo = cmbWard.ItemData(cmbWard.ListIndex)
        Else
'            MsgBox "Please select Ward", vbInformation
'            Exit Sub
             intWardNo = "%"
        End If
        chvShopName = txtShopName.Text + "%"
        chvLicenceeName = txtLesseName.Text + "%"
        chvSubItemName = txtSubItemName.Text + "%"
        If txtDeedRegNo.Text <> "" Then
            numDeedRegNo = txtDeedRegNo.Text
        Else
            numDeedRegNo = Null
        End If
'''        arrInput = Array(numDeedRegNo, _
'''            numMasterID, _
'''            numZoneID, _
'''            intWardNo, _
'''            Trim(chvShopName), _
'''            Trim(chvLicenceeName), _
'''            Trim(chvSubItemName) _
'''            )
            arrInput = Array(numDeedRegNo, _
            numMasterID, _
            numZoneID, _
            Trim(chvShopName), _
            Trim(chvLicenceeName), _
            Trim(chvSubItemName) _
            )
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya)) Then
            'Set Rec = objDB.ExecuteSP("spSanRentSearch2", arrInput, , , mCnn, adCmdStoredProc)
            Set Rec = objdb.ExecuteSP("spSanSnRentSearch", arrInput, , , mCnn, adCmdStoredProc)
            mCnt = 1
            vsDeedDetails.Visible = True
            vsDeedDetails.Clear 1
            vsDeedDetails.Rows = 1
            vsGrid.Visible = False
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF)
                    vsDeedDetails.Rows = vsDeedDetails.Rows + 1
                    If mCnt > 8 Then vsDeedDetails.Rows = vsDeedDetails.Rows + 1
                    vsDeedDetails.TextMatrix(mCnt, 0) = IIf(IsNull(Rec!numDeedRegNo), "", Rec!numDeedRegNo)
                    vsDeedDetails.TextMatrix(mCnt, 1) = IIf(IsNull(Rec!chvShopName), "", Rec!chvShopName)
                    vsDeedDetails.TextMatrix(mCnt, 2) = IIf(IsNull(Rec!chvLicenceeName), "", Rec!chvLicenceeName)
                    vsDeedDetails.TextMatrix(mCnt, 3) = Format(IIf(IsNull(Rec!chvSubItemName), "", Rec!chvSubItemName), "#0")
                    vsDeedDetails.TextMatrix(mCnt, 4) = IIf(IsNull(Rec!dtStartDate), "", Rec!dtStartDate)
                    vsDeedDetails.TextMatrix(mCnt, 5) = IIf(IsNull(Rec!dtEnddate), "", Rec!dtEnddate)
                    vsDeedDetails.TextMatrix(mCnt, 6) = IIf(IsNull(Rec!fltMonthlyLicenseFee), "", Rec!fltMonthlyLicenseFee)
                    vsDeedDetails.TextMatrix(mCnt, 7) = IIf(IsNull(Rec!intDueDay), "", Rec!intDueDay)
                    vsDeedDetails.TextMatrix(mCnt, 9) = IIf(IsNull(Rec!numMasterID), "", Rec!numMasterID)
                    Rec.MoveNext
                    mCnt = mCnt + 1
                Wend
                lblcaption.Caption = "Please Double Click On Selected Building"
                lblcaption.Visible = True
                fmeTotal.Visible = False
            Else
                lblcaption.Caption = "No Item Selected"
                lblcaption.Visible = True
            End If
        cmdCopyToReceipt.Enabled = False
        
        Else
            MsgBox "Didn't able to connect to the Sanchaya Server", vbApplicationModal
        End If
        
    End Sub

    Private Sub cmdMaster_Click()
        Call ClearSearch
        Call FillMaster
        vsMaster.Visible = True
        vsMaster.SetFocus
    End Sub
    Private Sub ClearSearch()
        txtShopName.Text = ""
        txtLesseName.Text = ""
        txtSubItemName.Text = ""
        txtMaster.Text = ""
        txtMaster.Tag = -1
        txtDeedRegNo.Text = ""
        txtDeedRegNo.Tag = -1
        vsDeedDetails.Clear (1)
    End Sub





    Private Sub Form_Load()
        Call FillZone
        Call FillWard
        Call FillGridPeriod
        Call FillCategory
        If mCategory = 1 Then
            lblDemandname.Caption = "Building Name"
            cmbCategory.Text = "Building"
            'cmbCategory.Enabled = False
        ElseIf mCategory = 2 Then
            lblDemandname.Caption = "Land/Bunk Name"
            cmbCategory.Text = "Bunk"
        Else
            lblDemandname.Caption = "Building Name"
        End If
    End Sub
    Private Sub FillCategory()
        Dim mSql As String
        If mCategory = 1 Then
            mSql = "SELECT chvCategoryDetails, intCategoryID FROM snMstCategoryRent Where intAccountHeadCode  in (12)"
        ElseIf mCategory = 2 Then
            mSql = "SELECT chvCategoryDetails, intCategoryID FROM snMstCategoryRent Where intAccountHeadCode  in (18)"
        Else
            mSql = "SELECT chvCategoryDetails, intCategoryID FROM snMstCategoryRent Where intAccountHeadCode  in (12,18)"
        End If
        PopulateList cmbCategory, mSql, , True, , True, enuSourceString.Sanchaya
        
    End Sub
    



    Private Sub txtDeedRegNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub FillGridPeriod()
        Dim mLoop As Integer
        Dim mItem As String
    
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        
        objdb.SetConnection mCnn
        mSql = "Select * From faPeriodicity"
        mItem = "#0; "
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                mItem = mItem & "|#" & Rec!intPeriodicityID & "; " & Rec!vchperiodicity
                Rec.MoveNext
            Wend
        End If
        vsGrid.ColComboList(3) = mItem
    End Sub
    Private Sub txtFine_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtGrandTotal_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtGrandTotal_LostFocus()
        If chkGrantTotalAdj Then
            If txtGrandTotal.Text <> "" Then
                Call ReCheckDemand
            End If
        End If
        If val(txtGrandTotal.Text) > val(lblGrandTotal.Caption) And val(txtGrandTotal.Text) > 0 Then
            cmdCopyToReceipt.Enabled = True
        Else
            cmdCopyToReceipt.Enabled = False
        End If
    End Sub
    Private Sub ReCheckDemand()
        Dim mCnt    As Integer
        Dim mCount  As Integer
        Dim mYear   As Integer
        Dim mPeriod As Integer
        '----------Rechecking of demand After Grand Total Entry
        If vsGrid.Rows > 1 Then
            chkSelectAll.Value = vbUnchecked
            vsGrid.Cell(flexcpChecked, 1, 12, vsGrid.Rows - 1, 12) = 2
        End If
        
'''''            If vsGrid.TextMatrix(mCnt, 0) = gbAcHeadCodeAdvanceLand Or vsGrid.TextMatrix(mCnt, 0) = gbAcHeadCodeAdvanceBuilding Then
'''''                vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
'''''                Call CalculateFineforRLB
'''''                Call Calculate
'''''            ElseIf vsGrid.TextMatrix(mCnt, 0) = gbAcHeadCodeBuildingArrear _
'''''                    Or vsGrid.TextMatrix(mCnt, 0) = gbAcHeadCodeBuildingCurrent _
'''''                    Or vsGrid.TextMatrix(mCnt, 0) = gbAcHeadCodeRentLandArrear _
'''''                    Or vsGrid.TextMatrix(mCnt, 0) = gbAcHeadCodeRentLandCurrent Then
'''''                If mCnt + 1 < vsGrid.Rows Then
'''''                    If vsGrid.TextMatrix(mCnt, 7) = vsGrid.TextMatrix(mCnt + 1, 7) And vsGrid.TextMatrix(mCnt, 8) = vsGrid.TextMatrix(mCnt + 1, 8) Then
'''''                            vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
'''''                            vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = vbChecked
'''''                        If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
'''''                            Call CalculateFineforRLB
'''''                            Call Calculate
'''''                        End If
'''''                        If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
'''''                            mYear = vsGrid.TextMatrix(mCnt, 7)
'''''                            mPeriod = vsGrid.TextMatrix(mCnt, 8)
'''''                            For mCount = 1 To vsGrid.Rows - 1
'''''                                If vsGrid.TextMatrix(mCount, 7) = mYear And vsGrid.TextMatrix(mCount, 8) = mPeriod Then
'''''                                    vsGrid.Cell(flexcpChecked, mCount, 12) = vbUnchecked
'''''                                End If
'''''                                If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodePenalInterest Then
'''''                                    vsGrid.Cell(flexcpChecked, mCount, 12) = vbChecked
'''''                                End If
'''''                            Next
'''''                            If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
'''''                            Call CalculateFineforRLB
'''''                            Call Calculate
'''''
'''''                        End If
'''''                            Exit For
'''''                        End If
'''''
'''''                    End If
'''''                End If
'''''            End If
            
        For mCnt = 1 To vsGrid.Rows - 1
            If mCnt + 1 < vsGrid.Rows Then
'                    --------------Checking Equal demands Period And YearID Should Be eequal
                    If vsGrid.TextMatrix(mCnt, 7) = vsGrid.TextMatrix(mCnt + 1, 7) And vsGrid.TextMatrix(mCnt, 8) = vsGrid.TextMatrix(mCnt + 1, 8) Then
                        If mCnt + 2 < vsGrid.Rows Then
'                           ----if Advance Head Exists
                            If vsGrid.TextMatrix(mCnt, 7) = vsGrid.TextMatrix(mCnt + 2, 7) And vsGrid.TextMatrix(mCnt, 8) = vsGrid.TextMatrix(mCnt + 2, 8) Then
'                                 Checking the Row
                                vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
                                vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = vbChecked
                                vsGrid.Cell(flexcpChecked, mCnt + 2, 12) = vbChecked
                                mCnt = mCnt + 2
                                If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
                                    Call CalculateFineforRLB
                                    Call Calculate
                                End If
                                If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
'                                     UnChecking the Rows
                                    vsGrid.Cell(flexcpChecked, mCnt, 12) = vbUnchecked
                                    vsGrid.Cell(flexcpChecked, mCnt - 1, 12) = vbUnchecked
                                    vsGrid.Cell(flexcpChecked, mCnt - 2, 12) = vbUnchecked
                                    Call CalculateFineforRLB
                                    Call Calculate
                                    Exit For
                                End If
                            Else
'                                 Demand With Out Advance Amount
                                vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
                                vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = vbChecked
                                mCnt = mCnt + 1
                                If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
                                    Call CalculateFineforRLB
                                    Call Calculate
                                End If
                                If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
                                    vsGrid.Cell(flexcpChecked, mCnt, 12) = vbUnchecked
                                    vsGrid.Cell(flexcpChecked, mCnt - 1, 12) = vbUnchecked
                                    Call CalculateFineforRLB
                                    Call Calculate
                                    Exit For
                                End If
                            End If
                        Else
                            vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
                            vsGrid.Cell(flexcpChecked, mCnt + 1, 12) = vbChecked
                            mCnt = mCnt + 1
                            If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
                                Call CalculateFineforRLB
                                Call Calculate
                            End If
                            If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
                                vsGrid.Cell(flexcpChecked, mCnt, 12) = vbUnchecked
                                vsGrid.Cell(flexcpChecked, mCnt - 1, 12) = vbUnchecked
                                Call CalculateFineforRLB
                                Call Calculate
                                Exit For
                            End If
                        End If
                    Else
                        vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
                        If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
                            Call CalculateFineforRLB
                            Call Calculate
                        End If
                        If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
                            vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
                            Call CalculateFineforRLB
                            Call Calculate
                            Exit For
                        End If
                    End If
                Else
                    vsGrid.Cell(flexcpChecked, mCnt, 12) = vbChecked
                    If vsGrid.TextMatrix(mCnt, 0) <> mAcHeadCodeFine Then
                        Call CalculateFineforRLB
                        Call Calculate
                    End If
                    If val(txtGrandTotal.Text) < val(txtNetAmount.Text) Then
                        vsGrid.Cell(flexcpChecked, mCnt, 12) = vbUnchecked
                        Call CalculateFineforRLB
                        Call Calculate
                        Exit For
                    End If
                End If
                For mCount = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodePenalInterest And vsGrid.TextMatrix(mCount, 11) > 0 Then
                       vsGrid.Cell(flexcpChecked, mCount, 12) = vbChecked
                    End If
                Next
        Next
    End Sub
    Private Sub GrandTotalEdit()
       Dim mCnt As Integer
       For mCnt = 1 To vsGrid.Rows - 1
        If vsGrid.TextMatrix(mCnt, 14) = 1 Then
        End If
       Next
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
        'Call ClearSearch
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
        mSql = mSql + " Order By intWardNo ,chvWardNameEnglish"
        PopulateList cmbWard, mSql, , , , True, enuSourceString.DBMaster
    End Sub
    Private Sub FillMaster()
        Dim mSql As String
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mRec As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim numWard As Variant
        Dim mCategory   As Variant
        If cmbWard.ListIndex > -1 Then
            objdb.CreateNewConnection mCnn, enuSourceString.DBMaster
            mSql = "SELECT numWardID From GM_Ward WHERE tnyWardType = 1 And intLBID = " & gbLocalBodyID & " And intWardNo = " & cmbWard.ItemData(cmbWard.ListIndex)
            Set mRec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (mRec.EOF And mRec.BOF) Then
                numWard = IIf(IsNull(mRec!numWardId), "", mRec!numWardId)
            End If
            mCnn.Close
        End If
        If cmbCategory.ListIndex > -1 Then
            mCategory = cmbCategory.ItemData(cmbCategory.ListIndex)
        End If
        mSql = ""
        mSql = "Select chvAssetName,numRegNo  From SnRentAssetDetails "
        mSql = mSql + " WHERE "
        mSql = mSql + " numZoneID = " & cmbZone.ItemData(cmbZone.ListIndex)
        If cmbCategory.ListIndex > -1 Then
            mSql = mSql + " And intCategory = " & mCategory
        End If
        If numWard <> "" Then
            mSql = mSql + " And numWardID = " & numWard
        End If
        
        mSql = mSql + " Order By chvAssetName"
        
        objdb.CreateNewConnection mCnn, enuSourceString.Sanchaya
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        vsMaster.Rows = 1
        If Not (Rec.BOF And Rec.EOF) Then
            vsMaster.Rows = Rec.RecordCount + 1
            vsMaster.Col = 0
            vsMaster.Row = 1
            vsMaster.ColSel = 1
            vsMaster.RowSel = vsMaster.Rows - 1
            mSql = Rec.GetString(, , vbTab, Chr(13))
            vsMaster.Clip = mSql
        End If
        Rec.Close
    End Sub

    Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsDeedDetails_DblClick()
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim arrInput    As Variant
        If (vsDeedDetails.TextMatrix(vsDeedDetails.Row, 0) <> "") And vsDeedDetails.Row > 0 Then
            vsDeedDetails.Visible = False
            vsGrid.Visible = True
            vsGrid.Rows = 1
            chkGrantTotalAdj.Value = vbUnchecked
            vsGrid.Editable = flexEDKbdMouse
            txtGrandTotal.Text = ""
            txtGrandTotal.Locked = True
            arrInput = Array(vsDeedDetails.TextMatrix(vsDeedDetails.Row, 0))
            mDueDay = vsDeedDetails.TextMatrix(vsDeedDetails.Row, 7)
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya)) Then
                'Set Rec = objDB.ExecuteSP("spSanRentDemandMain", arrInput, , , mCnn, adCmdStoredProc)
                Set Rec = objdb.ExecuteSP("spSanSnRentSearch", arrInput, , , mCnn, adCmdStoredProc)
                If Not (Rec.EOF And Rec.BOF) Then
                    txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                    txtShopName.Text = IIf(IsNull(Rec!chvShopName), "", Rec!chvShopName)
                    txtLesseName.Text = IIf(IsNull(Rec!chvLicenceeName), "", Rec!chvLicenceeName)
                    txtDeedRegNo.Text = IIf(IsNull(Rec!numDeedRegNo), "", Rec!numDeedRegNo)
                    txtDeedRegNo.Tag = IIf(IsNull(Rec!numDeedRegNo), "", Rec!numDeedRegNo)
                    txtSubItemName.Text = IIf(IsNull(Rec!chvSubItemName), "", Rec!chvSubItemName)
                    txtSubItemName.Tag = IIf(IsNull(Rec!numSubItemId), "", Rec!numSubItemId)
                    mDeedRegNo = IIf(IsNull(Rec!numDeedRegNo), "", Rec!numDeedRegNo)
                    Call FillGrid(mDeedRegNo)
                    lblcaption.Visible = False
                End If
            Else
                MsgBox "Didn't able to connect to the Sanchaya Server", vbApplicationModal
            End If
            
        End If
    End Sub
    
    Private Sub FillGrid(numDeedRegNo As Variant)
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mRows       As Integer
        Dim arrInput    As Variant
        Dim objAcc      As New clsAccounts
        Dim mYear       As String
        Dim mRLBFine    As Double
        Dim mArrearFlag As Boolean
        mAcHeadCodeFine = gbAcHeadCodePenalInterest
        arrInput = Array(numDeedRegNo)
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya)) Then
             Set Rec = objdb.ExecuteSP("spSanSnRentDemandChild", arrInput, , , mCnn, adCmdStoredProc)
                If Not (Rec.BOF And Rec.EOF) Then
                    vsGrid.Rows = 1
                    mRows = 1
                    While Not (Rec.EOF)
                        vsGrid.Rows = vsGrid.Rows + 1
                        With vsGrid
                             objAcc.SetAccounts (IIf(IsNull(Rec!chvSanHeadCode), -1, Rec!chvSanHeadCode))
                             If (objAcc.AccountCode = gbAcHeadCodeAdvanceLand) Or (objAcc.AccountCode = gbAcHeadCodeAdvanceBuilding) Then
                                MsgBox "Advance Head Exists with Sanchaya Demand", vbApplicationModal
                                Exit Sub
                             End If
                            .TextMatrix(mRows, 0) = objAcc.AccountCode
                            .TextMatrix(mRows, 1) = objAcc.AccountHead
                            If Rec!chvPeriodID - 10 < 4 Then
                            .Cell(flexcpText, mRows, 2) = str(Rec!intYearID - 1) & " - " & str(Rec!intYearID)
                            Else
                            .Cell(flexcpText, mRows, 2) = str(Rec!intYearID) & " - " & str(Rec!intYearID + 1)
                            End If
                            .Cell(flexcpText, mRows, 3) = IIf(IsNull(Rec!chvPeriodID), "", Rec!chvPeriodID)
                            If Rec!ArrearFlag = 0 Then
                                .TextMatrix(mRows, 5) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
                            Else
                                .TextMatrix(mRows, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            End If
                            .TextMatrix(mRows, 6) = objAcc.AccountHeadID
                            If Rec!chvPeriodID - 10 < 4 Then
                                .TextMatrix(mRows, 7) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID) - 1
                            Else
                                .TextMatrix(mRows, 7) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                            End If
                            .Cell(flexcpText, mRows, 8) = IIf(IsNull(Rec!chvPeriodID), "", Rec!chvPeriodID)
                            .Cell(flexcpText, mRows, 9) = IIf(IsNull(Rec!ArrearFlag), "", Rec!ArrearFlag)
                            .Cell(flexcpText, mRows, 11) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            .Cell(flexcpChecked, mRows, 12) = vbChecked
                            Call Calculate
                            .Cell(flexcpText, mRows, 10) = IIf(IsNull(Rec!intKeyID), "", Rec!intKeyID) 'txtDeedRegNo.Tag 'Rec!intKeyID 'Rec!numDemandID
                            If .TextMatrix(mRows, 0) = gbAcHeadCodeAdvanceBuilding Or .TextMatrix(mRows, 0) = gbAcHeadCodeAdvanceLand Then
                                .Cell(flexcpText, mRows, 14) = 1  'To identify Advance
                            End If
                            .Cell(flexcpText, mRows, 15) = mDueDay
                            CalculateFineforRLB
                        End With
                        chkSelectAll.Value = vbChecked
                        Rec.MoveNext
                        mRows = mRows + 1
                    Wend
                    If chkFineWaiver.Value = vbUnchecked Then
                       mRLBFine = CalculateFineforRLB
                    Else
                       mRLBFine = Trim(txtFine.Text)
                    End If
                    If mRLBFine > 0 Then
                        With vsGrid
                            .Rows = vsGrid.Rows + 1
                            objAcc.SetAccountCode (mAcHeadCodeFine)
                            .Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                            .Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                            .Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
                            .Cell(flexcpText, mRows, 3) = Month(Format(gbTransactionDate, "dd/mmm/yy")) + 10
                            .Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                            .Cell(flexcpText, mRows, 7) = gbFinancialYearID
                            .Cell(flexcpText, mRows, 8) = Month(Format(gbTransactionDate, "dd/mmm/yy")) + 10
                            .Cell(flexcpText, mRows, 9) = 0
                            .Cell(flexcpText, mRows, 11) = 0
                            .Cell(flexcpChecked, mRows, 12) = vbChecked
                            .Cell(flexcpText, mRows, 5) = Format(mRLBFine, "#0")
                            mRows = mRows + 1
                        End With
                    End If
                    '--------------------checking whether grid is Complete
                                
                    If GridComplete Then
                        fmeTotal.Visible = True
                        cmdCopyToReceipt.Enabled = False
                        chkFineWaiver.Enabled = False
                    Else
                        fmeTotal.Visible = True
                        If val(txtNetAmount.Text) > 0 Then
                            cmdCopyToReceipt.Enabled = True
                            chkFineWaiver.Enabled = True
                        End If
                    End If
                Else
                    MsgBox "Demand Doesn't Exists", vbApplicationModal
                    fmeTotal.Visible = True
                    cmdCopyToReceipt.Enabled = False
                    chkFineWaiver.Enabled = False
                End If
        Else
            MsgBox "Didn't able to connect to the Sanchaya Server", vbApplicationModal
        End If
    End Sub
    Private Function GridComplete()
        Dim mCnt    As Integer
        Dim mStatus As Boolean
        For mCnt = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCnt, 0) <> "" Then
                If vsGrid.TextMatrix(mCnt, 11) < 0 Then
                    mStatus = True
                    Exit For
                End If
            Else
                If vsGrid.TextMatrix(mCnt, 11) > 0 Then
                    mStatus = True
                    Exit For
                End If
            End If
           Next
           If mStatus Then
                MsgBox "Row is Incomplete"
                GridComplete = True
           End If
    End Function
    Private Sub Calculate()
       Dim mTotal       As Double
       Dim mCurrentAmt  As Double
       Dim mArrearAmt   As Double
       Dim mCount       As Integer
       Dim mFine        As Double
       Dim mAdv         As Double
       Dim mNetAmt      As Double
       Dim mGrantTot    As Double
       '---------- To Find the Arrear Total,Current Total and Grand Total
       For mCount = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mCount, 12) = vbChecked Then
                If vsGrid.TextMatrix(mCount, 0) <> gbAcHeadCodePenalInterest Then
                    'If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeAdvanceBuilding Or vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeAdvanceLand Then
                    If val(vsGrid.Cell(flexcpText, mCount, 14)) = 1 Then
                        mAdv = mAdv + Format(val(vsGrid.Cell(flexcpText, mCount, 11)), "0.00")
                    Else
                        If vsGrid.TextMatrix(mCount, 0) <> "" Then
                            mArrearAmt = mArrearAmt + Format(val(vsGrid.Cell(flexcpText, mCount, 4)), "0.00")
                        End If
                        If vsGrid.TextMatrix(mCount, 0) <> "" Then
                            mCurrentAmt = mCurrentAmt + Format(val(vsGrid.Cell(flexcpText, mCount, 5)), "0.00")
                        End If
                    End If
                End If
            End If
       Next
       lblTotalArrear.Caption = Format(mArrearAmt, "0.00")
       lblTotalCurrent.Caption = Format(mCurrentAmt, "0.00")
        
        If txtFine.Text = "" Then
            mGrantTot = Format(mArrearAmt + mCurrentAmt, "0.00")
        Else
            mGrantTot = Format(mArrearAmt + mCurrentAmt + txtFine.Text, "0.00")
        End If
        'txtGrandTotal.Text = Format(mGrantTot, "0.00")
        lblGrandTotal.Caption = Format(mGrantTot, "0.00")
        If chkGrantTotalAdj.Value = vbChecked Then
            
        End If
        If mAdv > 0 Then
            txtAdvance.Text = Format(mAdv, "0.00")
            mNetAmt = mGrantTot - mAdv
        Else
            mNetAmt = lblGrandTotal.Caption
        End If
        If mAdv < 0 Then
            cmdCopyToReceipt.Enabled = False
        End If
        txtNetAmount = Format(mNetAmt, "0.00")
        For mCount = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodePenalInterest Then
                vsGrid.TextMatrix(mCount, 5) = txtFine.Text
                vsGrid.TextMatrix(mCount, 11) = txtFine.Text
            End If
        Next
    End Sub
    Private Sub vsDeedDetails_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call vsDeedDetails_DblClick
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsGrid_Click()
        If vsGrid.Col = 12 Then
            If chkGrantTotalAdj.Value = vbUnchecked Then
                vsGrid.Editable = flexEDKbdMouse
                If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Or vsGrid.TextMatrix(vsGrid.Row, 0) <> gbAcHeadCodePenalInterest Then
                    Call CalculateFineforRLB
                    Call Calculate
                End If
            End If
        Else
            vsGrid.Editable = flexEDNone
        End If
    End Sub

    Private Sub vsGrid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        Dim mLoop As Long
        Dim mRowCount As Integer
        Dim mYearID As Integer
        Dim mPeriodID   As Integer
        Dim mDemand     As String
        If Row > 0 Then
            If vsGrid.Cell(flexcpChecked, Row, Col) = 2 Then
             For mLoop = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(Row, 7) = vsGrid.TextMatrix(mLoop, 7) And vsGrid.TextMatrix(Row, 8) = vsGrid.TextMatrix(mLoop, 8) Then
                    If Row - 1 <> 0 Then
                        If vsGrid.Cell(flexcpChecked, Row - 1, 12) = 1 Then
                            vsGrid.Cell(flexcpChecked, mLoop, 12) = 1
                            mNumberOfSelections = mNumberOfSelections + 1
                        Else
                            Cancel = True
                        End If
                    Else
                        vsGrid.Cell(flexcpChecked, mLoop, 12) = 1
                        mNumberOfSelections = mNumberOfSelections + 1
                    End If
                End If
                If vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodePenalInterest Then
                    vsGrid.Cell(flexcpChecked, mLoop, 12) = 1
                End If
            Next mLoop
            Else ' Already  Checked
                 If vsGrid.Cell(flexcpChecked, Row - 1, Col) = 1 Then
                    For mLoop = 1 To vsGrid.Rows - 1
                        If vsGrid.TextMatrix(Row, 7) = vsGrid.TextMatrix(mLoop, 7) And vsGrid.TextMatrix(Row, 8) = vsGrid.TextMatrix(mLoop, 8) And vsGrid.TextMatrix(Row, 0) <> gbAcHeadCodePenalInterest Then
                            vsGrid.Cell(flexcpChecked, mLoop, 12) = 2
                            mNumberOfSelections = mNumberOfSelections - 1
                        End If
                    Next mLoop
                    For mLoop = Row To vsGrid.Rows - 1
                        If vsGrid.TextMatrix(mLoop, 0) <> gbAcHeadCodePenalInterest Then
                            vsGrid.Cell(flexcpChecked, mLoop, 12) = 2
                            mNumberOfSelections = mNumberOfSelections - 1
                        Else
                            Cancel = True
                        End If
                    Next mLoop
                Else
                    Cancel = True
                End If
            End If
        End If
    End Sub


    Private Sub vsMaster_DblClick()
        If (vsMaster.TextMatrix(vsMaster.Row, 0) <> "") Then
            txtMaster.Text = vsMaster.TextMatrix(vsMaster.Row, 0) 'Master Name No
            txtMaster.Tag = vsMaster.TextMatrix(vsMaster.Row, 1)  'Master No
            vsMaster.Visible = False
            txtMaster.SetFocus
        End If
    End Sub

    Private Sub vsMaster_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call vsMaster_DblClick
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsMaster_LostFocus()
        vsMaster.Visible = False
 
    End Sub
    
    Private Function CalculateFineforRLB() As Double
        Dim dtFromDt As Date
        Dim mNoOfMonths As Long
        Dim mFineforRLB     As Double
        Dim mCount          As Integer
        Dim mActualFine     As Double
        Dim mYearID         As Integer
        Dim mDue            As Integer
        Dim mPeriodID       As Integer
        Dim mRLB            As Double
        Dim mAdvance        As Double
        Dim mAdvDate        As Date
        Dim mDemandDate     As Date
        Dim mPAmt           As Double
        Dim mSTax           As Double
            mActualFine = 0
            mAdvance = 0
            For mCount = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, mCount, 12) = vbChecked Then
                If vsGrid.TextMatrix(mCount, 0) <> "" And _
                (vsGrid.TextMatrix(mCount, 0) <> gbAcHeadCodePenalInterest) Then
'                    If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeAdvanceBuilding _
'                    Or vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeAdvanceLand Then
'                    Else
                    mPeriodID = vsGrid.TextMatrix(mCount, 8)
                    If mPeriodID - 10 < 4 Then
                        mYearID = vsGrid.TextMatrix(mCount, 7) + 1
                    Else
                        mYearID = vsGrid.TextMatrix(mCount, 7)
                    End If
                    mDue = vsGrid.TextMatrix(mCount, 15)
                    mRLB = vsGrid.TextMatrix(mCount, 11)
                    
''''                    If vsGrid.Cell(flexcpChecked, mCount, 12) = vbChecked Then
''''                        dtFromDt = DateSerial(mYearID, (mPeriodID - 10), mDue)
''''                        If dtFromDt < gbTransactionDate Then
''''                            mNoOfMonths = 1 + DateDiff("m", DdMmmYy(dtFromDt), DdMmmYy(gbTransactionDate)) 'gbTransactionDate))
''''                            mFineforRLB = mRLB * mNoOfMonths / 100
''''                        End If
''''                        mActualFine = mActualFine + mFineforRLB
''''                    End If
'''''                       End If
''''                End If
                  'Advance Collection of Revenues - Rent from Civic Amenities And Leased Land--
                    If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeAdvanceBuilding _
                    Or vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeAdvanceLand Then
                        mAdvance = mAdvance + vsGrid.TextMatrix(mCount, 11)
                        mAdvDate = DateSerial(mYearID, mPeriodID - 10, mDue)
                    End If
                    If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandArrear _
                    Or vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandCurrent _
                    Or vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesCurrent _
                    Or vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesArrear Then
                        mDemandDate = DateSerial(mYearID, mPeriodID - 10, mDue)
                        If mDemandDate < gbTransactionDate Then
                            If gbTransactionDate > mDemandDate Then
                            If mAdvance > 0 Then
                                If mAdvance > mActualFine Then
                                    mAdvance = mAdvance - mActualFine
                                Else
                                    mActualFine = mActualFine - mAdvance
                                    mAdvance = 0
                                End If
                                If mAdvDate <= mDemandDate Then
                                    If mCount + 1 < vsGrid.Rows Then
                                    mPAmt = Format(val(vsGrid.TextMatrix(mCount, 11)) + val(vsGrid.TextMatrix(mCount + 1, 11)), "0.00")
                                    End If
                                    If mAdvance > mPAmt Then
                                        mAdvance = mAdvance - mPAmt
                                    ElseIf mAdvance = 0 Then
                                            mPAmt = Format(val(vsGrid.TextMatrix(mCount, 11)), "0.00")
                                            mNoOfMonths = 1 + DateDiff("m", DdMmmYy(mDemandDate), DdMmmYy(gbTransactionDate))
                                            mFineforRLB = mPAmt * mNoOfMonths / 100
                                            mActualFine = mActualFine + mFineforRLB
                                    Else
                                        mPAmt = mPAmt - mAdvance
                                        mAdvance = 0
                                        mSTax = Format(mPAmt * 10.3 / 100, "0.00")
                                        mPAmt = Format(mPAmt - mSTax, "0.00")
                                        mNoOfMonths = 1 + DateDiff("m", DdMmmYy(mDemandDate), DdMmmYy(gbTransactionDate))
                                        mFineforRLB = Format(mPAmt * mNoOfMonths / 100, "0.00")
                                        mActualFine = mActualFine + mFineforRLB
                                    End If
                                End If
                            Else
                                mPAmt = val(vsGrid.TextMatrix(mCount, 11))
                                mNoOfMonths = 1 + DateDiff("m", DdMmmYy(mDemandDate), DdMmmYy(gbTransactionDate))
                                mFineforRLB = mPAmt * mNoOfMonths / 100
                                mActualFine = mActualFine + mFineforRLB
                            End If
                            End If
                        End If
                    End If
                End If
                End If
            Next
            CalculateFineforRLB = mActualFine
            txtFine.Text = mActualFine
    End Function
