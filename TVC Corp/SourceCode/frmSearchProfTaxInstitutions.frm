VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchProfTaxInstitutions 
   BackColor       =   &H00F2FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Profession Tax - Institutions"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDemand 
      BackColor       =   &H00F2FFFF&
      Height          =   5700
      Left            =   4845
      TabIndex        =   37
      Top             =   915
      Visible         =   0   'False
      Width           =   8505
      Begin VB.CheckBox chkFineWaiver 
         Caption         =   "Fine Waiver"
         Height          =   285
         Left            =   3330
         TabIndex        =   56
         Top             =   4500
         Width           =   1230
      End
      Begin VB.TextBox txtDemandTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6780
         MaxLength       =   12
         TabIndex        =   54
         Top             =   4860
         Width           =   1380
      End
      Begin VB.TextBox txtCurrentTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6780
         MaxLength       =   12
         TabIndex        =   52
         Top             =   4500
         Width           =   1380
      End
      Begin VB.TextBox txtArrearTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5370
         MaxLength       =   12
         TabIndex        =   50
         Top             =   4500
         Width           =   1380
      End
      Begin VB.TextBox txtFine 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   49
         Top             =   4500
         Width           =   1035
      End
      Begin VB.CommandButton cmdCopyToReceipt 
         Caption         =   "Copy to Receipt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2625
         TabIndex        =   41
         Top             =   4980
         Width           =   1575
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridDemand 
         Height          =   4125
         Left            =   90
         TabIndex        =   38
         Top             =   390
         Width           =   8460
         _cx             =   14922
         _cy             =   7276
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         BackColorAlternate=   -2147483626
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   2
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSearchProfTaxInstitutions.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         Begin VB.CheckBox chkSelectAll 
            Caption         =   "Check1"
            Height          =   195
            Left            =   8085
            TabIndex        =   55
            Top             =   30
            Width           =   225
         End
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   6210
         TabIndex        =   53
         Top             =   4890
         Width           =   405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   1905
         TabIndex        =   51
         Top             =   4545
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   4590
         TabIndex        =   48
         Top             =   4530
         Width           =   660
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   40
         Top             =   150
         Width           =   1890
      End
   End
   Begin VB.Frame fraInstitutionMaster 
      BackColor       =   &H00F2FFFF&
      Height          =   6825
      Left            =   4605
      TabIndex        =   44
      Top             =   -90
      Width           =   7890
      Begin VSFlex8LCtl.VSFlexGrid vsGridLeft 
         Height          =   5595
         Left            =   180
         TabIndex        =   45
         Top             =   570
         Width           =   6975
         _cx             =   12303
         _cy             =   9869
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         BackColorAlternate=   15398903
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSearchProfTaxInstitutions.frx":01E9
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   -1  'True
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ I n s t i t u t i o n       M a s t e r ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1935
         TabIndex        =   47
         Top             =   285
         Width           =   2625
      End
      Begin VB.Label lblCount 
         Caption         =   "#"
         Height          =   225
         Left            =   5775
         TabIndex        =   46
         Top             =   6225
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F2FFFF&
      Height          =   150
      Left            =   105
      TabIndex        =   43
      Top             =   4845
      Width           =   4395
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F2FFFF&
      Height          =   6825
      Left            =   0
      TabIndex        =   15
      Top             =   -90
      Width           =   4590
      Begin VB.ComboBox cmbZone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   615
         Width           =   1965
      End
      Begin VB.TextBox txtWard 
         Height          =   285
         Left            =   1245
         MaxLength       =   2
         TabIndex        =   1
         Top             =   960
         Width           =   525
      End
      Begin VB.TextBox txtDoorNo1 
         Height          =   285
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1290
         Width           =   1080
      End
      Begin VB.TextBox txtPhone 
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
         Left            =   1230
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3540
         Width           =   1605
      End
      Begin VB.TextBox txtPost 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3225
         Width           =   1605
      End
      Begin VB.TextBox txtMainPlace 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2910
         Width           =   3210
      End
      Begin VB.TextBox txtLocalPlace 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2595
         Width           =   3210
      End
      Begin VB.TextBox txtStreet 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2265
         Width           =   3210
      End
      Begin VB.TextBox txtHouseName 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1935
         Width           =   3210
      End
      Begin VB.TextBox txtInstName 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1605
         Width           =   3210
      End
      Begin VB.TextBox txtOwnersName 
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3885
         Width           =   3210
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   570
         TabIndex        =   13
         Top             =   4530
         Width           =   1380
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   2010
         TabIndex        =   14
         Top             =   4530
         Width           =   1380
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F2FFFF&
         Height          =   150
         Left            =   75
         TabIndex        =   27
         Top             =   390
         Width           =   4395
      End
      Begin VB.TextBox txtDoorNo2 
         Height          =   285
         Left            =   2355
         TabIndex        =   4
         Top             =   1290
         Width           =   525
      End
      Begin VB.ComboBox cmbWard 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1425
      End
      Begin VB.Frame fraAmount 
         BackColor       =   &H00F2FFFF&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   255
         TabIndex        =   28
         Top             =   4950
         Width           =   4245
         Begin VB.CommandButton cmdCopyToDemand 
            Caption         =   "Copy to Demand"
            Height          =   375
            Left            =   1290
            TabIndex        =   39
            Top             =   1260
            Width           =   1380
         End
         Begin VB.TextBox txtAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            MaxLength       =   12
            TabIndex        =   32
            Top             =   240
            Width           =   1080
         End
         Begin VB.TextBox txtNoOfEmp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3315
            MaxLength       =   10
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtEmpName 
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
            Left            =   960
            MaxLength       =   100
            TabIndex        =   30
            Top             =   540
            Width           =   3210
         End
         Begin VB.TextBox txtDesignation 
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
            Left            =   960
            MaxLength       =   100
            TabIndex        =   29
            Top             =   870
            Width           =   3210
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   375
            TabIndex        =   36
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Of Emp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   2535
            TabIndex        =   35
            Top             =   285
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Emp. Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   120
            TabIndex        =   34
            Top             =   585
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   75
            TabIndex        =   33
            Top             =   930
            Width           =   840
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Profession Tax - Traders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   42
         Top             =   210
         Width           =   4410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   45
         TabIndex        =   26
         Top             =   1605
         Width           =   1125
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House/Office"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   195
         TabIndex        =   25
         Top             =   1980
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   720
         TabIndex        =   24
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   330
         TabIndex        =   23
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   390
         TabIndex        =   22
         Top             =   2955
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   825
         TabIndex        =   21
         Top             =   3270
         Width           =   315
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   465
         TabIndex        =   20
         Top             =   3600
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner's Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   105
         TabIndex        =   19
         Top             =   3885
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   585
         TabIndex        =   18
         Top             =   1290
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   795
         TabIndex        =   17
         Top             =   990
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   810
         TabIndex        =   16
         Top             =   645
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSearchProfTaxInstitutions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private mProfTaxInstTypeMode As Integer  ' 1= Trader 2=Employees 3=Self Drawing Officer
    Dim arrInput As Variant
    Dim oldArrayIn As Variant
    Dim mInstitutionID As Variant
    Dim mYearID As Integer
    Dim mPeriodID As Integer
    Dim mFineAmt As Double
    Dim mAdvAmt     As Double   ' Total Advance Amount
    Dim dtUptoDate  As Date     ' Fine Upto Date
    Dim dtFromDate  As Date
    Dim mNumberOfSelections As Integer
    
    Private Sub chkFineWaiver_Click()
        If chkFineWaiver.value = vbChecked Then
            frmFineWaiver.Mode = 4
            frmFineWaiver.Show vbModal, frmPropertyTax
        End If
    End Sub

    Private Sub chkSelectAll_Click()
        If chkSelectAll.value = vbChecked Then
            If vsGridDemand.Rows > 1 Then
                vsGridDemand.Cell(flexcpChecked, 1, 12, vsGridDemand.Rows - 1, 12) = True
                Call Calculate
                Call calculateFine
            End If
        ElseIf chkSelectAll.value = vbUnchecked Then
            If vsGridDemand.Rows > 1 Then
                vsGridDemand.Cell(flexcpChecked, 1, 12, vsGridDemand.Rows - 1, 12) = False
                txtFine.Text = ""
                txtArrearTotal.Text = ""
                txtCurrentTotal.Text = ""
                txtDemandTotal.Text = ""
            End If
        End If
        
        Call Calculate
        Call calculateFine
        If chkSelectAll.value = vbUnchecked Then
            txtFine.Text = ""
        End If
    End Sub
    Private Sub cmbWard_Click()
        If cmbWard.ListIndex > -1 Then
            txtWard.Text = cmbWard.ItemData(cmbWard.ListIndex)
        End If
    End Sub

    Private Sub cmdClear_Click()
'        Call FormInitialise
        InputData (False)
    End Sub
    Private Sub cmdCopyToDemand_Click()
        Call copyToDemand
    End Sub

    Private Sub cmdCopyToReceipt_Click()
        Call copyToReceipt
    End Sub

    Private Sub cmdSearch_Click()
        fraDemand.Visible = False
        fraInstitutionMaster.Visible = True
        Call FillInstitutions
    End Sub
    Private Sub Form_Activate()
        Me.Left = frmMenu.Left + 10
    End Sub
    Private Sub Form_Load()
        ReDim arrInput(13)
        Call formInitialise
        Call FillZone
        Call FillWard
        txtFine.Locked = True
        txtArrearTotal.Locked = True
        txtCurrentTotal.Locked = True
        txtDemandTotal.Locked = True
    End Sub
    Private Sub formInitialise()
        Dim ctrl As Control
        Dim mRowCnt As Integer
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            ElseIf TypeOf ctrl Is ComboBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                ctrl.Tag = ""
            End If
        Next
        vsGridDemand.Visible = False
        If mProfTaxInstTypeMode = 1 Then
            fraAmount.Visible = False
            lblTitle.Caption = "Profession Tax - Traders"
        Else
            fraAmount.Visible = True
            lblTitle.Caption = "Profession Tax - Employees"
        End If
        
    End Sub
    '**************************************************************************'
    '           To Fill the Institution Details                                '
    '**************************************************************************'
    Private Sub FillInstitutions()
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mRowCnt As Integer
        Dim mRecCnt As Integer
        
        Dim numZoneID As Variant
        Dim numWardId As Variant
        Dim intInstitutionTypeID As Variant
        Dim mInputArray As Variant
        
        If cmbZone.ListIndex > -1 Then
            numZoneID = cmbZone.ItemData(cmbZone.ListIndex)
        End If
        If cmbWard.ListIndex > -1 Then
            numWardId = cmbWard.ItemData(cmbWard.ListIndex)
        End If
        Select Case ProfTaxInstTypeMode 'mProfTaxInstTypeMode
            Case Is = intInstitutionTypeID = 1   ' Traders
            Case Is = intInstitutionTypeID = 2   ' Employees
            Case Is = intInstitutionTypeID = 3   ' Self Drawing
            Case Else: intInstitutionTypeID = 1
        End Select
        
        If Trim(txtInstName.Text) = "" Then
            txtInstName.Tag = ""
        End If
        
        'mInputArray = Array(IIf(val(txtInstName.Tag) > 0, txtInstName.Tag, Null), _'
        mInputArray = Array(IIf(val(txtInstName.Tag) > 0, txtInstName.Tag, Null), _
                    numZoneID, _
                    intInstitutionTypeID, _
                    numWardId, _
                    txtDoorNo1.Text, _
                    txtDoorNo2.Text, _
                    txtInstName.Text, _
                    txtOwnersName.Text, _
                    txtHouseName.Text, _
                    txtStreet.Text, _
                    txtLocalPlace.Text, _
                    txtMainPlace.Text)
        
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Sanchaya)) Then
            Set Rec = objDB.ExecuteSP("spSanSnProfTaxSearch", mInputArray, , , mCnn, adCmdStoredProc)
            mRowCnt = 1
            mRecCnt = 1
            vsGridLeft.Rows = 1
            lblCount.Caption = "Count: 0"
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    If vsGridLeft.Rows = mRowCnt Then vsGridLeft.Rows = vsGridLeft.Rows + 15
                    vsGridLeft.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!chvInstName), "", Rec!chvInstName)
                    vsGridLeft.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvMainPlaceEng), "", Rec!chvMainPlaceEng)
                    vsGridLeft.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
                    vsGridLeft.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!chvOwnersEng), "", Rec!chvOwnersEng)
                    vsGridLeft.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!chvDoorNos), "", Rec!chvDoorNos)
                    vsGridLeft.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!numInstID), "", Rec!numInstID)
                    vsGridLeft.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!intInstTypeHeadId), "", Rec!intInstTypeHeadId)
                    vsGridLeft.TextMatrix(mRowCnt, 7) = mRecCnt
                    vsGridLeft.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                    vsGridLeft.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                    mInstitutionID = vsGridLeft.TextMatrix(mRowCnt, 5)
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                    mRecCnt = mRecCnt + 1
                Wend
                lblCount.Caption = "Count:" & mRecCnt - 1
            End If
            Rec.Close
        End If
       
    End Sub
    Private Sub FillZone()
        Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Where intLBID = " & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, enuSourceString.DBMaster)
    End Sub
    Private Sub FillWard()
        Dim mSQL As String
        On Error Resume Next
        mSQL = "SELECT chvWardNameEnglish, intWardNo, numWardID FROM GM_Ward"
        mSQL = mSQL + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID
        mSQL = mSQL + " AND numZoneID = " & cmbZone.ItemData(cmbZone.ListIndex)
        mSQL = mSQL + " Order By intWardNo ,chvWardNameEnglish"
        PopulateList cmbWard, mSQL, , , , True, enuSourceString.DBMaster
   End Sub
   Public Sub InputData(mSaveFlag As Boolean)
    If mSaveFlag Then
        ReDim arrInput(13)
        arrInput(0) = cmbZone.Text
        arrInput(1) = cmbWard.Text
        arrInput(2) = txtWard.Text
        arrInput(3) = txtDoorNo1.Text
        arrInput(4) = txtDoorNo2.Text
        arrInput(5) = txtInstName.Text
        arrInput(6) = txtHouseName.Text
        arrInput(7) = txtStreet.Text
        arrInput(8) = txtLocalPlace.Text
        arrInput(9) = txtMainPlace.Text
        arrInput(10) = txtPost.Text
        arrInput(11) = txtPhone.Text
        arrInput(12) = txtOwnersName.Text
    Else
       If Trim(arrInput(0)) <> "" Then cmbZone.Text = arrInput(0)
        If Trim(arrInput(1)) <> "" Then cmbWard.Text = arrInput(1)
        'txtWard.Text = arrInput(2)
        txtDoorNo1.Text = arrInput(3)
        txtDoorNo2.Text = arrInput(4)
        txtInstName.Text = arrInput(5)
        txtHouseName.Text = arrInput(6)
        txtStreet.Text = arrInput(7)
        txtLocalPlace.Text = arrInput(8)
        txtMainPlace.Text = arrInput(9)
        txtPost.Text = arrInput(10)
        txtPhone.Text = arrInput(11)
        txtOwnersName.Text = arrInput(12)
        
        ReDim arrInput(13)
    End If
End Sub
    '**************************************************************************'
    '           To Fill the Institution Details from Grid to TextBoxes         '
    '**************************************************************************'
    Private Sub DisplayInstitutionDetails(InstID As Variant)
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSQL As String
        Dim mRCnt As Integer
        Dim mArrDoorNo As Variant
        Dim mArrDoorno1 As Variant
        Dim mArrDoorNo2 As Variant
        Dim mCount As Integer
        Dim numZoneID   As Variant
        Dim numWardId   As Integer
        Dim intInstitutionTypeID As Variant
        
        Dim mInputArray As Variant
        
        'Note:- Store Search Parameters into Array
        Call InputData(True)
        
        If cmbZone.ListIndex > -1 Then
            numZoneID = cmbZone.ItemData(cmbZone.ListIndex)
        End If
        If cmbWard.ListIndex > -1 Then
            numWardId = cmbWard.ItemData(cmbWard.ListIndex)
        End If
        Select Case mProfTaxInstTypeMode
            Case Is = 15: intInstitutionTypeID = 1  ' Traders
            Case Is = 16: intInstitutionTypeID = 2  ' Employees
            Case Is = 17: intInstitutionTypeID = 3  ' Self Drawing
            Case Else: intInstitutionTypeID = 1
        End Select
        Dim s As String
        mInputArray = Array(InstID)
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Sanchaya)) Then
            Set Rec = objDB.ExecuteSP("spSanSnProfTaxSearch", mInputArray, , , mCnn, adCmdStoredProc)
            If Not (Rec.EOF And Rec.BOF) Then
                txtInstName.Text = IIf(IsNull(Rec!chvInstName), "", Rec!chvInstName)
         
'                If ((vsGridLeft.TextMatrix(vsGridLeft.Row, 4) <> "" Or _
'                    vsGridLeft.TextMatrix(vsGridLeft.Row, 4) <> vbNullString Or _
'                    vsGridLeft.TextMatrix(vsGridLeft.Row, 4) <> vbNull)) Then
                  If InStr(vsGridLeft.TextMatrix(vsGridLeft.Row, 4), "/") > 0 Then
                    mArrDoorNo = Split(vsGridLeft.TextMatrix(vsGridLeft.Row, 4), "/")
                    txtDoorNo1.Text = mArrDoorNo(1)
                    If InStr(1, mArrDoorNo(1), "(") > 0 Then
                        mArrDoorno1 = Split(mArrDoorNo(1), "(")
                        txtDoorNo1.Text = mArrDoorno1(0)
                        
                        mArrDoorNo2 = Split(mArrDoorno1(1), ")")
                        txtDoorNo2.Text = mArrDoorNo2(0)
                    End If
'                Else
'                    txtDoorNo1.Text = ""
'                    txtDoorNo2.Text = ""
                End If
                txtWard.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                txtStreet.Text = IIf(IsNull(Rec!chvStreetName), "", Rec!chvStreetName)
                txtLocalPlace.Text = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
                txtMainPlace.Text = IIf(IsNull(Rec!chvMainPlaceEng), "", Rec!chvMainPlaceEng)
                txtInstName.Tag = IIf(IsNull(Rec!numInstID), "", Rec!numInstID)
                txtPhone = IIf(IsNull(Rec!chvLandPhone), "", Rec!chvLandPhone)
                txtOwnersName = IIf(IsNull(Rec!chvOwnersEng), "", Rec!chvOwnersEng)
            End If
        Else
            MsgBox "Didn't able to connect to the Sanchaya Server", vbApplicationModal
        End If
        
        cmbZone.ListIndex = -1
        For mCount = 0 To cmbZone.ListCount - 1
            If val(vsGridLeft.TextMatrix(vsGridLeft.Row, 8)) = cmbZone.ItemData(mCount) Then
                cmbZone.ListIndex = mCount
                Exit For
            End If
        Next
        Rec.Close
    End Sub
    
   Private Sub copyToDemand()
    Dim mCnn As New ADODB.Connection
    Dim Rec  As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim objAcc  As New clsAccounts
    Dim mSQL As String
    Dim mRCnt As Integer
    
    '-------------------------------------------------------------------------------------------'
    '                                   Validations                                             '
    '-------------------------------------------------------------------------------------------'
        If Trim(txtAmount.Text) = "" Or Trim(txtAmount.Text) = 0 Then
            MsgBox "Please Enter the Amount"
            txtAmount.SetFocus
            Exit Sub
        End If
        
        If Trim(txtNoOfEmp.Text) = "" Then
            MsgBox "Please Enter the number of Employee"
            txtNoOfEmp.SetFocus
            Exit Sub
        End If
        If val(txtNoOfEmp.Text) = 1 Then
            If Trim(txtEmpName.Text) = "" Then
                MsgBox "Please enter the Name of Employee"
                txtEmpName.SetFocus
                Exit Sub
            End If
        End If
        
        objAcc.SetAccountCode (gbAcHeadCodeProfTaxEmployees)
        
            If txtInstName.Tag <> "" Then
                frmDemandInterface.cmbZone = cmbZone.Text
                frmDemandInterface.txtWardNo = txtWard.Text
                frmDemandInterface.txtDoorNo1 = txtDoorNo1.Text
                frmDemandInterface.txtDoorNo2 = txtDoorNo2.Text
                frmDemandInterface.txtMainPlace = txtMainPlace.Text
                frmDemandInterface.txtLocalPlace = txtLocalPlace.Text
                frmDemandInterface.vsGrid.TextMatrix(1, 0) = objAcc.AccountCode
                frmDemandInterface.vsGrid.TextMatrix(1, 1) = objAcc.AccountHead
                frmDemandInterface.vsGrid.TextMatrix(1, 5) = txtAmount.Text
                frmDemandInterface.txtName = txtEmpName.Text
                frmDemandInterface.vsGrid.Editable = flexEDNone
            Else
                MsgBox "Please select an Institution", vbInformation
                Exit Sub
            End If
       cmdCopyToDemand.Enabled = False
       Unload Me
       frmDemandInterface.Visible = True
    End Sub
    Private Sub copyToReceipt()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDB   As New clsDB
        Dim objAcc  As New clsAccounts
        Dim mSQL    As String
        Dim mRCnt   As Integer
        Dim mLoop  As Integer
        Dim mYearID As Integer
        Dim mPeriodID As Integer
        Dim mUptoDate As Date
        Dim mProfTax As Double
       
        If txtInstName.Tag <> "" Then
            frmReceiptsCounter.cmbZone = cmbZone.Text
            frmReceiptsCounter.txtWardNo = txtWard.Text
            frmReceiptsCounter.txtDoorNo1 = txtDoorNo1.Text
            frmReceiptsCounter.txtDoorNo2 = txtDoorNo2.Text
            frmReceiptsCounter.txtMainPlace = txtMainPlace.Text
            frmReceiptsCounter.txtLocalPlace = txtLocalPlace.Text
            frmReceiptsCounter.txtName = txtOwnersName.Text
            frmReceiptsCounter.txtHouse = txtInstName.Text
            frmReceiptsCounter.txtHouse.Tag = txtInstName.Tag
            
            frmReceiptsCounter.cmbZone.Enabled = False
            frmReceiptsCounter.txtWardNo.Enabled = False
            frmReceiptsCounter.txtDoorNo1.Enabled = False
            frmReceiptsCounter.txtDoorNo2.Enabled = False
            frmReceiptsCounter.txtMainPlace.Enabled = False
            frmReceiptsCounter.txtLocalPlace.Enabled = False
            frmReceiptsCounter.txtName.Enabled = False
            frmReceiptsCounter.txtHouse.Enabled = False
'            frmReceiptsCounter.
            mRCnt = 1
            frmReceiptsCounter.vsGrid.Rows = 1
            For mLoop = 1 To vsGridDemand.Rows - 1
                If vsGridDemand.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
                    mRCnt = mRCnt + 1
                    frmReceiptsCounter.vsGrid.Rows = mRCnt
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 0) = vsGridDemand.TextMatrix(mLoop, 0)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 1) = vsGridDemand.TextMatrix(mLoop, 1)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 2) = vsGridDemand.TextMatrix(mLoop, 2)
                    frmReceiptsCounter.vsGrid.Cell(flexcpText, mRCnt - 1, 3) = vsGridDemand.TextMatrix(mLoop, 8)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 4) = vsGridDemand.TextMatrix(mLoop, 4)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 5) = vsGridDemand.TextMatrix(mLoop, 5)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 6) = vsGridDemand.TextMatrix(mLoop, 6)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 7) = vsGridDemand.TextMatrix(mLoop, 7)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 8) = vsGridDemand.TextMatrix(mLoop, 8)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 9) = vsGridDemand.TextMatrix(mLoop, 9)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 10) = vsGridDemand.TextMatrix(mLoop, 10)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 11) = vsGridDemand.TextMatrix(mLoop, 11)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 12) = vsGridDemand.TextMatrix(mLoop, 12)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 13) = vsGridDemand.TextMatrix(mLoop, 13)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 14) = vsGridDemand.TextMatrix(mLoop, 14)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 15) = vsGridDemand.TextMatrix(mLoop, 15)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRCnt - 1, 16) = vsGridDemand.TextMatrix(mLoop, 16)
                    frmReceiptsCounter.vsGrid.Editable = flexEDNone
                End If
            Next
        End If
       cmdCopyToReceipt.Enabled = False
       Unload Me
       Unload frmDemandInterface
       frmReceiptsCounter.Visible = True
        
    End Sub
    Public Property Let ProfTaxInstTypeMode(mData As Integer)
        mProfTaxInstTypeMode = mData
    End Property

    Public Property Get ProfTaxInstTypeMode() As Integer
        ProfTaxInstTypeMode = mProfTaxInstTypeMode
    End Property

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtArrearTotal_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtCurrentTotal_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtDemandTotal_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtFine_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
'        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
'            KeyAscii = 0
'        End If
    End Sub
    Private Sub txtNoOfEmp_Change()
    If IsNumeric(txtNoOfEmp.Text) Then
        If vsGridLeft.IsSelected(vsGridLeft.Row) And vsGridLeft.TextMatrix(vsGridLeft.Row, vsGridLeft.Col) <> "" Then
            If val(txtNoOfEmp.Text) = 1 Then
                txtEmpName.Enabled = True
                txtDesignation.Enabled = True
            ElseIf val(txtNoOfEmp.Text) > 1 Then
                txtEmpName.Text = ""
                txtDesignation.Text = ""
                txtEmpName.Enabled = False
                txtDesignation.Enabled = False
            End If
        End If
    End If
    End Sub
    Private Sub txtNoOfEmp_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtPhone_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtPost_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtWard_Change()
        Dim mCount As Integer
        cmbWard.ListIndex = -1
        For mCount = 0 To cmbWard.ListCount - 1
            If val(txtWard.Text) = cmbWard.ItemData(mCount) Then
                cmbWard.ListIndex = mCount
                Exit For
            End If
        Next
    End Sub
    Private Sub txtWard_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub FillInstitutionDetails()
        Dim mRowCnt As Integer
        Dim numZoneID As Variant
        Dim numWardId As Variant
        Dim intInstitutionTypeID As Variant
        Dim mRecCnt As Integer
        
        If cmbZone.ListIndex > -1 Then
            numZoneID = cmbZone.ItemData(cmbZone.ListIndex)
        End If
        If cmbWard.ListIndex > -1 Then
            numWardId = cmbWard.ItemData(cmbWard.ListIndex)
        End If
        Select Case ProfTaxInstTypeMode 'mProfTaxInstTypeMode
            Case Is = intInstitutionTypeID = 1   ' Traders
            Case Is = intInstitutionTypeID = 2   ' Employees
            Case Is = intInstitutionTypeID = 3   ' Self Drawing
            Case Else: intInstitutionTypeID = 1
        End Select
        
        If Trim(txtInstName.Text) = "" Then
            txtInstName.Tag = ""
        End If
        
        arrInput = Array(IIf(val(txtInstName.Tag) > 0, txtInstName.Tag, Null), _
                    numZoneID, _
                    intInstitutionTypeID, _
                    numWardId, _
                    txtDoorNo1.Text, _
                    txtDoorNo2.Text, _
                    txtInstName.Text, _
                    txtOwnersName.Text, _
                    txtHouseName.Text, _
                    txtStreet.Text, _
                    txtLocalPlace.Text, _
                    txtMainPlace.Text)
    End Sub
    '**************************************************************************'
    '           To Fill the Demand Details on Grid                             '
    '**************************************************************************'
    Private Sub FillDemandDetailsGrid()
        Dim mRowCnt As Integer
        Dim numZoneID As Variant
        Dim numWardId As Variant
        Dim intInstitutionTypeID As Variant
        Dim mRecCnt As Integer
        Dim arrIn As Variant
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim objAcc As New clsAccounts
        Dim mSQL As String
        Dim mPenalPeriodID As Integer
        
         If Trim(txtInstName.Text) = "" Then
            txtInstName.Tag = ""
        End If
        
        arrIn = Array(IIf(val(txtInstName.Tag) > 0, txtInstName.Tag, Null))
        dtUptoDate = gbTransactionDate
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Sanchaya)) Then
            Set Rec = objDB.ExecuteSP("spSanSnProfTaxDemand", arrIn, , , mCnn, adCmdStoredProc)
           vsGridDemand.Clear 1, 1
            mRowCnt = 1
            mRecCnt = 1
            vsGridDemand.Rows = 1
            lblCount.Caption = "Count: 0"
            vsGridDemand.Tag = -1
           
            If Rec.State <> 0 Then
                  While Not (Rec.EOF Or Rec.BOF)
                    If vsGridDemand.Rows = mRowCnt Then vsGridDemand.Rows = vsGridDemand.Rows + 1
                    vsGridDemand.TextMatrix(mRowCnt, 17) = IIf(IsNull(Rec!chvSanHeadCode), "", Rec!chvSanHeadCode) 'Poornima
                    If vsGridDemand.TextMatrix(mRowCnt, 17) = 431190101 Then
                        objAcc.SetAccountCode (gbAcHeadCodeProfTaxTradersCurrent)
                    ElseIf vsGridDemand.TextMatrix(mRowCnt, 17) = 431190102 Then
                        objAcc.SetAccountCode (gbAcHeadCodeProfTaxTradersArrears) 'Poornima
                    End If
                    vsGridDemand.TextMatrix(mRowCnt, 0) = objAcc.AccountCode
                    vsGridDemand.TextMatrix(mRowCnt, 1) = objAcc.AccountHead
                    vsGridDemand.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                    vsGridDemand.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!chvPeriodID), "", Rec!chvPeriodID)
                    
                    mYearID = vsGridDemand.TextMatrix(mRowCnt, 7)
                    mPeriodID = vsGridDemand.TextMatrix(mRowCnt, 8)
                    vsGridDemand.Cell(flexcpText, mRowCnt, 2) = CStr(mYearID) & "-" & CStr(mYearID + 1)
                    If mPeriodID = 2 Then
                        vsGridDemand.Cell(flexcpText, mRowCnt, 3) = "IInd Half"
                    ElseIf mPeriodID = 1 Then
                        vsGridDemand.Cell(flexcpText, mRowCnt, 3) = "Ist Half"
                    End If
                    vsGridDemand.Tag = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    If Rec!ArrearFlag = 1 Then
                        vsGridDemand.TextMatrix(mRowCnt, 4) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
                    Else
                        vsGridDemand.TextMatrix(mRowCnt, 5) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
                    End If
                    vsGridDemand.TextMatrix(mRowCnt, 6) = objAcc.AccountHeadID
                    vsGridDemand.TextMatrix(mRowCnt, 9) = ""
                    vsGridDemand.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!intKeyID), "", Rec!intKeyID)
                    vsGridDemand.TextMatrix(mRowCnt, 11) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
                    vsGridDemand.TextMatrix(mRowCnt, 12) = vbChecked
                    vsGridDemand.TextMatrix(mRowCnt, 13) = ""
                    vsGridDemand.TextMatrix(mRowCnt, 14) = ""
                    vsGridDemand.TextMatrix(mRowCnt, 15) = ""
                    vsGridDemand.TextMatrix(mRowCnt, 16) = ""
                    
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                    mRecCnt = mRecCnt + 1
                Wend
          Else
            MsgBox "Demand Doesn't Exists", vbApplicationModal
          End If
          Call calculateFine
''            Rec.Close
            If val(txtFine.Text) > 0 Then
                objAcc.SetAccountCode gbAcHeadCodePenalInterest
                vsGridDemand.Rows = vsGridDemand.Rows + 1
                vsGridDemand.TextMatrix(mRowCnt, 0) = objAcc.AccountCode
                vsGridDemand.TextMatrix(mRowCnt, 1) = objAcc.AccountHead
                vsGridDemand.TextMatrix(mRowCnt, 2) = gbFinancialYearID & "-" & gbFinancialYearID + 1
                vsGridDemand.TextMatrix(mRowCnt, 8) = gbCurrentPeriodID
                mPenalPeriodID = vsGridDemand.TextMatrix(mRowCnt, 8)
                If mPenalPeriodID = 2 Then
                    vsGridDemand.Cell(flexcpText, mRowCnt, 3) = "IInd Half"
                ElseIf mPenalPeriodID = 1 Then
                    vsGridDemand.Cell(flexcpText, mRowCnt, 3) = "Ist Half"
                End If
                vsGridDemand.TextMatrix(mRowCnt, 4) = ""
                vsGridDemand.TextMatrix(mRowCnt, 5) = val(txtFine.Text)
                vsGridDemand.TextMatrix(mRowCnt, 6) = objAcc.AccountHeadID
                vsGridDemand.TextMatrix(mRowCnt, 7) = gbFinancialYearID
                vsGridDemand.TextMatrix(mRowCnt, 9) = ""
                vsGridDemand.TextMatrix(mRowCnt, 10) = ""
                vsGridDemand.TextMatrix(mRowCnt, 11) = val(txtFine.Text)
                vsGridDemand.TextMatrix(mRowCnt, 12) = vbChecked
                vsGridDemand.TextMatrix(mRowCnt, 13) = ""
                vsGridDemand.TextMatrix(mRowCnt, 14) = ""
                vsGridDemand.TextMatrix(mRowCnt, 15) = ""
                vsGridDemand.TextMatrix(mRowCnt, 16) = ""
            End If
            If Rec.State Then
                Rec.Close
            End If
            Call Calculate
            chkSelectAll.value = vbChecked
        End If
    End Sub
  Private Sub FillInputData()
        txtInstName.Tag = IIf(IsNull(arrInput(0)), -1, arrInput(0))
        Call gSubSetComboItem2(cmbZone, val(arrInput(1)))
        txtWard.Text = val(IIf(IsNull(arrInput(3)), "", arrInput(3)))
        txtDoorNo1.Text = IIf(IsNull(arrInput(4)), "", (arrInput(4)))
        txtDoorNo2.Text = IIf(IsNull(arrInput(5)), "", (arrInput(5)))
        txtInstName.Text = IIf(IsNull(arrInput(6)), "", (arrInput(6)))
        txtHouseName.Text = IIf(IsNull(arrInput(8)), "", (arrInput(8)))
        txtStreet.Text = IIf(IsNull(arrInput(9)), "", (arrInput(9)))
        txtLocalPlace.Text = IIf(IsNull(arrInput(10)), "", (arrInput(10)))
        txtMainPlace.Text = IIf(IsNull(arrInput(11)), "", (arrInput(11)))
        txtOwnersName.Text = IIf(IsNull(arrInput(7)), "", (arrInput(7)))
    End Sub

    Private Sub vsGridDemand_Click()
        If vsGridDemand.Col = 12 Then
                vsGridDemand.Editable = flexEDKbdMouse
                If vsGridDemand.TextMatrix(vsGridDemand.Row, 0) <> "" Or vsGridDemand.TextMatrix(vsGridDemand.Row, 0) <> gbAcHeadCodePenalInterest Then
                    Call calculateFine
                    Call Calculate
                End If
        Else
            vsGridDemand.Editable = flexEDNone
        End If
    End Sub
    Private Sub vsGridDemand_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        Dim mLoop As Long
        Dim mRowCount As Integer
        Dim mYearID As Integer
        Dim mPeriodID   As Integer
        Dim mDemand     As String
        If Row > 0 Then
            If vsGridDemand.Cell(flexcpChecked, Row, Col) = 2 Then
             For mLoop = 1 To vsGridDemand.Rows - 1
                If vsGridDemand.TextMatrix(Row, 10) = vsGridDemand.TextMatrix(mLoop, 10) Then
                    If Row - 1 <> 0 Then
                        If vsGridDemand.Cell(flexcpChecked, Row - 1, 12) = 1 Then
                            vsGridDemand.Cell(flexcpChecked, mLoop, 12) = 1
                            mNumberOfSelections = mNumberOfSelections + 1
                        Else
                            Cancel = True
                        End If
                    Else
                        vsGridDemand.Cell(flexcpChecked, mLoop, 12) = 1
                        mNumberOfSelections = mNumberOfSelections + 1
                    End If
                End If
                If vsGridDemand.TextMatrix(mLoop, 0) = gbAcHeadCodePenalInterest Then
                    vsGridDemand.Cell(flexcpChecked, mLoop, 12) = 1
                End If
            Next mLoop
            Else ' Already  Checked
                 If vsGridDemand.Cell(flexcpChecked, Row - 1, Col) = 1 Then
                    For mLoop = 1 To vsGridDemand.Rows - 1
                        If vsGridDemand.TextMatrix(Row, 10) = vsGridDemand.TextMatrix(mLoop, 10) And vsGridDemand.TextMatrix(Row, 0) <> gbAcHeadCodePenalInterest Then
                            vsGridDemand.Cell(flexcpChecked, mLoop, 12) = 2
                            mNumberOfSelections = mNumberOfSelections - 1
                        End If
                    Next mLoop
                    For mLoop = Row To vsGridDemand.Rows - 1
                        If vsGridDemand.TextMatrix(mLoop, 0) <> gbAcHeadCodePenalInterest Then
                            vsGridDemand.Cell(flexcpChecked, mLoop, 12) = 2
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
        Call Calculate
        Call calculateFine
    End Sub

    Private Sub vsGridLeft_DblClick()
        Dim mInstID As Variant
        
        If vsGridLeft.Row > 0 Then
            mInstID = vsGridLeft.TextMatrix(vsGridLeft.Row, 5) ' Hidden Field Institution ID
            If val(mInstID) > 0 Then
                Call formInitialise
                Call DisplayInstitutionDetails(mInstID)
                'Note:- If Profession Tax - Traders Fetch Demand
                If mProfTaxInstTypeMode = 1 Then
                    fraInstitutionMaster.Visible = False
                    vsGridDemand.Visible = True
                    fraDemand.Visible = True
                    Call FillDemandDetailsGrid
                    Call calculateFine
                End If
            End If
       End If
    End Sub
                    '----------------------------------------'
                    '       To calculate Total amount        '
                    '----------------------------------------'
    Private Sub Calculate()
       Dim mTotal       As Double
       Dim mArrearAmt         As Double
       Dim mCurrentAmt         As Double
       Dim mCount       As Integer
       
       For mCount = 1 To vsGridDemand.Rows - 1
        If vsGridDemand.Cell(flexcpChecked, mCount, 12) = 1 Then
            If val(vsGridDemand.TextMatrix(mCount, 4)) <> 0 Then
                mArrearAmt = mArrearAmt + Format(val(vsGridDemand.TextMatrix(mCount, 4)), "0.00")
                txtArrearTotal.Text = Format(mArrearAmt, "0.00")
            ElseIf val(vsGridDemand.TextMatrix(mCount, 5)) <> 0 Then
                mCurrentAmt = mCurrentAmt + Format(val(vsGridDemand.TextMatrix(mCount, 5)), "0.00")
                txtCurrentTotal.Text = Format(mCurrentAmt, "0.00")
            End If
        End If
       Next
       txtDemandTotal.Text = Format(val(txtArrearTotal.Text) + val(txtCurrentTotal.Text), "0.00")
    End Sub
    
    Private Function FindNoofMonths(mYearID As Integer, mPeriodID As Integer, Optional dtUptoDate As Date) As Integer
        Dim mDemandDate As Variant
        Dim mNoOfMonths As Integer
        Dim mAmount     As Double
'        If vsGridDemand.Rows = mRowCnt Then vsGridDemand.Rows = vsGridDemand.Rows + 15
            'KERALA PANCHAYAT RAJ NIYAMAGALUM CHATTANGALUM - ANUBANDHA NIYAMAGAL 8th Edition
             'PAGE NO : 1000
            If mPeriodID = 1 Then
                    'mDemandDate = DateSerial(mYearID, 5, 1)
                    mDemandDate = DateSerial(mYearID, 9, 1)
            Else
                    'mDemandDate = DateSerial(mYearID + 1, 11, 1)
                    'mDemandDate = DateSerial(mYearID, 11, 1)
                    mDemandDate = DateSerial(mYearID + 1, 3, 1)
            End If
                
            If mYearID = gbFinancialYearID And mPeriodID = 2 Then
    '                Fine = 0
                mDemandDate = gbTransactionDate
                Exit Function
            End If
            If mDemandDate <= gbTransactionDate Then 'dtUptoDate Then
                mNoOfMonths = DateDiff("M", mDemandDate, gbTransactionDate)
            End If
'            If mYearID < gbFinancialYearID And mPeriodID = 1 Then
'                mNoOfMonths = Abs(DateDiff("M", mDemandDate, gbTransactionDate)) * 2
'                mNoOfMonths = mNoOfMonths + 1
'            ElseIf mYearID < gbFinancialYearID And mPeriodID = 2 Then
'                mNoOfMonths = DateDiff("M", mDemandDate, gbTransactionDate)
'            End If
            FindNoofMonths = mNoOfMonths
    End Function
    Private Sub calculateFine()
        Dim mFine, mProfTaxAmount As Double
        Dim mCnt As Variant
        Dim mCount As Integer
        For mCnt = 1 To vsGridDemand.Rows - 1 Step 1
            If vsGridDemand.Cell(flexcpChecked, mCnt, 12) = vbChecked Then
                If vsGridDemand.TextMatrix(mCnt, 0) = gbAcHeadCodeProfTaxTradersCurrent Or vsGridDemand.TextMatrix(mCnt, 0) = gbAcHeadCodeProfTaxTradersArrears And (vsGridDemand.TextMatrix(mCnt, 0) <> gbAcHeadCodePenalInterest) Then ' poornima
                    mProfTaxAmount = val(vsGridDemand.TextMatrix(mCnt, 11))
                    If Not (val(vsGridDemand.TextMatrix(mCnt, 7)) = gbFinancialYearID And val(vsGridDemand.TextMatrix(mCnt, 8)) = gbCurrentPeriodID) Then
                        mFine = mFine + (mProfTaxAmount * FindNoofMonths(vsGridDemand.TextMatrix(mCnt, 7), vsGridDemand.TextMatrix(mCnt, 8))) / 100
                    End If                                                          ' poornima
                End If
            End If
        Next
        txtFine.Text = Format(mFine, "0.00")
'        Call FillDemandDetailsGrid
        Dim mPenalPeriodID As Integer
        Dim mRowCnt As Integer
        Dim mPenal   As Double
        If val(txtFine.Text) > 0 Then
            mPenal = val(txtFine.Text)
        Else
            mPenal = 0
        End If
            For mCnt = 1 To vsGridDemand.Rows - 1
                If vsGridDemand.TextMatrix(mCnt, 0) = gbAcHeadCodePenalInterest Then
                    vsGridDemand.TextMatrix(mCnt, 5) = mPenal
                    vsGridDemand.TextMatrix(mCnt, 11) = mPenal
                    
                    vsGridDemand.TextMatrix(mCnt, 8) = 3
                    vsGridDemand.Cell(flexcpText, mRowCnt, 2) = CStr(gbFinancialYearID) & "-" & CStr(gbFinancialYearID + 1)
                    vsGridDemand.Cell(flexcpText, mRowCnt, 3) = "Full Year"
                End If
            Next
    End Sub



