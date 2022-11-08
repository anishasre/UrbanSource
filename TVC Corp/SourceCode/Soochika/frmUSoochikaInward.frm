VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUSoochikaInward 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Inward"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15210
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   945
      Top             =   7035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   7860
      TabIndex        =   70
      Top             =   30
      Width           =   7335
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   2760
         TabIndex        =   145
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   600
         TabIndex        =   144
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdgo 
         Caption         =   "GO"
         Height          =   375
         Left            =   6480
         TabIndex        =   137
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtyear 
         Height          =   360
         Left            =   5520
         TabIndex        =   136
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtInwNo 
         Height          =   360
         Left            =   3600
         TabIndex        =   134
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkaddr 
         BackColor       =   &H80000005&
         Caption         =   "Previous Address"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   720
         TabIndex        =   133
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtNoofPages 
         Height          =   345
         Left            =   480
         TabIndex        =   132
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtAtt 
         Height          =   630
         Left            =   1380
         TabIndex        =   130
         Top             =   6960
         Width           =   4200
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   7200
         Top             =   7080
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.ComboBox cmbSeatID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   5160
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ComboBox cmbSeat 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   5160
         Width           =   1425
      End
      Begin VB.ComboBox cmbDepartment 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   5130
         Width           =   3165
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "Cance&l"
         Height          =   405
         Left            =   5160
         TabIndex        =   63
         Top             =   6480
         Width           =   1395
      End
      Begin VB.CommandButton cmdReprint 
         Appearance      =   0  'Flat
         Caption         =   "Re&Print"
         Height          =   405
         Left            =   2040
         TabIndex        =   62
         Top             =   6480
         Width           =   1395
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   405
         Left            =   3600
         TabIndex        =   0
         Top             =   6480
         Width           =   1395
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Height          =   405
         Left            =   480
         TabIndex        =   61
         Top             =   6480
         Width           =   1395
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4365
         Left            =   60
         TabIndex        =   33
         Top             =   600
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   7699
         _Version        =   393216
         Tabs            =   6
         Tab             =   2
         TabsPerRow      =   6
         TabHeight       =   917
         BackColor       =   -2147483634
         TabCaption(0)   =   "Check List"
         TabPicture(0)   =   "frmUSoochikaInward.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "grvCheckList"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Valuables"
         TabPicture(1)   =   "frmUSoochikaInward.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grvValuables"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Certificate Address"
         TabPicture(2)   =   "frmUSoochikaInward.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label23"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label24"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label25"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label26"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label27"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Label28"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Label29"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Label30"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Label31"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Label32"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "txtCertHouseName"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "cmbCertDistrict"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "cmbCertState"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "txtCertPostOffice"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "txtCertPincode"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "txtCertMainPlace"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "txtCertLocalPlace"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "txtCertDoorNo1"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "txtCertWardNo"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "txtCertName"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "cmbCertGender"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "txtCertDoorNo2"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "chkAsInward"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).ControlCount=   23
         TabCaption(3)   =   "Reg Post Details"
         TabPicture(3)   =   "frmUSoochikaInward.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtRegDesg"
         Tab(3).Control(1)=   "txtRegPostNo"
         Tab(3).Control(2)=   "txtRegToWhome"
         Tab(3).Control(3)=   "Label35"
         Tab(3).Control(4)=   "Label34"
         Tab(3).Control(5)=   "Label33"
         Tab(3).ControlCount=   6
         TabCaption(4)   =   "Other"
         TabPicture(4)   =   "frmUSoochikaInward.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmbWardMember"
         Tab(4).Control(1)=   "cmbWardName"
         Tab(4).Control(2)=   "cmbBillReceiptType"
         Tab(4).Control(3)=   "txtBillReceiptDescription"
         Tab(4).Control(4)=   "txtBillReceiptAmount"
         Tab(4).Control(5)=   "txtBillReceiptNo"
         Tab(4).Control(6)=   "Label41"
         Tab(4).Control(7)=   "Label40"
         Tab(4).Control(8)=   "Label39"
         Tab(4).Control(9)=   "Label38"
         Tab(4).Control(10)=   "Label37"
         Tab(4).Control(11)=   "Label36"
         Tab(4).ControlCount=   12
         TabCaption(5)   =   "Reference Details"
         TabPicture(5)   =   "frmUSoochikaInward.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "grvRefDetails"
         Tab(5).ControlCount=   1
         Begin VB.CheckBox chkAsInward 
            Caption         =   "As Inward"
            Height          =   465
            Left            =   90
            TabIndex        =   128
            Top             =   585
            Width           =   1185
         End
         Begin VB.ComboBox cmbWardMember 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   -70590
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   3270
            Width           =   2505
         End
         Begin VB.ComboBox cmbWardName 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   -74310
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   3300
            Width           =   2505
         End
         Begin VB.ComboBox cmbBillReceiptType 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   -72960
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   780
            Width           =   2505
         End
         Begin VB.TextBox txtBillReceiptDescription 
            Height          =   360
            Left            =   -72960
            TabIndex        =   54
            Top             =   2310
            Width           =   2505
         End
         Begin VB.TextBox txtBillReceiptAmount 
            Height          =   360
            Left            =   -72960
            TabIndex        =   53
            Top             =   1800
            Width           =   2505
         End
         Begin VB.TextBox txtBillReceiptNo 
            Height          =   360
            Left            =   -72960
            TabIndex        =   52
            Top             =   1290
            Width           =   2505
         End
         Begin VB.TextBox txtRegDesg 
            Height          =   360
            Left            =   -73350
            TabIndex        =   49
            Top             =   1740
            Width           =   2295
         End
         Begin VB.TextBox txtRegPostNo 
            Height          =   360
            Left            =   -73350
            TabIndex        =   50
            Top             =   2430
            Width           =   2295
         End
         Begin VB.TextBox txtRegToWhome 
            Height          =   360
            Left            =   -73350
            TabIndex        =   48
            Top             =   1140
            Width           =   2295
         End
         Begin VB.TextBox txtCertDoorNo2 
            Height          =   375
            Left            =   5580
            TabIndex        =   41
            Top             =   1860
            Width           =   1245
         End
         Begin VB.ComboBox cmbCertGender 
            Height          =   360
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   960
            Width           =   1155
         End
         Begin VB.TextBox txtCertName 
            Height          =   405
            Left            =   2310
            TabIndex        =   37
            Top             =   930
            Width           =   4545
         End
         Begin VB.TextBox txtCertWardNo 
            Height          =   375
            Left            =   1110
            TabIndex        =   39
            Top             =   1860
            Width           =   2445
         End
         Begin VB.TextBox txtCertDoorNo1 
            Height          =   375
            Left            =   4620
            TabIndex        =   40
            Top             =   1860
            Width           =   915
         End
         Begin VB.TextBox txtCertLocalPlace 
            Height          =   375
            Left            =   1110
            TabIndex        =   42
            Top             =   2310
            Width           =   2445
         End
         Begin VB.TextBox txtCertMainPlace 
            Height          =   375
            Left            =   4620
            TabIndex        =   43
            Top             =   2310
            Width           =   2235
         End
         Begin VB.TextBox txtCertPincode 
            Height          =   375
            Left            =   1110
            MaxLength       =   6
            TabIndex        =   44
            Top             =   2760
            Width           =   2445
         End
         Begin VB.TextBox txtCertPostOffice 
            Height          =   375
            Left            =   4620
            TabIndex        =   45
            Top             =   2760
            Width           =   2235
         End
         Begin VB.ComboBox cmbCertState 
            Height          =   360
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   3210
            Width           =   2445
         End
         Begin VB.ComboBox cmbCertDistrict 
            Height          =   360
            Left            =   4620
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   3210
            Width           =   2235
         End
         Begin VB.TextBox txtCertHouseName 
            Height          =   375
            Left            =   1110
            TabIndex        =   38
            Top             =   1410
            Width           =   5745
         End
         Begin VSFlex8LCtl.VSFlexGrid grvValuables 
            Height          =   3585
            Left            =   -74910
            TabIndex        =   35
            Top             =   690
            Width           =   6975
            _cx             =   12303
            _cy             =   6324
            Appearance      =   2
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
            Rows            =   5
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmUSoochikaInward.frx":00A8
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
            ComboSearch     =   0
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
         Begin VSFlex8LCtl.VSFlexGrid grvCheckList 
            Height          =   3585
            Left            =   -74880
            TabIndex        =   34
            Top             =   690
            Width           =   7215
            _cx             =   12726
            _cy             =   6324
            Appearance      =   0
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
            Rows            =   10
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmUSoochikaInward.frx":0184
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
         Begin VSFlex8LCtl.VSFlexGrid grvRefDetails 
            Height          =   3525
            Left            =   -74910
            TabIndex        =   57
            Top             =   750
            Width           =   6975
            _cx             =   12303
            _cy             =   6218
            Appearance      =   2
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmUSoochikaInward.frx":0214
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
            ComboSearch     =   0
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
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Ward Member"
            Height          =   285
            Left            =   -71820
            TabIndex        =   107
            Top             =   3330
            Width           =   1155
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "Ward Name"
            Height          =   285
            Left            =   -74880
            TabIndex        =   106
            Top             =   3360
            Width           =   465
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "Bill/Receipt Type"
            Height          =   285
            Left            =   -74760
            TabIndex        =   105
            Top             =   840
            Width           =   1545
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   " Description"
            Height          =   285
            Left            =   -74760
            TabIndex        =   104
            Top             =   2340
            Width           =   1545
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
            Height          =   285
            Left            =   -74760
            TabIndex        =   103
            Top             =   1830
            Width           =   1545
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Bill/Receipt No"
            Height          =   285
            Left            =   -74760
            TabIndex        =   102
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Postal Number"
            Height          =   255
            Left            =   -74790
            TabIndex        =   101
            Top             =   2490
            Width           =   1245
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Designation"
            Height          =   255
            Left            =   -74790
            TabIndex        =   100
            Top             =   1785
            Width           =   1245
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "To Whome"
            Height          =   255
            Left            =   -74790
            TabIndex        =   99
            Top             =   1185
            Width           =   1245
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   225
            Left            =   270
            TabIndex        =   81
            Top             =   1005
            Width           =   735
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Ward No"
            Height          =   225
            Left            =   300
            TabIndex        =   80
            Top             =   1935
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Door No"
            Height          =   225
            Left            =   3720
            TabIndex        =   79
            Top             =   1935
            Width           =   735
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Local Place"
            Height          =   225
            Left            =   90
            TabIndex        =   78
            Top             =   2385
            Width           =   945
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Main Place"
            Height          =   225
            Left            =   3600
            TabIndex        =   77
            Top             =   2385
            Width           =   855
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Pincode"
            Height          =   225
            Left            =   90
            TabIndex        =   76
            Top             =   2835
            Width           =   945
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Postoffice"
            Height          =   225
            Left            =   3600
            TabIndex        =   75
            Top             =   2835
            Width           =   855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "State"
            Height          =   225
            Left            =   360
            TabIndex        =   74
            Top             =   3285
            Width           =   585
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "District"
            Height          =   225
            Left            =   3750
            TabIndex        =   73
            Top             =   3285
            Width           =   705
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "House Name"
            Height          =   225
            Left            =   60
            TabIndex        =   72
            Top             =   1485
            Width           =   945
         End
      End
      Begin MSComCtl2.DTPicker dtpDeliveryDate 
         Height          =   360
         Left            =   1530
         TabIndex        =   60
         Top             =   5580
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   635
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   16515073
         CurrentDate     =   40544
      End
      Begin VB.Label lblusername 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5040
         TabIndex        =   142
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label lbluser 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   5040
         TabIndex        =   139
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label lblyear 
         BackColor       =   &H80000005&
         Caption         =   "Year"
         Height          =   375
         Left            =   5040
         TabIndex        =   138
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblinwNo 
         BackColor       =   &H80000005&
         Caption         =   "InwardNo"
         Height          =   255
         Left            =   2640
         TabIndex        =   135
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Attachments"
         Height          =   225
         Left            =   120
         TabIndex        =   129
         Top             =   7200
         Width           =   1155
      End
      Begin VB.Label lblInwardDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inward Date : 01/01/1990"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   4890
         TabIndex        =   127
         Top             =   5880
         Width           =   2190
      End
      Begin VB.Label lblMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " * "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   10
         Left            =   7110
         TabIndex        =   123
         Top             =   5190
         Width           =   195
      End
      Begin VB.Label lblMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " * "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   9
         Left            =   4350
         TabIndex        =   122
         Top             =   5190
         Width           =   195
      End
      Begin VB.Label lblLastInward 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Last Inward : 000000"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   4890
         TabIndex        =   112
         Top             =   6120
         Width           =   1785
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         Height          =   225
         Left            =   270
         TabIndex        =   111
         Top             =   5640
         Width           =   1155
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         Height          =   285
         Left            =   4590
         TabIndex        =   109
         Top             =   5175
         Width           =   345
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   285
         Left            =   150
         TabIndex        =   108
         Top             =   5175
         Width           =   1005
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Pages"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   71
         Top             =   300
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7755
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   7785
      Begin VB.ListBox lstSubject 
         Appearance      =   0  'Flat
         Height          =   990
         Left            =   1470
         TabIndex        =   5
         Top             =   1140
         Width           =   5745
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Informar Details"
         ForeColor       =   &H80000008&
         Height          =   5625
         Left            =   0
         TabIndex        =   82
         Top             =   1980
         Width           =   7695
         Begin VB.TextBox txtasssyear 
            Height          =   360
            Left            =   3480
            TabIndex        =   147
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtseatid 
            Height          =   285
            Left            =   5640
            TabIndex        =   141
            Top             =   5160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtuserid 
            Height          =   285
            Left            =   4440
            TabIndex        =   140
            Top             =   5160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkMalInw 
            BackColor       =   &H80000005&
            Caption         =   "Malayalam Inward"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   1920
            TabIndex        =   131
            Top             =   5160
            Width           =   1995
         End
         Begin VB.ComboBox cmbGender 
            Height          =   360
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1110
            Width           =   1155
         End
         Begin VB.TextBox txtApplicantName 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   2310
            TabIndex        =   16
            Top             =   1080
            Width           =   5115
         End
         Begin VB.TextBox txtWardNo 
            Height          =   330
            Left            =   1110
            TabIndex        =   18
            Top             =   2010
            Width           =   1125
         End
         Begin VB.TextBox txtDoorNo1 
            Height          =   495
            Left            =   5160
            TabIndex        =   19
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtLocalPlace 
            Height          =   300
            Left            =   1110
            TabIndex        =   22
            Top             =   2460
            Width           =   2445
         End
         Begin VB.TextBox txtMainPlace 
            Height          =   300
            Left            =   4980
            TabIndex        =   23
            Top             =   2460
            Width           =   2445
         End
         Begin VB.TextBox txtPincode 
            Height          =   270
            Left            =   1110
            MaxLength       =   6
            TabIndex        =   24
            Top             =   2910
            Width           =   2445
         End
         Begin VB.TextBox txtPostOffice 
            Height          =   270
            Left            =   4980
            TabIndex        =   25
            Top             =   2910
            Width           =   2445
         End
         Begin VB.ComboBox cmbState 
            Height          =   360
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   3360
            Width           =   2445
         End
         Begin VB.ComboBox cmbDistrict 
            Height          =   360
            Left            =   4980
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   3360
            Width           =   2445
         End
         Begin VB.TextBox txtContactNo 
            Height          =   375
            Left            =   1110
            MaxLength       =   11
            TabIndex        =   28
            Top             =   3780
            Width           =   2445
         End
         Begin VB.TextBox txtContactMail 
            Height          =   300
            Left            =   4980
            TabIndex        =   29
            Top             =   3780
            Width           =   2445
         End
         Begin VB.CheckBox chkBPL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "BPL"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   1260
            TabIndex        =   30
            Top             =   4560
            Width           =   1035
         End
         Begin VB.CheckBox chkSCST 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "SC/ST"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   2490
            TabIndex        =   31
            Top             =   4680
            Width           =   1035
         End
         Begin VB.TextBox txtDocProof 
            Height          =   615
            Left            =   4920
            TabIndex        =   32
            Top             =   4320
            Width           =   2445
         End
         Begin VB.CheckBox chkInstitution 
            BackColor       =   &H80000005&
            Caption         =   "Institution"
            Height          =   255
            Left            =   1320
            TabIndex        =   10
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox chkInsideLB 
            BackColor       =   &H80000005&
            Caption         =   "Inside LB"
            Height          =   285
            Left            =   3180
            TabIndex        =   11
            Top             =   270
            Width           =   1035
         End
         Begin VB.CheckBox chkCourtfee 
            BackColor       =   &H80000005&
            Caption         =   "Courtfee Stamp"
            Height          =   315
            Left            =   5250
            TabIndex        =   12
            Top             =   270
            Width           =   1695
         End
         Begin VB.TextBox txtInstitutionName 
            Height          =   375
            Left            =   1110
            TabIndex        =   13
            Top             =   630
            Width           =   2445
         End
         Begin VB.TextBox txtInstitutionDesg 
            Height          =   375
            Left            =   4980
            TabIndex        =   14
            Top             =   630
            Width           =   2445
         End
         Begin VB.TextBox txtHouseName 
            Height          =   360
            Left            =   1110
            TabIndex        =   17
            Top             =   1560
            Width           =   6315
         End
         Begin VB.TextBox txtDoorNo2 
            Height          =   450
            Left            =   6120
            TabIndex        =   20
            Top             =   1920
            Width           =   795
         End
         Begin VB.CommandButton cmdTaxSearch 
            Caption         =   "Tax"
            Enabled         =   0   'False
            Height          =   465
            Left            =   7080
            TabIndex        =   21
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label46 
            BackColor       =   &H80000005&
            Caption         =   "AssessYear"
            Height          =   375
            Left            =   2400
            TabIndex        =   146
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   795
            Index           =   13
            Left            =   7410
            TabIndex        =   126
            Top             =   4320
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   12
            Left            =   7410
            TabIndex        =   125
            Top             =   720
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   11
            Left            =   3540
            TabIndex        =   124
            Top             =   690
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   795
            Index           =   8
            Left            =   7380
            TabIndex        =   121
            Top             =   3420
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   7
            Left            =   3510
            TabIndex        =   120
            Top             =   3420
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   795
            Index           =   6
            Left            =   7410
            TabIndex        =   119
            Top             =   2580
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   5
            Left            =   3570
            TabIndex        =   118
            Top             =   2100
            Width           =   195
         End
         Begin VB.Label lblMandatory 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " * "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   795
            Index           =   4
            Left            =   7440
            TabIndex        =   117
            Top             =   1170
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00000000&
            Height          =   825
            Left            =   270
            TabIndex        =   98
            Top             =   1155
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ward No"
            Height          =   825
            Left            =   300
            TabIndex        =   97
            Top             =   2085
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Door No"
            Height          =   225
            Left            =   4320
            TabIndex        =   96
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Local Place"
            Height          =   825
            Left            =   90
            TabIndex        =   95
            Top             =   2535
            Width           =   945
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Main Place"
            Height          =   825
            Left            =   3960
            TabIndex        =   94
            Top             =   2535
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pincode"
            Height          =   825
            Left            =   90
            TabIndex        =   93
            Top             =   2985
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Postoffice"
            Height          =   825
            Left            =   3960
            TabIndex        =   92
            Top             =   2985
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   825
            Left            =   360
            TabIndex        =   91
            Top             =   3435
            Width           =   585
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "District"
            Height          =   825
            Left            =   3720
            TabIndex        =   90
            Top             =   3435
            Width           =   1095
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No"
            Height          =   945
            Left            =   120
            TabIndex        =   89
            Top             =   3840
            Width           =   945
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Mail"
            Height          =   825
            Left            =   3780
            TabIndex        =   88
            Top             =   3855
            Width           =   1035
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            Height          =   825
            Left            =   90
            TabIndex        =   87
            Top             =   4305
            Width           =   945
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Doc Prof"
            Height          =   345
            Left            =   3960
            TabIndex        =   86
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Institution Name"
            Height          =   525
            Left            =   210
            TabIndex        =   85
            Top             =   555
            Width           =   795
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Desgination"
            Height          =   255
            Left            =   3780
            TabIndex        =   84
            Top             =   690
            Width           =   1005
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "House Name"
            Height          =   825
            Left            =   60
            TabIndex        =   83
            Top             =   1635
            Width           =   945
         End
      End
      Begin VB.TextBox txtSubID 
         Height          =   345
         Left            =   840
         MaxLength       =   3
         TabIndex        =   3
         Top             =   750
         Width           =   615
      End
      Begin VB.CheckBox chkByMember 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "By Member"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4230
         TabIndex        =   7
         Top             =   1170
         Width           =   1485
      End
      Begin VB.CheckBox chkByRef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "By  Ref"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2190
         TabIndex        =   6
         Top             =   1170
         Width           =   1485
      End
      Begin VB.TextBox txtRefNo 
         Height          =   360
         Left            =   1260
         TabIndex        =   8
         Top             =   1500
         Width           =   1695
      End
      Begin VB.TextBox txtSubject 
         Height          =   465
         Left            =   1470
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   690
         Width           =   5745
      End
      Begin VB.ComboBox cmbPriority 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1965
      End
      Begin VB.ComboBox cmbCorrespondance 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpRefDate 
         Height          =   360
         Left            =   5130
         TabIndex        =   9
         Top             =   1500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   40544
      End
      Begin VB.Label lblsubjectmaster 
         BackColor       =   &H80000005&
         Caption         =   "SubjectMaster"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " * "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   7440
         TabIndex        =   116
         Top             =   3180
         Width           =   195
      End
      Begin VB.Label lblMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " * "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   7230
         TabIndex        =   115
         Top             =   810
         Width           =   195
      End
      Begin VB.Label lblMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " * "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   7230
         TabIndex        =   114
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblMandatory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " * "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   113
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Date"
         Height          =   225
         Left            =   3270
         TabIndex        =   69
         Top             =   1568
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No"
         Height          =   225
         Left            =   180
         TabIndex        =   68
         Top             =   1568
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   225
         Left            =   180
         TabIndex        =   67
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Priority"
         Height          =   285
         Left            =   4020
         TabIndex        =   66
         Top             =   248
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Correspondance"
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   255
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmUSoochikaInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objdb As New clsDB
Dim mSql As Variant
Dim i As Integer
Dim J As Integer
Dim SoochikaFileID  As Variant
Dim DistributionID  As Integer
Dim FunctionID  As Integer
Dim ReferenceID As Integer

'paperless
Dim fso As New FileSystemObject
Dim fld As Folder
Public Sub DisableControls()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = False
            'ctl.Locked = True
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is CheckBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is Buttons Then
            ctl.Enabled = False
        End If
        cmdSave.Enabled = False
        cmdReprint.Enabled = False
    Next ctl
End Sub
Private Sub EnableControls()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is CheckBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is Buttons Then
            ctl.Enabled = True
        End If
        cmdSave.Enabled = True
        cmdReprint.Enabled = True
    Next ctl
End Sub
Private Sub InwardAddress()
            cmbCertGender.ListIndex = cmbGender.ListIndex
            txtCertName.Text = txtApplicantName.Text
            txtCertHouseName.Text = txtHouseName.Text
            txtCertWardNo.Text = txtWardNo.Text
            txtCertDoorNo1.Text = txtDoorNo1.Text
            txtCertDoorNo2.Text = txtDoorNo2.Text
            txtCertLocalPlace.Text = txtLocalPlace.Text
            txtCertMainPlace.Text = txtMainPlace.Text
            txtCertPincode.Text = txtPincode.Text
            txtCertPostOffice.Text = txtPostOffice.Text
            cmbCertDistrict.ListIndex = cmbDistrict.ListIndex 'gbDistID//changed by soumya vs 13jan15
            cmbCertState.ListIndex = cmbState.ListIndex
End Sub



Private Function CheckVersion()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    
    Rec.Open "SpCheckVersion '" & gbSoochikaDBVer & "','" & gbSoochikaScriptVer & "'", mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        If (Rec.Fields(0) = "1") Then
            CheckVersion = True
        Else
            CheckVersion = False
        End If
    Else
        CheckVersion = False
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Function CheckInwardDate()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    
    CheckInwardDate = False
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    
    Rec.Open "select max(dtDateofreceipt) as [MaxDate] from tInwardDetails", mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        If (DateDiff("d", CDate(Rec!MaxDate), Date) < 0) Then
            CheckInwardDate = False
        Else
            CheckInwardDate = True
        End If
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Function CheckInwardUser(ByVal UserID As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim arrOut As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    mSql = "select numUserID from tUserDetails where numUserID=" & UserID & " and((intUserTypeID=6) or (intUserTypeID=5 and flgClerical=1))"
    Set Rec = mCnn.Execute(mSql)
    If Not (Rec.EOF Or Rec.BOF) Then
        CheckInwardUser = True
    Else
        CheckInwardUser = False
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Sub IntegrationSuites()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim arrOut As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    mSql = "select * from tSuiteInstallDetails"
    Set Rec = mCnn.Execute(mSql)
    If Not (Rec.EOF Or Rec.BOF) Then
        gbSevanaIntegration = IIf(IsNull(Rec!flgSevana), 0, Rec!flgSevana)
        gbSanchayaIntegration = IIf(IsNull(Rec!flgSanchaya), 0, Rec!flgSanchaya)
        gbSaankhyaIntegration = IIf(IsNull(Rec!flgSaankhyaDouble), 0, Rec!flgSaankhyaDouble)
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Function SaveUserLog()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim arrOut As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    ReDim arrIn(7)
    ReDim arrOut(0)
    arrIn(0) = 0
    arrIn(1) = gbnumUserId
    arrIn(2) = gbnumSeatID
    arrIn(3) = Null
    arrIn(4) = GetMacAddress
    arrIn(5) = GetIPAddress
    arrIn(6) = "Inward Module from saankhya"
    arrIn(7) = "Saankhya"
    objdb.ExecuteSP "SpSaveUserlog", arrIn, arrOut, , mCnn, adCmdStoredProc
    SaveUserLog = arrOut(0, 0)
End Function
Private Function GetlastInward()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    
    Rec.Open "select  right('000000'+cast(isnull(max(numCurrentNo),0) as varchar),6) as [MaxInwardNo] from tInwardDetails where year(dtDateofReceipt)=year(getdate())", mCnn
    GetlastInward = Rec!Maxinwardno
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function

Private Sub FillCombos()
    Call PopulateList(cmbCorrespondance, "SP_SelectCorrespondance 1,1", , False, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbPriority, "Sp_SelectPriority 1", , False, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbGender, "Sp_selectGenderCode 1", , False, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbCertGender, "Sp_selectGenderCode 1", , False, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbState, "Sp_SelectState 1", , False, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbCertState, "Sp_SelectState 1", , False, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbCertDistrict, "Sp_SelectDistrict 1,32", , True, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbDepartment, "SP_SelectDepartment 1", , True, True, True, enuSourceString.SoochikaUnicode)
    Call FillFlexGridCombo(grvValuables, 1, "Sp_SelectInstrumentType 1", adCmdText, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbBillReceiptType, "Sp_SelectBillReceiptType 1", , True, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbWardName, "SpSelectWard 1", , True, True, True, enuSourceString.SoochikaUnicode)
    Call PopulateList(cmbWardMember, "SpSelectWardMember 1", , True, True, True, enuSourceString.SoochikaUnicode)
End Sub
Public Sub Clear()
    lblInwardDate.Caption = "Inward Date : " & Format(Date, "DD/MM/YYYY")
    Form_Activate
    cmbCorrespondance.ListIndex = 0
    chkByMember.Value = 0
    chkByRef.Value = 0
    chkByMember.Enabled = False
    chkByRef.Enabled = False
    chkInstitution.Value = 0
    cmbGender.ListIndex = 0
    txtRefNo.Text = ""
    dtpRefDate.Value = Date
    dtpRefDate.Enabled = False
    dtpDeliveryDate.Value = Date
    'dtpDeliveryDate.value = Null
    CheckInstitution (0)
    txtInstitutionDesg.Text = ""
    txtInstitutionName.Text = ""
    InsideLB (1)
    chkCourtfee.Value = 0
    lstSubject.Visible = False
    txtSubID.Text = ""
    txtSubject.Text = ""
    txtApplicantName.Text = ""
    txtHouseName.Text = ""
    txtWardNo.Text = ""
    txtDoorNo1.Text = ""
    txtDoorNo2.Text = ""
    txtMainPlace.Text = ""
    txtLocalPlace.Text = ""
    txtPincode.Text = ""
    txtPostOffice.Text = ""
    txtContactMail.Text = ""
    txtContactNo.Text = ""
    chkBPL.Value = 0
    chkSCST.Value = 0
    txtDocProof.Text = ""
    txtNoofPages.Text = ""
    grvCheckList.Rows = 2
    grvCheckList.Clear 1
    grvValuables.Rows = 2
    grvValuables.Clear 1
    cmbCertGender.ListIndex = 0
    txtCertName.Text = ""
    txtCertHouseName.Text = ""
    txtCertWardNo.Text = ""
    txtCertDoorNo1.Text = ""
    txtCertDoorNo2.Text = ""
    txtCertLocalPlace.Text = ""
    txtCertMainPlace.Text = ""
    txtCertPincode.Text = ""
    txtCertPostOffice.Text = ""
    cmbBillReceiptType.ListIndex = 0
    txtBillReceiptAmount.Text = ""
    txtBillReceiptDescription.Text = ""
    txtBillReceiptNo.Text = ""
    txtRegDesg.Text = ""
    txtRegPostNo.Text = ""
    txtRegToWhome.Text = ""
    cmbWardName.ListIndex = 0
    cmbWardMember.ListIndex = 0
    grvRefDetails.Rows = 2
    grvRefDetails.Clear 1
    cmbDepartment.ListIndex = 0
    cmbSeat.Clear
    cmbSeatID.Clear
    chkAsInward.Value = 0
    chkaddr.Value = 0
    'CHANGED
    lblusername.Caption = ""
    txtInwNo.Text = ""
    txtYear.Text = ""
    lblLastInward.Caption = " Last Inward : " & GetlastInward & " "
    InwardMode = 0
    If (gbSanchayaIntegration = 0) Then
        cmdTaxSearch.Enabled = False
    Else
        cmdTaxSearch.Enabled = True
    End If
    txtAtt.Text = ""
End Sub
Private Sub CheckInstitution(Check As Variant)
    If (Check = 1) Then
        txtInstitutionDesg.Enabled = True
        txtInstitutionName.Enabled = True
        lblMandatory(11).Visible = True
        lblMandatory(12).Visible = True
    Else
        txtInstitutionDesg.Enabled = False
        txtInstitutionName.Enabled = False
        lblMandatory(11).Visible = False
        lblMandatory(12).Visible = False
    End If
End Sub
Private Sub InsideLB(inside As Variant)
    If (inside = 1) Then
        cmbState.ListIndex = 31
        'changed by soumya V s
        cmbCertState.ListIndex = 31
    Call PopulateList(cmbCertDistrict, "Sp_SelectDistrict 1," & cmbCertState.ItemData(cmbCertState.ListIndex), , False, True, True, enuSourceString.SoochikaUnicode)
        For i = 0 To cmbCertDistrict.ListCount - 1
            If (cmbCertDistrict.ItemData(i) = gbDistID) Then
                cmbCertDistrict.ListIndex = i
                'cmbCertDistrict.ListIndex = i
            End If
        Next i
        Call PopulateList(cmbDistrict, "Sp_SelectDistrict 1," & cmbState.ItemData(cmbState.ListIndex), , False, True, True, enuSourceString.SoochikaUnicode)
        For i = 0 To cmbDistrict.ListCount - 1
            If (cmbDistrict.ItemData(i) = gbDistID) Then
                cmbDistrict.ListIndex = i
                'cmbCertDistrict.ListIndex = i
            End If
        Next i
        cmbState.Enabled = False
        cmbDistrict.Enabled = False
        'chkInsideLB.value = 1
        lblMandatory(5).Visible = True
    Else
        cmbState.Enabled = True
        cmbDistrict.Enabled = True
        lblMandatory(5).Visible = False
    End If
End Sub
Private Sub GetSubjectSeat()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
   
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If

    ReDim arrIn(1)
    arrIn(0) = txtSubID.Text
    arrIn(1) = txtWardNo.Text
    Set Rec = objdb.ExecuteSP("SpSelectSubjectSeatCoding", arrIn, , False, mCnn, adCmdStoredProc)
    'changed by soumya V S
    
     If Not (Rec.EOF Or Rec.BOF) Then
   
     If (Rec!numSubTypeID <> Null) Then
     Label42.Visible = False
     Label43.Visible = False
    
     
   
    
     Else
     'CHANGED
    lblusername.Caption = ""
        For i = 0 To cmbDepartment.ListCount - 1
            If (cmbDepartment.ItemData(i) = Rec!intDeptID) Then
                cmbDepartment.ListIndex = i
                'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & Rec!intDeptID, , True, True, True, enuSourceString.SoochikaUnicode)
                'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & Rec!intDeptID, , True, True, True, enuSourceString.SoochikaUnicode)
                
                
                'LATEST 24NOV
                Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                For J = 0 To cmbSeat.ListCount - 1
                    If (cmbSeatID.List(J) = Rec!numSeatID) Then
                    'chnaged by soumya V S
                        cmbSeatID.ListIndex = J
                        cmbSeat.ListIndex = J
                    End If
                Next
            End If
        Next
     End If

If (cmbSeat.ListIndex > 0) Then
ReDim arrIn(0)
arrIn(0) = cmbSeatID.Text
Set Rec = objdb.ExecuteSP("spSelectUser", arrIn, , , mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblusername.Caption = Rec!chvUserNameEng
Else
lblusername.Caption = ""

    Rec.Close
    End If
    End If
    
     Else
      If (cmbDepartment.ListIndex <> 0) Then
       
        lblusername.Caption = ""
        'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex), , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex), , True, True, True, enuSourceString.SoochikaUnicode)
        
        'add  by vipin 21-09-2012
        'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
         
         
        'LATEST 24NOV
                Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
    End If

   End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub getSubjectDeliverydate(ByVal SubID As Integer)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim Rec1 As New ADODB.Recordset
    Dim Rec2 As New ADODB.Recordset
    Dim arrIn As Variant
    Dim a As Integer
    Dim deliveryDate As Variant
    Dim strDate As Date
    
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    'changed by soumya V S on 14.05.2014
    'changed by soumya V S on 14.10.2014
    Rec.Open "select intPeriod,FlgCurrespondance from tSubjectDeliveryPeriod where numSubjectID=" & SubID, mCnn
    'Rec1.Open "Set dateformat DMY SELECT COUNT(*) as cnt FROM mholiday WHERE dtDate>getdate() and dtDate<=dATEADD(DD,15,GETDATE())", mCnn
    Rec2.Open "select numSubTypeID from tSubjectDeliveryPeriod where numSubTypeID<>0  and numSubTypeID IS NOT NULL and numSubjectID=" & SubID, mCnn
    
   
        If (cmbCorrespondance.ListIndex = 8) Then
    'added by soumya on  Sept2015 for RTI-30days
    '**************************************************
    strDate = Format(Now, "dd/MM/yyyy")
        deliveryDate = strDate
       
       ' If Not (Rec.EOF Or Rec.BOF) Then
        
        For i = 1 To 29
        deliveryDate = DateAdd("d", 1, deliveryDate)
       ' While (CheckHoliday(deliveryDate) = True)
        'deliveryDate = DateAdd("d", 1, deliveryDate)
       ' Wend
        Next i
        'End If
    If Not (Rec2.EOF Or Rec2.BOF) Then
    
    dtpDeliveryDate.Visible = False
    
    Label44.Visible = False
     
    
    
    Else
    dtpDeliveryDate.Visible = True
    Label44.Visible = True
    'CHNAGED on JUL
aa:    While (CheckHoliday(deliveryDate) = True)
    deliveryDate = DateAdd("d", 1, deliveryDate)
    GoTo aa
    Wend
    dtpDeliveryDate.Value = deliveryDate
    
   
    
    End If
  Else
    
    
   
        strDate = Format(Now, "dd/MM/yyyy")
        deliveryDate = strDate
       
        If Not (Rec.EOF Or Rec.BOF) Then
        
        For i = 1 To CInt(Rec!intPeriod)
        deliveryDate = DateAdd("d", 1, deliveryDate)
        While (CheckHoliday(deliveryDate) = True)
        deliveryDate = DateAdd("d", 1, deliveryDate)
        Wend
        Next i
        End If
    If Not (Rec2.EOF Or Rec2.BOF) Then
    
    dtpDeliveryDate.Visible = False
    
    Label44.Visible = False
     
    
    
    Else
    dtpDeliveryDate.Visible = True
    
    Label44.Visible = True
    dtpDeliveryDate.Value = deliveryDate
    
   
    
    End If
    End If
    
    If Not (Rec.EOF Or Rec.BOF) Then
    If (Rec!FlgCurrespondance) = 1 Then
   cmbCorrespondance.ListIndex = 9
    End If
    End If
    
    
    'changed by soumya VS on 14/10/2014
    
'    Rec.Open "select intPeriod from tSubjectDeliveryPeriod where numSubjectID=" & SubID, mCnn
'    If Not (Rec.EOF Or Rec.BOF) Then
'        deliveryDate = DateAdd("d", Rec!intPeriod, Date)
'        If (CheckHoliday(deliveryDate) = True) Then
'            MsgBox "Delivery date falls in Holiday ", vbInformation, "SOOCHIKA"
'        End If
'        dtpDeliveryDate.value = deliveryDate
'    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub GetRCReference(ByVal Subject As Variant)
    Dim flgRef As Variant
    
    flgRef = InStr(Subject, "Residential")
    If (flgRef <> 0) Then
        chkByMember.Enabled = True
        chkByRef.Enabled = True
    Else
        chkByMember.Enabled = False
        chkByRef.Enabled = False
    End If
End Sub
Private Function CheckSevanaInward(ByVal SubID As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    ''-------
    ''' Added On 28.09.12 By Vipin Kelembath
    If SubID = "" Then
        SubID = 0
    End If
    ''-------
    mSql = "select numSubjectSuiteID from mSubjectSuite where numSubjectID=" & SubID & " and intSuiteID=112"
    Rec.Open mSql, mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        CheckSevanaInward = IIf(IsNull(Rec!numSubjectSuiteID), 0, Rec!numSubjectSuiteID)
    Else
        CheckSevanaInward = 0
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Sub SaveSevanaStatus(mCnn As ADODB.Connection, InwNo As Variant, Year As Variant, Dt As Variant, mID As Variant, SID As Variant)
    Dim Sevarr As Variant
    ReDim Sevarr(4)
    Sevarr(0) = InwNo
    Sevarr(1) = Year
    Sevarr(2) = Dt
    Sevarr(3) = mID
    Sevarr(4) = SID
   objdb.ExecuteSP "sp_insertinwardstatusSoochika", Sevarr, , , mCnn, adCmdStoredProc
End Sub
Private Function CheckReceipt(ByVal SubID As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    
    CheckReceipt = 0
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connectoin Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    
    mSql = "select tnyReceipt from mSubjectReceipt where numSubjectID=" & SubID
    Set Rec = mCnn.Execute(mSql)
    If Not (Rec.EOF Or Rec.BOF) Then
        If (Rec!tnyReceipt = 1) Then
            CheckReceipt = 1
        End If
    End If
    Rec.Close
    
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function

Private Function CheckHoliday(ByVal dtDate As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    mSql = " select * from mHoliday where convert(datetime,dtdate,103)=convert(datetime,'" & dtDate & "',103)"
    Rec.Open mSql, mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        CheckHoliday = True
    Else
        CheckHoliday = False
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Sub GetPostoffice()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    
    Rec.Open "select chvPostofficeEng as [PostOffice] from mPostoffice where intPincode=" & txtPincode.Text, mCnn
    
    If Not (Rec.EOF Or Rec.BOF) Then
        txtPostOffice.Text = Rec!PostOffice
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub FillEnclosureGrid(ByVal SubID As Integer)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim arrOut As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    
    ReDim arrIn(1)
    arrIn(0) = 1
    arrIn(1) = SubID
    Set Rec = objdb.ExecuteSP("Sp_SelectSubjectEnclosure", arrIn, arrOut, , mCnn, adCmdStoredProc)
    If IsArray(arrOut) Then
        For i = 0 To UBound(arrOut, 2)
            If i > 0 Then
                   grvCheckList.Rows = grvCheckList.Rows + 1
            End If
            grvCheckList.TextMatrix(i + 1, 3) = arrOut(1, i)
            grvCheckList.TextMatrix(i + 1, 2) = arrOut(0, i)
        Next i
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Function ValidateData()
 Dim strDate As Date
    ValidateData = False
    If (txtSubject.Text = "") Then
        MsgBox "Please enter subject ", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtSubject.SetFocus
    ElseIf (chkInstitution.Value = 1 And txtInstitutionName.Text = "") Then
        MsgBox "Please enter Institution Name", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtInstitutionName.SetFocus
        
           'Contact No Validating
        'Changed by 06 march 2015 soumyavs
        
        ElseIf (txtContactNo.Text) <> "" And Len(txtContactNo.Text) < 10 Then
        MsgBox "Please enter a valid telephone number.", vbInformation, "SOOCHIKA"
         ValidateData = False
        txtContactNo.SetFocus
        'CHNAGED
     ElseIf (txtContactMail <> "") Then
    
        If InStr(1, txtContactMail, "@") = 0 Then
       MsgBox "Email Address is Invalid", vbInformation, "SOOCHIKA"
       ValidateData = False
       txtContactMail.SetFocus
       ElseIf InStr(1, txtContactMail, ".") = 0 Then
       MsgBox "Email Address is Invalid", vbInformation, "SOOCHIKA"
       ValidateData = False
       txtContactMail.SetFocus
       End If
       
       'Nov18
        'ElseIf (gbSevanaMainTypeID = 0) Then
        
        
        
        
        ElseIf (dtpDeliveryDate.Value > 1) Then
            strDate = Format(Now, "dd/MM/yyyy")
           If (dtpDeliveryDate.Value < strDate) Then
           MsgBox "Delivery date should be greater than today !!!", vbInformation, "SOOCHIKA"
           ValidateData = False
           dtpDeliveryDate.SetFocus
           
           
           ElseIf (CheckHoliday(dtpDeliveryDate.Value) = True) Then
           MsgBox "The selected delivery date is holiday !!!", vbInformation, "SOOCHIKA"
           ValidateData = False
           dtpDeliveryDate.SetFocus
           'Else
         
           'ValidateData = True
           'End If

     
    ElseIf (dtpDeliveryDate.Value = "") Then
         MsgBox "Please enter Service Delivery Date !!!", vbInformation, "SOOCHIKA"
           ValidateData = False
           dtpDeliveryDate.SetFocus
         

    ElseIf (chkInstitution.Value = 1 And txtInstitutionDesg.Text = "") Then
        MsgBox "Please enter Institution Designation", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtInstitutionDesg.SetFocus
    ElseIf (Trim(txtApplicantName.Text) = "") Then
        MsgBox "Please enter Applicant Name ", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtApplicantName.SetFocus
    ElseIf (chkInsideLB.Value = 1 And txtWardNo.Text = "") Then
        MsgBox "Please enter Ward No", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtWardNo.SetFocus
    ElseIf txtMainPlace.Text = "" Then
        MsgBox "Please enter Main place", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtMainPlace.SetFocus
    ElseIf (chkBPL.Value = 1 Or chkSCST.Value = 1) And txtDocProof.Text = "" Then '' Changed by poornima on 04/02/2012
        MsgBox "Please enter Doc Proof", vbInformation, "SOOCHIKA"
        ValidateData = False
        txtDocProof.SetFocus
        
    ElseIf (ValidateEnclosure() = False) Then
        MsgBox "Please select any Enclosures", vbInformation, "SOOCHIKA"
        ValidateData = False
        SSTab1.Tab = 0
    ElseIf (ValidateValuables() = False) Then
        MsgBox "Please fill the required data of valuables ", vbInformation, "SOOCHIKA"
        ValidateData = False
        SSTab1.Tab = 1
        'CHANGED
    'ElseIf (gbSevanaMainTypeID = 0) Then
        ElseIf cmbDepartment.ListIndex = 0 Then
        MsgBox "Please select the Department", vbInformation, "SOOCHIKA"
        ValidateData = False
        cmbDepartment.SetFocus
        'Else
        'ValidateData = True
        'End If
    'ElseIf (gbSevanaMainTypeID = 0) Then
        ElseIf cmbSeat.ListIndex = -1 Then
        MsgBox "Please select the Seat", vbInformation, "SOOCHIKA"
        ValidateData = False
        cmbSeat.SetFocus
        'Else
        'ValidateData = True
        'End If
        
        
      

    Else
        ValidateData = True
    End If
    Else
    ValidateData = True
      End If
      'End If
      'End If
      'End If
End Function
Private Function ValidateEnclosure()
    ValidateEnclosure = True
    If (val(txtNoofPages.Text) > 0) Then
        ValidateEnclosure = False
        For i = 1 To grvCheckList.Rows - 1
            If (grvCheckList.TextMatrix(i, 1) = "-1") Then
                ValidateEnclosure = True
                Exit Function
            End If
        Next i
    End If
End Function
Private Function ValidateValuables()
    ValidateValuables = True
    If (grvValuables.Rows > 2) Then
        For i = 1 To grvValuables.Rows - 2
            If ((grvValuables.TextMatrix(i, 1) <> "" And grvValuables.TextMatrix(i, 2) <> "" And grvValuables.TextMatrix(i, 3) <> "" And grvValuables.TextMatrix(i, 4) <> "" And grvValuables.TextMatrix(i, 5) <> "")) Or ((grvValuables.TextMatrix(i, 1) = "" And grvValuables.TextMatrix(i, 2) = "" And grvValuables.TextMatrix(i, 3) = "" And grvValuables.TextMatrix(i, 4) = "" And grvValuables.TextMatrix(i, 5) = "")) Then
                ValidateValuables = True
            Else
                ValidateValuables = False
                Exit Function
            End If
        Next
    End If
End Function
Public Function SaveSoochika(mCnn As ADODB.Connection)
    Dim Rec As New ADODB.Recordset
    Dim arrOut As Variant
    SoochikaFileID = SaveSoochikaInwardDetails(mCnn)
    SaveSMS '15-06-2012 SMS(Ranjitha)
    SaveSoochikaKeywordDetails mCnn
    SaveSoochikaInwardTrackDetails mCnn
    If txtCertName.Text = "" Then 'ADDED BY VIPIN ON 29-09-2012
        Call InwardAddress
    End If
    'If (val(txtNoofPages.Text) > 0) Then
        SaveSoochikaEnclosureDetails mCnn
    'End If
    If (cmbBillReceiptType.ListIndex > 0) Then
        SaveSoochikaBillReceiptDetails mCnn
    End If
    If (txtRegToWhome.Text <> "") Then
        SaveSoochikaRegisteredPostDetails mCnn
    End If
    If (txtCertName.Text <> "") Then
        SaveSoochikaCertificateDetails mCnn
    End If
    If (grvRefDetails.TextMatrix(1, 1) <> "") Then
        SaveSoochikaReferenceDetails mCnn
    End If
    If (grvValuables.TextMatrix(1, 1) <> "") Then
        SaveSoochikaValuableDetails mCnn
    End If
    SaveSoochika = SoochikaFileID
    
    
End Function
Private Function SaveSoochikaInwardDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    Dim Rec As New ADODB.Recordset
    Dim arrOut As Variant

    ReDim arrIn(36)
    arrIn(0) = cmbCorrespondance.ItemData(cmbCorrespondance.ListIndex)
    arrIn(1) = cmbPriority.ItemData(cmbPriority.ListIndex)
    If (chkInstitution.Value = 1) Then
        arrIn(2) = 1
        arrIn(3) = txtInstitutionName.Text
        arrIn(4) = txtInstitutionDesg.Text
    Else
        arrIn(2) = Null
        arrIn(3) = Null
        arrIn(4) = Null
    End If
    arrIn(5) = cmbGender.ItemData(cmbGender.ListIndex)
    arrIn(6) = txtApplicantName.Text
    arrIn(7) = txtHouseName.Text
    arrIn(8) = txtWardNo.Text
    arrIn(9) = txtDoorNo1.Text
    arrIn(10) = txtDoorNo2.Text
    arrIn(11) = txtMainPlace.Text
    If (txtLocalPlace.Text = "") Then
        arrIn(12) = txtMainPlace.Text
    Else
        arrIn(12) = txtLocalPlace.Text
    End If
    arrIn(13) = txtPostOffice.Text
    arrIn(14) = txtPincode.Text
    arrIn(15) = cmbDistrict.ItemData(cmbDistrict.ListIndex)
    arrIn(16) = cmbState.ItemData(cmbState.ListIndex)
    arrIn(17) = txtContactNo.Text
    arrIn(18) = txtContactMail.Text
    arrIn(19) = txtSubID.Text
    arrIn(20) = frmUSevanaInward.txtSubTypeID.Text
    arrIn(21) = txtSubject.Text
    'CHNAGED
    If (gbSevanaMainTypeID = 0) Then
    arrIn(22) = dtpDeliveryDate.Value
    Else
    arrIn(22) = frmUSevanaInward.dtpDeliveryDate1.Value
    End If
    arrIn(23) = gbSuitID
    arrIn(24) = Null
    If (chkCourtfee.Value = 1) Then
        arrIn(25) = 1
    Else
        arrIn(25) = Null
    End If
    arrIn(26) = txtNoofPages.Text
    'NOV 13
    If (gbSevanaMainTypeID = 0) Then
    arrIn(27) = cmbSeatID.Text
    Else
     arrIn(27) = frmUSevanaInward.cmbSeatID.Text
     End If
    'arrIn(27) = cmbSeatID.Text
    arrIn(28) = gbnumSeatID
    arrIn(29) = arrIn(27)
    arrIn(30) = 0   'Changed by Renjitha on 29.05.2012
    arrIn(31) = gbLBID
    arrIn(32) = gbnumZonalID
    arrIn(33) = "Soochika Saankhya inward module "
    arrIn(34) = Null
    If (gbLBID = 167) Then
    arrIn(35) = Null
    arrIn(36) = Null
    End If
   ' Unload frmUSevanaInward
    Set Rec = objdb.ExecuteSP("SpSaveInwardDetails", arrIn, arrOut, , mCnn, adCmdStoredProc)
    If (IsArray(arrOut) = True) Then
        SaveSoochikaInwardDetails = arrOut(0, 0)
    End If
End Function
Private Sub SaveSoochikaKeywordDetails(ByVal mCnn As ADODB.Connection)
    Dim arrIn As Variant
    ReDim arrIn(5)
    
    arrIn(0) = SoochikaFileID
    arrIn(1) = DistributionID
    arrIn(2) = FunctionID
    arrIn(3) = ReferenceID
    arrIn(4) = cmbWardName.ItemData(cmbWardName.ListIndex)
    arrIn(5) = cmbWardMember.ItemData(cmbWardMember.ListIndex)
    
    objdb.ExecuteSP "spSaveInwardKeywords ", arrIn, , , mCnn, adCmdStoredProc
End Sub
Private Sub SaveSoochikaInwardTrackDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    Dim Rec As New ADODB.Recordset
    ReDim arrIn(9)

    arrIn(0) = SoochikaFileID
    If (gbSevanaMainTypeID = 0) Then
    arrIn(1) = cmbSeatID.Text
    Else
     arrIn(1) = frmUSevanaInward.cmbSeatID.Text
     End If
    arrIn(2) = gbnumSeatID
    If (gbSevanaMainTypeID = 0) Then
    Set Rec = mCnn.Execute("SpSelectSeatDetails " & CDbl(cmbSeatID.Text))
    Else
    Set Rec = mCnn.Execute("SpSelectSeatDetails " & CDbl(frmUSevanaInward.cmbSeatID.Text))
    End If
    If Not (Rec.EOF Or Rec.BOF) Then
        arrIn(3) = IIf(IsNull(Rec!numCurrentUserID), 0, Rec!numCurrentUserID)
    Else
        arrIn(3) = Null
    End If
    Rec.Close
    arrIn(4) = gbnumUserId
    arrIn(5) = "Processing"
    arrIn(6) = Null
    arrIn(7) = Null
    arrIn(8) = 0  'Changed by Renjitha on 29.02.2012 Form 1 to 0
    arrIn(9) = Null
    
    objdb.ExecuteSP "SpSaveInwardTrackDetails", arrIn, , , mCnn, adCmdStoredProc
End Sub

Public Function updateseat(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    Dim Rec As New ADODB.Recordset
    ReDim arrIn(2)

    arrIn(0) = SoochikaFileID
    arrIn(1) = txtseatid.Text
    arrIn(2) = txtuserid.Text
   objdb.ExecuteSP "spupdateseat", arrIn, , , mCnn, adCmdStoredProc
End Function


Private Sub SaveSoochikaReceiptDetails(mCnn As ADODB.Connection, ByVal ReceiptNO As Variant, ByVal ReceiptBookNo As Variant, ByVal ReceiptAmount As Variant)
    Dim arrIn As Variant
    ReDim arrIn(6)
    
    arrIn(0) = SoochikaFileID
    arrIn(1) = Null
    arrIn(2) = Date
    arrIn(3) = Null
    arrIn(4) = Null
    arrIn(5) = gbnumUserId
    arrIn(7) = gbnumSeatID
    
    objdb.ExecuteSP "SpSaveInwardReceiptDetails", arrIn, , , mCnn, adCmdStoredProc
End Sub
Private Sub SaveSoochikaCertificateDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    ReDim arrIn(13)
    
    arrIn(0) = SoochikaFileID
    arrIn(1) = txtCertName.Text
    arrIn(2) = txtCertHouseName.Text
    arrIn(3) = txtCertWardNo.Text
    arrIn(4) = txtCertDoorNo1.Text
    arrIn(5) = txtCertDoorNo2.Text
    arrIn(6) = txtCertMainPlace.Text
    arrIn(7) = txtCertLocalPlace.Text
    arrIn(8) = txtPostOffice.Text
    arrIn(9) = txtPincode.Text
    'changed by soumya VS on 18.08
    arrIn(10) = cmbCertDistrict.ItemData(cmbCertDistrict.ListIndex)
    arrIn(11) = cmbCertState.ItemData(cmbCertState.ListIndex)
    arrIn(12) = cmbCertGender.ItemData(cmbCertGender.ListIndex)
    arrIn(13) = 1
    
    objdb.ExecuteSP "SpSaveInwardCertificateAddress", arrIn, , , mCnn, adCmdStoredProc
End Sub
Private Sub SaveSoochikaBillReceiptDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    ReDim arrIn(4)
    
    arrIn(0) = SoochikaFileID
    arrIn(1) = cmbBillReceiptType.ItemData(cmbBillReceiptType.ListIndex)
    arrIn(2) = txtBillReceiptNo.Text
    arrIn(3) = txtBillReceiptAmount.Text
    arrIn(4) = txtBillReceiptDescription.Text
    
    objdb.ExecuteSP "SpSaveInwardBillReceipt", arrIn, , , mCnn, adCmdStoredProc
End Sub
Private Sub SaveSoochikaRegisteredPostDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    ReDim arrIn(3)
    
    arrIn(0) = SoochikaFileID
    arrIn(1) = txtRegToWhome.Text
    arrIn(2) = txtRegDesg.Text
    arrIn(3) = txtRegPostNo.Text
    
    objdb.ExecuteSP "SpSaveInwardRegisteredPost", arrIn, , , mCnn, adCmdStoredProc
End Sub
Private Sub SaveSoochikaReferenceDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    
    For i = 1 To grvRefDetails.Rows - 1
        ReDim arrIn(2)
        arrIn(0) = SoochikaFileID
        arrIn(1) = grvRefDetails.TextMatrix(i, 1)
        arrIn(2) = grvRefDetails.TextMatrix(i, 2)
        
        objdb.ExecuteSP "SpSaveInwardReferenceDetails", arrIn, , , mCnn, adCmdStoredProc
    Next i
End Sub
Private Sub SaveSoochikaEnclosureDetails(mCnn As ADODB.Connection)
'changed by soumya vs
    Dim arrIn As Variant
    If (gbSevanaIntegration = 1 And gbSevanaMainTypeID <> 0) Then
      ReDim arrIn(2)
    For i = 1 To frmUSevanaInward.grvCheckList.Rows - 1
     If (frmUSevanaInward.grvCheckList.Cell(flexcpChecked, i, 1) = vbChecked) Then
           arrIn(0) = SoochikaFileID
            arrIn(1) = frmUSevanaInward.grvCheckList.TextMatrix(i, 3)
            arrIn(2) = frmUSevanaInward.grvCheckList.TextMatrix(i, 2)
          objdb.ExecuteSP "SpSaveInwardEnclosure", arrIn, , , mCnn, adCmdStoredProc
          End If
        Next
    Else
    ReDim arrIn(2)
    For i = 1 To grvCheckList.Rows - 1
         
          If (frmUSoochikaInward.grvCheckList.Cell(flexcpChecked, i, 1) = vbChecked) Then
          
          'vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked
        'If (frmUSoochikaInward.grvCheckList.TextMatrix(i, 1) = -1) Then
            arrIn(0) = SoochikaFileID
            arrIn(1) = frmUSoochikaInward.grvCheckList.TextMatrix(i, 3)
            arrIn(2) = frmUSoochikaInward.grvCheckList.TextMatrix(i, 2)
        
            objdb.ExecuteSP "SpSaveInwardEnclosure", arrIn, , , mCnn, adCmdStoredProc
            End If
        'End If
    Next
    End If
End Sub
Private Sub SaveSoochikaValuableDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    
    For i = 1 To grvValuables.Rows - 1
        If (grvValuables.TextMatrix(i, 1) <> "") Then
            ReDim arrIn(5)
            arrIn(0) = SoochikaFileID
            arrIn(1) = grvValuables.TextMatrix(i, 1)
            arrIn(2) = grvValuables.TextMatrix(i, 2)
            arrIn(3) = grvValuables.TextMatrix(i, 3)
            arrIn(4) = grvValuables.TextMatrix(i, 4)
            arrIn(5) = grvValuables.TextMatrix(i, 5)
            
            objdb.ExecuteSP "SpSaveInwardValuables", arrIn, , , mCnn, adCmdStoredProc
        End If
    Next
End Sub
Private Sub SaveSoochikaSevanaInwardDetails(mCnn As ADODB.Connection, SevanaID As Variant, ReceiptNO As Variant, ReceiptBookNo As Variant, ReceiptDate As Variant, ReceiptAmount As Variant)
    Dim arrIn As Variant
    ReDim arrIn(12)
    
    arrIn(0) = SoochikaFileID
    arrIn(1) = SevanaID
    arrIn(2) = Null
    arrIn(3) = frmSevanaInward.txtSubTypeID.Text
    arrIn(4) = frmSevanaInward.cboHospitals.ItemData(frmSevanaInward.cboHospitals.ListIndex)
    arrIn(5) = frmSevanaInward.cboHospitals.Text
    arrIn(6) = ReceiptNO
    arrIn(7) = ReceiptBookNo
    arrIn(8) = ReceiptDate
    arrIn(9) = ReceiptAmount
    arrIn(10) = chkBPL.Value
    arrIn(11) = chkSCST.Value
    arrIn(12) = txtDocProof.Text
    
    objdb.ExecuteSP "SpSaveInwardSevanaDetails", arrIn, , , mCnn, adCmdStoredProc
End Sub

Public Function SaveSevana(ByVal FileID As Variant, ByVal TypeID As Variant, ByVal KioskID As Variant, ByVal mReceiptNo As Variant, mAmt As Double, ByRef mCnn As Connection) As Boolean
    
    Dim arrIn As Variant
    Dim arrOut As Variant
    Dim ForwardTo As Variant
    Dim mVarrReceipt As Variant
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    'NOV18
    Dim mconnection As New ADODB.Connection
    ReDim arrIn(25)
    
    arrIn(23) = 0
    arrIn(0) = Right(FileID, 6)                             'Inward No
    arrIn(1) = Format(Date, "DD/MM/YYYY")                   'Inward Date
    arrIn(2) = gbSevanaMainTypeID                           'Main Sub ID
    If frmUSevanaInward.cboHospitals.ListIndex >= 0 Then     'Hospital
        arrIn(3) = frmUSevanaInward.cboHospitals.ItemData(frmUSevanaInward.cboHospitals.ListIndex)
    Else
        arrIn(3) = 0
    End If
    arrIn(4) = KioskID                                      'Forward To
    'arrIn(5) = Format(frmUSevanaInward.DTPApplDate.value, "DD/MM/YYYY")         'Application Date
    'arrIn(5) = Format(Date, "DD/MM/YYYY")
    'CHANGED
     arrIn(5) = Format(frmUSevanaInward.dtpEventDate.Value, "DD/MM/YYYY")
    If txtWardNo.Text = "" Then                             'Ward No
        arrIn(6) = 0
    Else
        arrIn(6) = txtWardNo.Text
    End If
    arrIn(7) = txtMainPlace.Text                            'Place(Locality)
    If txtDoorNo1.Text = "" Then                            'House No
        arrIn(8) = ""
    Else
        arrIn(8) = IIf(IsNull(txtDoorNo1.Text), 0, txtDoorNo1.Text) & "/" & IIf(IsNull(txtDoorNo2.Text), "", txtDoorNo2.Text) 'House Number
    End If
    arrIn(9) = txtHouseName.Text                          'House Name
    arrIn(10) = ""                                        'Street Name
    arrIn(11) = ""                                        'Via
    arrIn(12) = 0                                         'Postoffice
    arrIn(13) = 0                                         'Village
    arrIn(14) = txtApplicantName.Text                     'Name of Applicant
    arrIn(15) = 0                                         'Taluk
    arrIn(16) = cmbDistrict.ItemData(cmbDistrict.ListIndex) 'District
    arrIn(17) = cmbState.ItemData(cmbState.ListIndex)     'State
    arrIn(18) = 0                                         'Care off ID
    arrIn(19) = frmUSevanaInward.cboSubType.ItemData(frmUSevanaInward.cboSubType.ListIndex) 'SubTypeID
    If chkInsideLB.Value = 1 Then
        arrIn(20) = chkInsideLB.Value                      'Polocn
    Else
        arrIn(20) = 2
    End If
    arrIn(21) = ""                                        'Covering Letter
    frmUSevanaInward.txtRemarks.Text = "Data entered by " & gbUserName & ". " & frmUSevanaInward.txtRemarks.Text
    If frmUSevanaInward.chkZonal.Value = 1 Then
        arrIn(22) = "Inward from Zonal office " & frmUSevanaInward.txtRemarks.Text
    Else
        arrIn(22) = frmUSevanaInward.txtRemarks.Text       'Remarks
    End If
    arrIn(24) = ""                                        'Careoff Name
    arrIn(25) = 0                                         'Inward sequential flag
    
    objdb.ExecuteSP "spSaveInwardFromSoochika", arrIn, arrOut, , mCnn, adCmdStoredProc
    
            'NOV18
    ReDim arrIn(12)
    objdb.CreateNewConnection mconnection, enuSourceString.SoochikaUnicode
    arrIn(0) = SoochikaFileID
    arrIn(1) = Right(FileID, 6)
    arrIn(2) = txtSubID.Text
    arrIn(3) = frmUSevanaInward.txtSubTypeID.Text
    'LATEST 24NOV
     If frmUSevanaInward.cboHospitals.ListIndex > -1 Then
     arrIn(4) = frmUSevanaInward.cboHospitals.ItemData(frmUSevanaInward.cboHospitals.ListIndex)
     Else
     arrIn(4) = Null
     End If
    arrIn(5) = frmUSevanaInward.cboHospitals.Text
    arrIn(6) = mReceiptNo
    arrIn(7) = 0
    arrIn(8) = Format(Now, "dd/MM/yyyy")
    arrIn(9) = mAmt
    arrIn(10) = chkBPL.Value
    arrIn(11) = chkSCST.Value
    arrIn(12) = txtDocProof.Text
    
    objdb.ExecuteSP "SpSaveInwardSevanaDetails", arrIn, , , mconnection, adCmdStoredProc
    
    If TypeID = 1 Or TypeID = 2 Then
        ReDim arrIn(13)
        
        arrIn(0) = arrOut(0, 0)                             'IntID from tInward
        arrIn(1) = Format(Date, "DD/MM/YYYY")               'Receipt No
        arrIn(2) = 0                                        'Receipt Book
        arrIn(3) = mReceiptNo                               'Receipt No
        arrIn(4) = mAmt                                     'Receipt Amount
        If gbSevanaMainTypeID = 5 Then
            arrIn(5) = frmUSevanaInward.txtNoCopeis.Text
        Else
            arrIn(5) = frmUSevanaInward.txtNoOfCertificate.Text    'No of copies
        End If
        arrIn(6) = frmUSevanaInward.txtEnglishname.Text      'English Name
        arrIn(7) = frmUSevanaInward.txtMalayalamname.Text    'Malayalam Name
        If frmUSevanaInward.cboRelationship.ListIndex > -1 Then
            arrIn(8) = frmUSevanaInward.cboRelationship.ItemData(frmUSevanaInward.cboRelationship.ListIndex) 'CFM
        Else
            arrIn(8) = Null
        End If
        arrIn(9) = frmUSevanaInward.cboLanguage.ItemData(frmUSevanaInward.cboLanguage.ListIndex) 'Language
        arrIn(10) = Right(FileID, 6)                        'Inward No
        arrIn(11) = 0 'gbFldUserId                          'Issue User
        arrIn(12) = frmUSevanaInward.txtRegNo.Text           'Register No
        arrIn(13) = Trim(frmUSevanaInward.txtBookNo.Text)          'Book No
        
        objdb.ExecuteSP "InsertReceiptDetails", arrIn, , , mCnn, adCmdStoredProc
        
        'changed by soumya 04 July
        
        'SaveSevanaStatus Right(FileID, 6), Year(Now()), Now(), frmUSevanaInward.txtSubTypeID.Text, gbSevanaMainTypeID
        

    
       ReDim arrIn(4)
       arrIn(0) = Right(FileID, 6)
       arrIn(1) = Year(Now())
       arrIn(2) = Format(Now, "dd/MM/yyyy")
       arrIn(3) = gbSevanaMainTypeID
       arrIn(4) = frmUSevanaInward.txtSubTypeID.Text
       objdb.ExecuteSP "sp_insertinwardstatusSoochika", arrIn, , , mCnn, adCmdStoredProc
      
       SaveSevana = True
    End If
End Function
Public Sub ShowAckReport(ByVal FileID As Variant)
    If (FileID = 0) Then
        MsgBox "Print is not Possible !!!", vbInformation, "SOOCHIKA"
        Exit Sub
    Else
        Dim arrIn(2)
        arrIn(0) = CStr(SoochikaFileID)
        arrIn(1) = 1
        frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptAckSlip.rpt", arrIn
        frmCRViewer.Show 1
    End If
    
End Sub

Private Sub chkaddr_Click()
     Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim Rec1 As New ADODB.Recordset
        Dim arrIn As Variant
        Dim mSql As String
       If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
        If (chkaddr.Value = 1) Then
        txtInwNo.Enabled = False
        lblinwNo.Enabled = False
        lblYear.Enabled = False
        txtYear.Enabled = False
        Else
        txtInwNo.Enabled = True
        lblinwNo.Enabled = True
        lblYear.Enabled = True
        txtYear.Enabled = True
        End If
        
        
        
        mSql = "select top 1 intGender,chvApplicantName,chvHouseName,intWardNo,intDoorNO1,chvDoorNo2,chvLocalPlace,chvMainPlace,intPincode,chvPostoffice,chvContactNo,chvContactMail,numFileid from tInwardDetails order by tInwardDetails.numFileid desc"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
        cmbGender.ListIndex = IIf(IsNull(Rec!intGender), 0, Rec!intGender)
        txtApplicantName.Text = Rec!chvApplicantName
        txtHouseName.Text = IIf(IsNull(Rec!chvHouseName), 0, Rec!chvHouseName)
        txtWardNo.Text = IIf(IsNull(Rec!intWardNo), 0, Rec!intWardNo)
        txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo1), 0, Rec!intDoorNo1)
        txtDoorNo1.Text = IIf(IsNull(Rec!chvDoorNo2), 0, Rec!chvDoorNo2)
        txtLocalPlace.Text = IIf(IsNull(Rec!chvLocalPlace), 0, Rec!chvLocalPlace)
        txtMainPlace.Text = IIf(IsNull(Rec!chvMainPlace), 0, Rec!chvMainPlace)
        txtPincode.Text = IIf(IsNull(Rec!intPincode), 0, Rec!intPincode)
        txtPostOffice.Text = IIf(IsNull(Rec!chvPostoffice), 0, Rec!chvPostoffice)
        txtContactNo.Text = IIf(IsNull(Rec!chvContactNo), 0, Rec!chvContactNo)
        txtContactMail.Text = IIf(IsNull(Rec!chvContactMail), 0, Rec!chvContactMail)
        End If
        
        
        mSql = "SELECT  intSalutation,chvName,chvHouseName,intWardNo,intDoorNo1,chvDoorNo2,chvMainPlace,chvLocalPlace,chvPostoffice,intPincode,intDistrictID,intStateID,numFileID FROM tInwardCertificateAddress WHERE numFileID=" & Rec!numFileID
        Rec1.Open mSql, mCnn
        If Not (Rec1.EOF And Rec1.BOF) Then
        cmbCertGender.ListIndex = IIf(IsNull(Rec1!intSalutation), 0, Rec1!intSalutation)
        txtCertName.Text = IIf(IsNull(Rec1!chvName), 0, Rec1!chvName)
        txtCertHouseName.Text = IIf(IsNull(Rec1!chvHouseName), 0, Rec1!chvHouseName)
        txtCertWardNo.Text = IIf(IsNull(Rec1!intWardNo), 0, Rec1!intWardNo)
        txtCertDoorNo1.Text = IIf(IsNull(Rec1!intDoorNo1), 0, Rec1!intDoorNo1)
        txtCertDoorNo2.Text = IIf(IsNull(Rec1!chvDoorNo2), 0, Rec1!chvDoorNo2)
        txtCertMainPlace.Text = IIf(IsNull(Rec1!chvMainPlace), 0, Rec1!chvMainPlace)
        txtCertLocalPlace.Text = IIf(IsNull(Rec1!chvLocalPlace), 0, Rec1!chvLocalPlace)
        txtCertPostOffice.Text = IIf(IsNull(Rec1!chvPostoffice), 0, Rec1!chvPostoffice)
        txtCertPincode.Text = IIf(IsNull(Rec!intPincode), 0, Rec1!intPincode)
        cmbCertDistrict.ListIndex = IIf(IsNull(Rec1!intDistrictID), 0, Rec1!intDistrictID)
        cmbCertState.ListIndex = IIf(IsNull(Rec1!intStateID), 0, Rec1!intStateID)
        End If
    
        If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub

Private Sub chkAsInward_Click()
If (chkAsInward.Value = 1) Then
Call InwardAddress
Else
Call CertificateAddressClear
End If
End Sub
Private Sub CertificateAddressClear()
            cmbCertGender.ListIndex = 0
            txtCertName.Text = ""
            txtCertHouseName.Text = ""
            txtCertWardNo.Text = ""
            txtCertDoorNo1.Text = ""
            txtCertDoorNo2.Text = ""
            txtCertLocalPlace.Text = ""
            txtCertMainPlace.Text = ""
            txtCertPincode.Text = ""
            txtCertPostOffice.Text = ""
            cmbCertDistrict.ListIndex = 0
            cmbCertState.ListIndex = cmbState.ListIndex
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmbseat_Change()
cmbSeatID.ListIndex = cmbSeat.ListIndex
End Sub

Private Sub cmbseat_Click()
'changed by soumya
cmbSeatID.ListIndex = cmbSeat.ListIndex
End Sub

Private Sub cmbSeatID_Change()
Dim arrIn As Variant
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
ReDim arrIn(0)
arrIn(0) = cmbSeatID.Text
Set Rec = objdb.ExecuteSP("spSelectUser", arrIn, , , mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblusername.Caption = Rec!chvUserNameEng
Else
lblusername.Caption = ""
End If
End Sub

Private Sub cmbSeatID_Click()
Dim arrIn As Variant
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
ReDim arrIn(0)
arrIn(0) = cmbSeatID.Text
Set Rec = objdb.ExecuteSP("spSelectUser", arrIn, , , mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblusername.Caption = Rec!chvUserNameEng
Else
lblusername.Caption = ""
End If
End Sub

Private Sub cmdGo_Click()
'CHANGED
     Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim Rec1 As New ADODB.Recordset
        Dim arrIn As Variant
        Dim mSql As String
     
If (txtInwNo.Text = "") Then
MsgBox "Please enter Inward Number for Searching"
txtInwNo.SetFocus
chkaddr.Enabled = False
End If
If (txtYear.Text = "") Then
MsgBox "Please enter Inward year for Searching"
txtYear.SetFocus

Else
  If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
        If (chkaddr.Value = 0) Then
         
        mSql = "select top 1 intCorrespondanceType,intPriority,numMainSubjectID,chvSubject,intGender,chvApplicantName,chvHouseName, "
        mSql = mSql + " intWardNo,intDoorNO1,chvDoorNo2,chvLocalPlace,chvMainPlace,intPincode,chvPostoffice,intStateID,intDistrictID,chvContactNo,chvContactMail,"
        mSql = mSql + " dtDeliveryDate,tInwardTrackDetails.numCurrentSeatID,tInwardDetails.numFileid"
        mSql = mSql + " from tInwardDetails inner join tInwardTrackDetails on tInwardTrackDetails.numFileid=tInwardDetails.numFileid"
        'CHANGED
        'mSQl = mSQl + " inner join tSeatdetails on tSeatDetails.numseatID=tInwardTrackDetails.numPreviousSeatID"
        mSql = mSql + " where numCurrentno=" + txtInwNo.Text + " and year(dtDateofReceipt)=" + txtYear.Text + " order by tINwardTrackDetails.intid asc"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
        'If (chkaddr.value = 1) Then
        'NOV18
        'cmbCorrespondance.ListIndex = Rec!intCorrespondanceType
        'cmbPriority.ListIndex = Rec!intPriority
        'changed
        'txtSubID.Text = Rec!numMainSubjectID
         'txtSubID.Text = IIf(IsNull(Rec!numMainSubjectID), "", Rec!numMainSubjectID)
        'txtSubject.Text = Rec!chvSubject
        cmbGender.ListIndex = IIf(IsNull(Rec!intGender), "", Rec!intGender)
        txtApplicantName.Text = IIf(IsNull(Rec!chvApplicantName), "", Rec!chvApplicantName)
        txtHouseName.Text = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
        txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
        txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo1), "", Rec!intDoorNo1)
        txtDoorNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
        txtLocalPlace.Text = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
        txtMainPlace.Text = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
        txtPincode.Text = IIf(IsNull(Rec!intPincode), "", Rec!intPincode)
        txtPostOffice.Text = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
        Call PopulateList(cmbState, "Sp_SelectState 1", , False, True, True, enuSourceString.SoochikaUnicode)
        cmbState.ListIndex = IIf(IsNull(Rec!intStateID), "", Rec!intStateID)
        Call PopulateList(cmbDistrict, "Sp_SelectDistrict 1," & cmbState.ItemData(cmbState.ListIndex), , False, True, True, enuSourceString.SoochikaUnicode)
        cmbDistrict.ListIndex = IIf(IsNull(Rec!intDistrictID), "", Rec!intDistrictID)
        txtContactNo.Text = IIf(IsNull(Rec!chvContactNo), "", Rec!chvContactNo)
        txtContactMail.Text = IIf(IsNull(Rec!chvContactMail), "", Rec!chvContactMail)
        dtpDeliveryDate.Value = IIf(IsNull(Rec!dtDeliveryDate), "", Rec!dtDeliveryDate)
        'cmbDepartment.ListIndex = IIf(IsNull(Rec!intDeptID), 0, Rec!intDeptID)
        'cmbSeat.ListIndex = IIf(IsNull(Rec!numCurrentSeatID), 0, Rec!numCurrentSeatID)
        cmdSave.Enabled = False
    
        
        'End If
        mSql = "SELECT     intSalutation, chvName, chvHouseName, intWardNo, intDoorNo1, chvDoorNo2, chvMainPlace, chvLocalPlace, chvPostoffice, intPincode, intDistrictID,intStateID,numFileID"
        mSql = mSql + " FROM  tInwardCertificateAddress WHERE numFileID=" & Rec!numFileID
        Rec1.Open mSql, mCnn
        If Not (Rec1.EOF And Rec1.BOF) Then
        cmbCertGender.ListIndex = IIf(IsNull(Rec1!intSalutation), "", Rec1!intSalutation)
        txtCertName.Text = IIf(IsNull(Rec1!chvName), "", Rec1!chvName)
        txtHouseName.Text = IIf(IsNull(Rec1!chvHouseName), "", Rec1!chvHouseName)
        txtWardNo.Text = IIf(IsNull(Rec1!intWardNo), "", Rec1!intWardNo)
        txtDoorNo1.Text = IIf(IsNull(Rec1!intDoorNo1), "", Rec1!intDoorNo1)
        txtDoorNo2.Text = IIf(IsNull(Rec1!chvDoorNo2), "", Rec1!chvDoorNo2)
        txtMainPlace.Text = IIf(IsNull(Rec1!chvMainPlace), "", Rec1!chvMainPlace)
        txtLocalPlace.Text = IIf(IsNull(Rec1!chvLocalPlace), "", Rec1!chvLocalPlace)
        txtPostOffice.Text = IIf(IsNull(Rec1!chvPostoffice), "", Rec1!chvPostoffice)
        txtPincode.Text = IIf(IsNull(Rec1!intPincode), "", Rec1!intPincode)
        cmbCertDistrict.ListIndex = Rec1!intDistrictID
        cmbCertState.ListIndex = Rec1!intStateID
        Call chkInsideLB_Click
      End If
      
        Else
        MsgBox "Inward No Doesnot Exist"
        chkaddr.Enabled = True
        End If
      
        End If
        'CHNAGED
        cmdSave.Enabled = True
        
        End If
    
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitIDESubClassing
    'CheckInwardDate,User and Inward NO
    MsgBox "Inward date : " & Date, vbInformation, "SOOCHIKA"
    SetSoochkaEnvironment
    DisableControls
    If (CheckVersion = False) Then
        MsgBox "Version is Out Dated !!!" & vbCrLf & "Please Contact Administrator .....", vbInformation, "SOOCHIKA"
        frmUSoochikaInward.Caption = "Soochika Application version is out dated"
        cmdNew.Enabled = False
        Exit Sub
    ElseIf (CheckInwardDate = False) Then
        MsgBox "Previous date inward entry not supported !!!" & vbCrLf & "Please Check System Date .....", vbInformation, "SOOCHIKA"
        cmdNew.Enabled = False
        Exit Sub
    ElseIf (CheckInwardUser(gbnumUserId) = False) Then
        MsgBox "User not permitted to take inwards !!!" & vbCrLf & "Please Contact Administrator .....", vbInformation, "SOOCHIKA"
        cmdNew.Enabled = False
        Exit Sub
    ElseIf lInitialise = False Then
        cmdNew.Enabled = False
        frmSoochikaStartup.Show vbModal
        Exit Sub
    Else
        frmUSoochikaInward.Caption = "New Inward"
        gbSoochikaUserLogID = SaveUserLog
        IntegrationSuites
        FillCombos
        Clear
        DisableControls
    End If
End Sub
Private Function lInitialise() As Boolean
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim mCount As Integer
        
    If objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False Then
        MsgBox "Cannot Continue.., Connection not present", vbInformation, "Soochika"
        Exit Function
    End If
    mSql = "Select * from tInterruption where flgReason=1"
    Rec.Open mSql, mCnn
    If (Rec.EOF And Rec.BOF) Then
        lInitialise = True
    Else
        lInitialise = False
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Sub cmbCorrespondance_Click()
    If ((cmbCorrespondance.ListIndex > 0 And cmbCorrespondance.ListIndex <= 4) Or (cmbCorrespondance.ListIndex = 8)) Then
        cmbPriority.ListIndex = 3
    Else
        cmbPriority.ListIndex = 4
    End If
End Sub

Private Sub Label3_Click()
    frmUSoochikaSubjectMaster.Show vbModal
End Sub

Private Sub txtSearch_Change()

End Sub

Private Sub lblsubjectmaster_Click()
'CHANGED
frmUSoochikaSubjectMaster.Show vbModal
End Sub

Private Sub txtContactNo_Change()
'changed by soumya VS
Dim textval As String
Dim numval As String
textval = txtContactNo.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    txtContactNo.Text = CStr(numval)
  End If
End Sub

Private Sub txtSubID_Change()
    If (txtSubID.Text <> "") Then
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim arrIn As Variant
        
        ReDim arrIn(1)
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
        
        
        arrIn(0) = 1
        arrIn(1) = txtSubID.Text
        'CHANGED
        lblusername.Caption = ""
        
        Set Rec = objdb.ExecuteSP("SP_SelectSubject", arrIn, , False, mCnn, adCmdStoredProc)
        If Not (Rec.EOF Or Rec.BOF) Then
            txtSubject.Text = Rec!chvSubject
            lstSubject.Visible = False
            DistributionID = IIf(IsNull(Rec!intDistrID), 0, Rec!intDistrID)
            FunctionID = IIf(IsNull(Rec!intFuncID), 0, Rec!intFuncID)
            ReferenceID = IIf(IsNull(Rec!intRefID), 0, Rec!intRefID)
        Else
            MsgBox "Invalid Subject ID !!!", vbInformation, "SOOCHIKA"
            txtSubject.Text = ""
            txtSubID.Text = ""
            DistributionID = 0
            FunctionID = 0
            ReferenceID = 0
        End If
        'CHNAGED
        'NOV18
         'gbSevanaMainTypeID = CheckSevanaInward(txtSubID.Text)
        'If (gbSevanaMainTypeID = 0) Then
        GetSubjectSeat
        'Else
        'cmbDepartment.Enabled = False
       ' cmbSeat.Enabled = False
       ' lblusername.Caption = ""
        
        
        'End If
        grvCheckList.Rows = 2
        grvCheckList.Clear 1
        'CHNAGED
        gbSevanaMainTypeID = CheckSevanaInward(txtSubID.Text)
        If (gbSevanaMainTypeID = 0) Then
        FillEnclosureGrid val(txtSubID.Text)
        Else
        grvCheckList.Rows = 0
    
        End If
        getSubjectDeliverydate val(txtSubID.Text)
        GetRCReference txtSubject.Text
         gbSevanaMainTypeID = CheckSevanaInward(txtSubID.Text)
        If (gbSevanaMainTypeID = 0) Then
         cmbDepartment.Enabled = True
        cmbSeat.Enabled = True
        End If
        
        
        'gbSaankhya = CheckReceipt(txtSubID.Text)
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
    End If
End Sub
Private Sub txtSubID_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtSubject_Change()
    If (txtSubject.Text <> "") Then
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
        
        mSql = "select chvSubject from mSubject where chvsubject like '%" & txtSubject.Text & "%'"
        Rec.Open mSql, mCnn
        i = 0
        If Not (Rec.EOF Or Rec.BOF) Then
            lstSubject.Clear
            While (Not Rec.EOF)
                lstSubject.AddItem Rec!chvSubject, i
                Rec.MoveNext
                i = i + 1
            Wend
            lstSubject.Visible = True
        End If
        Rec.Close
    End If
End Sub
Private Sub txtSubject_LostFocus()
     If (txtSubject.Text <> "") Then
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
        
        mSql = "select numSubjectID from mSubject where chvsubject = '" & txtSubject.Text & "'"
        Rec.Open mSql, mCnn
        i = 0
        If Not (Rec.EOF Or Rec.BOF) Then
            txtSubID.Text = Rec!numSubjectID
        Else
            txtSubID.Text = ""
            gbSevanaMainTypeID = 0 'Ranjitha 09/10
        End If
        Rec.Close
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
        GetRCReference txtSubject.Text
    End If
End Sub
Private Sub lstSubject_DblClick()
    txtSubject.Text = lstSubject.Text
    lstSubject.Visible = False
    txtSubject.SetFocus
End Sub
Private Sub txtRefNo_Change()
    If (txtRefNo.Text <> "") Then
        dtpRefDate.Enabled = True
        grvRefDetails.TextMatrix(1, 1) = txtRefNo.Text
        grvRefDetails.TextMatrix(1, 2) = dtpRefDate.Value
    Else
        dtpRefDate.Enabled = False
        grvRefDetails.Clear 1
    End If
End Sub
Private Sub dtpRefDate_Change()
    grvRefDetails.TextMatrix(1, 1) = txtRefNo.Text
    grvRefDetails.TextMatrix(1, 2) = dtpRefDate.Value
End Sub

Private Sub chkInstitution_Click()
    CheckInstitution (chkInstitution.Value)
End Sub
Private Sub chkInsideLB_Click()
    InsideLB (chkInsideLB.Value)
End Sub

Private Sub txtWardNo_KeyPress(KeyAscii As Integer)


    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtWardNo_LostFocus()
'CHANGED
    If (txtWardNo.Text <> "") Then
    'NOV18
    'If (gbSevanaMainTypeID = 0) Then
       GetSubjectSeat
       'Else
       'lblusername.Caption = ""
       'cmbDepartment.Enabled = False
       'cmbSeat.Enabled = False
       'End If
       
       
    End If
End Sub

Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub cmdTaxSearch_Click()
    If (txtWardNo.Text = "") Then
        MsgBox "Please enter Ward No ", vbOKOnly, "SOOCHIKA"
        txtWardNo.SetFocus
        Exit Sub
    ElseIf (txtDoorNo1.Text = "") Then
        MsgBox "Please enter Door No 1 ", vbOKOnly, "SOOCHIKA"
        txtDoorNo1.SetFocus
        Exit Sub
    ElseIf (txtasssyear.Text = "") Then
        MsgBox "Please enter Assessment Year ", vbOKOnly, "SOOCHIKA"
        txtasssyear.SetFocus
        Exit Sub
    Else
        If (txtDoorNo2.Text = "") Then
            txtDoorNo2.Text = 0
        End If
        frmSoochikaBuildingDetails.Show (vbModal)
    End If
End Sub
Private Sub txtPincode_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtPincode_LostFocus()
    If (txtPincode.Text <> "") Then
        If (Len(txtPincode.Text) <> 6) Then
            MsgBox "Invalid Pincode ", vbInformation, "SOOCHIKA"
            Exit Sub
        Else
            GetPostoffice
        End If
    End If
End Sub
Private Sub cmbState_Click()
  
    Call PopulateList(cmbDistrict, "Sp_SelectDistrict 1," & cmbState.ItemData(cmbState.ListIndex), , False, True, True, enuSourceString.SoochikaUnicode)
    cmbDistrict.ListIndex = 0
End Sub
Private Sub chkBPL_Click()
    If (chkBPL.Value = 1 Or chkSCST.Value = 1) Then
        lblMandatory(13).Visible = True
    Else
        lblMandatory(13).Visible = False
    End If
End Sub
Private Sub chkSCST_Click()
    If (chkBPL.Value = 1 Or chkSCST.Value = 1) Then
        lblMandatory(13).Visible = True
    Else
        lblMandatory(13).Visible = False
    End If
End Sub
Private Sub txtNoofPages_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNoofPages_LostFocus()
    If (txtNoofPages.Text <> "" And val(txtNoofPages.Text) > 0) Then
        grvCheckList.Rows = 2
        grvCheckList.Clear 1
        FillEnclosureGrid val(txtSubID.Text)
    End If
End Sub
Private Sub grvValuables_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If (grvValuables.TextMatrix(Row, 3) <> "" And IsDate(grvValuables.TextMatrix(Row, 3)) = False) Then
        MsgBox "Invalid Instrument Date ", vbInformation, "SOOCHIKA"
        grvValuables.TextMatrix(Row, 3) = ""
        Exit Sub
    End If
    If (grvValuables.TextMatrix(Row, 4) <> "" And IsNumeric(grvValuables.TextMatrix(Row, 4)) = False) Then
        MsgBox "Invalid Amount ", vbInformation, "SOOCHIKA"
        grvValuables.TextMatrix(Row, 4) = ""
        Exit Sub
    End If
    If (grvValuables.Rows - 1 = Row) Then
        If (grvValuables.TextMatrix(Row, 5) <> "") Then
            grvValuables.Rows = grvValuables.Rows + 1
        End If
    End If
End Sub
Private Sub grvRefDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If (grvRefDetails.TextMatrix(Row, 2) <> "") Then
        If (IsDate(grvRefDetails.TextMatrix(Row, 2)) = True) Then
            If (grvRefDetails.Rows - 1 = Row) Then
                grvRefDetails.Rows = grvRefDetails.Rows + 1
            End If
        Else
            MsgBox "Invalid Reference Date ", vbInformation, "SOOCHIKA"
            grvRefDetails.TextMatrix(Row, 2) = ""
        End If
    End If
End Sub

Private Sub cmbDepartment_Click()
     If (cmbDepartment.ListIndex <> 0) Then
        'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex), , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex), , True, True, True, enuSourceString.SoochikaUnicode)
        
        'add  by vipin 21-09-2012
       ' Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)


      'LATEST 24NOV
                Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
    End If
End Sub

Private Sub lstSubject_LostFocus()
lstSubject.Visible = False
End Sub
Private Sub cmdSave_Click()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
        'NOV18
    Text1.Text = cmbSeatID.Text
    Text2.Text = cmbDepartment.ItemData(cmbDepartment.ListIndex)
    frmUSevanaInward.txtWardNo.Text = txtWardNo.Text
  
    If (ValidateData = True) Then
        
        If (gbSevanaIntegration = 1 And gbSevanaMainTypeID <> 0) Then
            frmUSevanaInward.Form_Load
            frmUSevanaInward.Show (vbModal)
        Else
            If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
                MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
                Exit Sub
            End If
            On Error GoTo rollback
            mCnn.BeginTrans
            SaveSoochika mCnn
            mCnn.CommitTrans
            MsgBox "Data Save with inward no " & Right(SoochikaFileID, 6), vbInformation, "SOOCHIKA"
            mSql = "SELECT tLBSettings.flgAttachment FROM tLBSettings"
            Set Rec = mCnn.Execute(mSql)
            If (Rec.Fields(0) = "1") Then
            SaveAttachment (SoochikaFileID) 'paperless
            End If
            ShowAckReport (SoochikaFileID)
            DisableControls
            cmdNew.Enabled = True
            cmdReprint.Enabled = True
            cmdNew.SetFocus
        End If
    End If
    
    Exit Sub
rollback:
        MsgBox Error$, vbInformation, "SOOCHIKA"
        mCnn.RollbackTrans
        Exit Sub
SkipUnload:
        
End Sub
Private Sub cmdNew_Click()
    EnableControls
    FillCombos
    Clear
    txtSubID.SetFocus
    cmdNew.Enabled = True
    'changed by soumya vs on 13.01.15
    Label44.Visible = True
    dtpDeliveryDate.Visible = True
    
    On Error GoTo SkipUnload:
    Unload frmReceiptsCounter
SkipUnload:
End Sub
Private Sub cmdReprint_Click()
    ShowAckReport (SoochikaFileID)
End Sub
Private Sub cmdCancel_Click()
    If (MsgBox("Do you want to close the window ? ", vbYesNo, "SOOCHIKA") = vbYes) Then
        Dim mCnn As New ADODB.Connection
    
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
        mSql = "update tUserLog set dtlogoutTime=getdate(),flgLogOut=1 where intID= " & IIf(IsEmpty(gbSoochikaUserLogID), 0, gbSoochikaUserLogID)
        mCnn.Execute mSql
        
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
        Unload Me
    End If
End Sub
'Private Sub txtSearch_LostFocus()
'    If txtSearch.Text <> "" Then
'        Dim mCnn As New ADODB.Connection
'        Dim Rec As New ADODB.Recordset
'        Dim CurrentNo As Variant
'        Dim CurrentYear As Variant
'        Dim Pos As Variant
'
'        If ((Pos = InStr(txtSearch.Text, "/")) = 0) Then
'            MsgBox "Invalid No " & vbCrLf & "InwardNo/Year", vbInformation, "SOOCHIKA"
'            Exit Sub
'        End If
'
'        If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
'            MsgBox "Connection Failure !!", vbInformation, "SOOCHIKA"
'            Exit Sub
'        End If
'
'        mSQL = "Select * from tInwardDetails "
'        mSQL = mSQL & "  left join tInwardReferenceDetails on tInwardDetails.numFileid=tInwardReferenceDetails.numfileid"
'        mSQL = mSQL & "  left join tInwardEnclosureDetails on tInwardDetails.numfileid=tInwardEnclosureDetails.numfileid"
'        mSQL = mSQL & "  left join tInwardValuableDetails on tInwardDetails.numfileid=tInwardValuableDetails.numfileid  "
'        mSQL = mSQL & "  left join tInwardBillReceiptDetails on tInwardDetails.numFileid=tInwardBillReceiptDetails.numfileid"
'        mSQL = mSQL & "  left join tInwardCertificateAddress on tInwardDetails.numfileid=tInwardCertificateAddress.numfileid"
'        mSQL = mSQL & "  left join tInwardRegisteredPostDetails on tInwardDetails.numfileid=tInwardRegisteredPostDetails.numfileid"
'        mSQL = mSQL & "  left join tInwardRCReferenceDetails on tInwardDetails.numfileid=tInwardRCReferenceDetails.numFileid"
'        mSQL = mSQL & "  where numCurrentNo=" & mID(txtSearch.Text, 1, Pos - 1) & " and year(dtdateofreceipt)= " & mID(txtSearch.Text, Pos + 1, Len(txtSearch.Text) - Pos)
'
'        Set Rec = mCnn.Execute(mSQL)
'        If Not (Rec.EOF Or Rec.BOF) Then
'            txtApplicantName.Text = Rec!chvApplicantName
'        End If
'        Rec.Close
'
'        If (mCnn.State = 1) Then
'            mCnn.Close
'        End If
'    Else
'        txtSearch.Text = "InwardNo/Year"
'        txtSearch.ForeColor = vbGrayText
'    End If
'End Sub

'15-06-2012 SMS(Ranjitha)

Private Sub SaveSMS()
    Dim InwardNo As Double
    Dim Amt As Double
    Dim smsMsg As String
    Dim Subject As String
    Dim lbType As String
    Dim arrIn As Variant
    ReDim arrIn(3)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Amt = 0
    smsMsg = ""
    lbType = ""
    If gbLBType = 3 Then
            lbType = "MP"
        ElseIf gbLBType = 5 Then
            lbType = "GP"
        ElseIf gbLBType = 1 Then
            lbType = "DP"
        ElseIf gbLBType = 4 Then
            lbType = "Corp."
        ElseIf gbLBType = 2 Then
            lbType = "BP"
        End If
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    InwardNo = CDbl(Right(SoochikaFileID, 6))
   ' smsMsg = GetPaidAmt(InwardNo, Year(Now))
    smsMsg = gbLBName + " " + lbType + "-" + "Inw.No." + CStr(InwardNo) + "/" + CStr(Year(Now)) + "-" + txtSubject.Text + "-" + txtApplicantName.Text
    arrIn(0) = SoochikaFileID
    arrIn(1) = InwardNo
    arrIn(2) = smsMsg
    arrIn(3) = txtContactNo.Text
    objdb.ExecuteSP "SaveSMS_Inward", arrIn, , , mCnn, adCmdStoredProc
End Sub

'paperless
Private Sub lstSubject_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 39 Then
        lstSubject.Visible = False
        txtRefNo.SetFocus
     ElseIf KeyCode = 13 Then
        txtSubject.Text = lstSubject.Text
        lstSubject.Visible = False
        txtSubject.SetFocus
     End If
End Sub


'paperless
Public Sub SaveAttachment(ByVal FileID As Variant)
Dim res As String
Dim app1 As String
    Dim nCnt As Integer
    Dim X As Variant
    Dim st As String
    Dim st1 As Integer

    If (FileID = 0) Then
        MsgBox "Attachment is not Possible !!!", vbInformation, "SOOCHIKA"
        Exit Sub
    Else
        res = MsgBox("Do you want to Attach a File", vbYesNo)
    If res = vbYes Then
Att:
        CommonDialog1.Filter = "All files (*.*)|*.*"
        CommonDialog1.DialogTitle = "Select File"
        CommonDialog1.ShowOpen
        'CommonDialog1.ShowSave
         app1 = CommonDialog1.FileName
        st1 = InStrRev(CommonDialog1.FileName, "\")
        st = mID(CommonDialog1.FileName, 1, st1 - 1)
        FindFile1 st, app1
        
        res = MsgBox("Do you want to Attach a File", vbYesNo)
    If res = vbYes Then
    GoTo Att
    End If
    End If
    End If
End Sub
'paperless
Private Function FindFile1(ByVal sFol As String, sFile As String) As Long
  Dim tFld As Folder, tFil As File, FileName As String
  Dim st1 As String
  Dim st As String
  Dim strPath As String
  Dim strSp As Variant
  Dim strFile As Variant
  Dim FileID As Variant
  Dim arrIn As Variant
  ReDim arrIn(3)
  Dim Rec As New ADODB.Recordset
  Set fld = fso.GetFolder(sFol)
  strPath = ""
  strFile = ""
  st1 = mID(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\") + 1)
  Dim mCnn As New ADODB.Connection
    
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        End If
  For Each tFil In fld.Files
  
  If tFil.Name = st1 Then _
      st = tFil.Name
    DoEvents
  Next
  
   Rec.Open "SELECT  chvPath From tLBSettings", mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        strPath = Rec!chvPath
    Else
        strPath = ""
    End If
  FileID = CStr(SoochikaFileID)
 strSp = Split(st, ".")
 
 strFile = strPath + strSp(0) + "-" + FileID + "." + strSp(1)
 If (fso.FileExists(strFile)) Then
    MsgBox "Fie exist", vbInformation, "SOOCHIKA"
 Else
 fso.CopyFile st, strPath + strSp(0) + "-" + FileID + "." + strSp(1)
    arrIn(0) = strSp(0)
    arrIn(1) = SoochikaFileID
    arrIn(2) = gbnumUserId
    arrIn(3) = "." + strSp(1)
    objdb.ExecuteSP "Sp_SaveAttachment", arrIn, , , mCnn, adCmdStoredProc
    txtAtt.Text = txtAtt.Text + st + ","
    End If
End Function


