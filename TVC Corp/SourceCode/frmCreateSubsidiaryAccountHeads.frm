VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCreateSubsidiaryAccountHeads 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  "
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13770
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmCreateSubsidiaryAccountHeads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   11940
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   390
      Width           =   1815
   End
   Begin VB.Frame fmeSulekha 
      Height          =   7335
      Left            =   3285
      TabIndex        =   36
      Top             =   8055
      Visible         =   0   'False
      Width           =   8835
      Begin VB.Frame fmeOfficials 
         Height          =   1155
         Left            =   90
         TabIndex        =   47
         Top             =   420
         Width           =   8625
         Begin VB.ComboBox cmbDepartment 
            BackColor       =   &H80000004&
            Height          =   390
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   240
            Width           =   5445
         End
         Begin VB.ComboBox cmbDesignations 
            Height          =   390
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   630
            Width           =   5445
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department :"
            Height          =   240
            Left            =   840
            TabIndex        =   51
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designations :"
            Height          =   240
            Left            =   840
            TabIndex        =   50
            Top             =   690
            Width           =   1065
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid2 
         Height          =   5655
         Left            =   60
         TabIndex        =   37
         Top             =   1590
         Width           =   8715
         _cx             =   15372
         _cy             =   9975
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
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
         Rows            =   21
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCreateSubsidiaryAccountHeads.frx":1CCA
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
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Interface From Sulekha Data Base"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Left            =   90
         TabIndex        =   19
         Top             =   210
         Width           =   8445
      End
      Begin VB.Label lblExitSulekhaPanal 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8580
         TabIndex        =   39
         Top             =   120
         Width           =   165
      End
   End
   Begin VB.Frame fmeCodeTitle 
      Height          =   3225
      Left            =   90
      TabIndex        =   34
      Top             =   4530
      Width           =   6075
      Begin VB.TextBox txtDesignation 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1650
         TabIndex        =   14
         Top             =   1920
         Width           =   4245
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1650
         TabIndex        =   13
         Top             =   1500
         Width           =   4245
      End
      Begin VB.TextBox txtDDOCode 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1650
         TabIndex        =   12
         Top             =   1080
         Width           =   4245
      End
      Begin VB.TextBox txtSubTitle 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1650
         TabIndex        =   11
         Top             =   660
         Width           =   4245
      End
      Begin VB.TextBox txtOpeningBalance 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1650
         MaxLength       =   11
         TabIndex        =   15
         Top             =   2340
         Width           =   4245
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   10
         Top             =   300
         Width           =   4245
      End
      Begin VB.Label lblMsgDept 
         AutoSize        =   -1  'True
         Caption         =   "Disabled for this SubLedger Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   3090
         TabIndex        =   46
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         Height          =   270
         Left            =   90
         TabIndex        =   45
         Top             =   2010
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Height          =   270
         Left            =   90
         TabIndex        =   44
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "DDO Code"
         Height          =   270
         Left            =   90
         TabIndex        =   43
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Title"
         Height          =   270
         Left            =   90
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance"
         Height          =   270
         Left            =   90
         TabIndex        =   38
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   270
         Left            =   90
         TabIndex        =   35
         Top             =   330
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Subsidiary Account Heads"
      ForeColor       =   &H000000C0&
      Height          =   6885
      Left            =   6240
      TabIndex        =   32
      Top             =   870
      Width           =   7485
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   6435
         Left            =   150
         TabIndex        =   20
         Top             =   360
         Width           =   7245
         _cx             =   12779
         _cy             =   11351
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCreateSubsidiaryAccountHeads.frx":1D36
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
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   7598
      TabIndex        =   18
      Top             =   7830
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&New"
      Height          =   375
      Left            =   4905
      TabIndex        =   17
      Top             =   7830
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6255
      TabIndex        =   16
      Top             =   7830
      Width           =   1335
   End
   Begin VB.ComboBox cmbSubLegerType 
      Height          =   390
      Left            =   4605
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   390
      Width           =   5445
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   13140
      Top             =   8340
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fmePersonal 
      Caption         =   "Personal Information"
      ForeColor       =   &H000000C0&
      Height          =   3495
      Left            =   90
      TabIndex        =   24
      Top             =   870
      Width           =   6045
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3660
         MaxLength       =   13
         TabIndex        =   9
         Top             =   2430
         Width           =   2295
      End
      Begin VB.TextBox txtDoorNo2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2370
         TabIndex        =   8
         Top             =   2850
         Width           =   525
      End
      Begin VB.TextBox txtDoorNo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   7
         Top             =   2850
         Width           =   795
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   5
         Top             =   2010
         Width           =   4545
      End
      Begin VB.TextBox txtStreet 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   4
         Top             =   1590
         Width           =   4545
      End
      Begin VB.TextBox txtWardNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2430
         Width           =   1485
      End
      Begin VB.TextBox txtLocalPlace 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         Top             =   1200
         Width           =   4545
      End
      Begin VB.TextBox txtHouseName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   2
         Top             =   780
         Width           =   4545
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Top             =   390
         Width           =   4545
      End
      Begin VB.Label lblMsgPersonal 
         AutoSize        =   -1  'True
         Caption         =   "Disabled for this SubLedger Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   3150
         TabIndex        =   42
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         Height          =   270
         Left            =   3060
         TabIndex        =   33
         Top             =   2460
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2250
         Y1              =   2850
         Y2              =   3150
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Door No"
         Height          =   270
         Left            =   120
         TabIndex        =   31
         Top             =   2910
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Main Place"
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   2070
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Street"
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   1650
         Width           =   525
      End
      Begin VB.Label lblWardNo 
         AutoSize        =   -1  'True
         Caption         =   "Ward No"
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   2490
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Local Place"
         Height          =   270
         Left            =   120
         TabIndex        =   27
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "House Name"
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   420
         Width           =   450
      End
   End
   Begin VB.Label lblCode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SubLedger Code"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   10140
      TabIndex        =   40
      Top             =   390
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SubLedger Type"
      Height          =   270
      Left            =   3015
      TabIndex        =   23
      Top             =   450
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Creating Subsidiary Account Heads"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmCreateSubsidiaryAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSubLedgerType As Boolean

    Private Sub cmbDepartment_Click()
        Dim mCommonSql As String
        If cmbDepartment.ListIndex > 0 Then
            'If mUserFlag = False Then
                PopulateList cmbDesignations, "Select Distinct chvDesigName,TB_EmployeeDetails_TRN.intDesigId [intDesignationID] from TB_EmployeeDetails_TRN Inner join TB_Designation_Lcl_Mst On TB_Designation_Lcl_Mst.intDesigId=TB_EmployeeDetails_TRN.intDesigId where intDeptId=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & " Order By chvDesigName", , True, True, True, enuSourceString.Sthapana
            'Else
            '    mCommonSql = "SELECT Distinct case when DB_Masters..GM_User.intDesignationID =0 then 'Temperary Staff'else chvDesigName end as chvDesigName,DB_Masters..GM_User.intDesignationID [intDesignationID] FROM DB_Masters.dbo.GM_User "
            '    mCommonSql = mCommonSql + "LEFT JOIN DB_Sthapana..TB_Designation_Lcl_Mst ON DB_Sthapana..TB_Designation_Lcl_Mst.intDesigId=DB_Masters.dbo.GM_User.intDesignationID "
            '    mCommonSql = mCommonSql + "INNER JOIN DB_Sthapana..TB_Department_Lcl_Mst ON DB_Sthapana..TB_Department_Lcl_Mst.intDeptId=DB_Masters.dbo.GM_User.intDeptID"
            '    mCommonSql = mCommonSql + " Where tnyActive = 0 And DB_Sthapana..TB_Department_Lcl_Mst.intDeptID = " & cmbDepartment.ItemData(cmbDepartment.ListIndex) & " Order By chvDesigName"
        
            '   PopulateList cmbDesignations, mCommonSql, , True, True, True, enuSourceString.DBMaster
            '    cmbDesignations.AddItem "Temperary Staff", 1
            'End If
            Call GetEmployees
        End If
    End Sub
    
    Private Sub cmbDesignations_Click()
        Call GetEmployees
    End Sub

    Private Sub cmbSubLegerType_Click()
        If cmbSubLegerType.ListIndex = -1 Then Exit Sub
        
        Call FormInitialize
        If mSubLedgerType = True Then
            Call FillGrid
        End If
        fmeCodeTitle.Enabled = True
        fmePersonal.Enabled = True
        cmdSave.Enabled = True
        cmdClear.Enabled = True
        fmeSulekha.Visible = False
        Select Case cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex)
       
            Case 1:
                If mSubLedgerType = True Then
                    fmeSulekha.Visible = True
                    fmePersonal.Enabled = True
                    lblMsgPersonal.Visible = False
                    Label15.Caption = "Interface from Sulekha Plan Monitoring"
                    
                    fmeOfficials.Visible = False
                    vsGrid2.Height = 6705
                    vsGrid2.Width = 8715
                    vsGrid2.Top = 510
                    Call GetImpementingOfficer
                End If
            Case 2:
                If mSubLedgerType = True Then
                    fmeSulekha.Visible = True
                    fmePersonal.Enabled = False
                    lblMsgPersonal.Visible = True
                    fmeCodeTitle.Enabled = True
                    Label15.Caption = "Interface from Sulekha Plan Monitoring"
                    
                    fmeOfficials.Visible = False
                    vsGrid2.Height = 6705
                    vsGrid2.Width = 8715
                    vsGrid2.Top = 510
                    Call GetImplementingAgencies
                End If
            Case 3:
                If mSubLedgerType = True Then
                    fmeSulekha.Visible = True
                    fmeCodeTitle.Enabled = True
                    fmePersonal.Enabled = False
                    lblMsgPersonal.Visible = True
                    Label15.Caption = "Interface from Sulekha Plan Monitoring"
                    
                    fmeOfficials.Visible = False
                    vsGrid2.Height = 6705
                    vsGrid2.Width = 8715
                    vsGrid2.Top = 510
                    Call GetAccreditedAgencies
                End If
            Case 4:
                If mSubLedgerType = True Then
                    fmeSulekha.Visible = True
                    fmeCodeTitle.Enabled = True
                    fmePersonal.Enabled = False
                    lblMsgPersonal.Visible = True
                    Label15.Caption = "Interface from Sulekha Plan Monitoring"
                    
                    fmeOfficials.Visible = False
                    vsGrid2.Height = 6705
                    vsGrid2.Width = 8715
                    vsGrid2.Top = 510
                    Call GetAuthorisedAgencies
                End If
            Case 6, 7, 8, 13:
                fmePersonal.Enabled = True
'                fmeCodeTitle.Enabled = True
                fmeCodeTitle.Visible = False
                lblMsgDept.Visible = False
                lblMsgPersonal.Visible = False
            Case 10: 'Sthapana
'                fmeSulekha.Visible = True
'                'fmePersonal.Enabled = False
'                'lblMsgPersonal.Visible = True
'                fmeCodeTitle.Enabled = True
'                Label15.Caption = "Interface from Sthapana Pay Bill Module"
'
'                fmeOfficials.Visible = True
'                vsGrid2.Height = 5565
'                vsGrid2.Width = 8715
'                vsGrid2.Top = 1680
'                Call GetEmployees
'                frmSearchEmplyees.Show vbModal
                fmeCodeTitle.Enabled = False
                fmePersonal.Enabled = True ' CHANGED BY AIBY 11-Feb-2015
                cmdSave.Enabled = True      ' CHANGED BY AIBY 11-Feb-2015
                'cmdClear.Enabled = False
                If mSubLedgerType = True Then
                    'frmSearchEmplyees.Show vbModal
                End If
            Case 12:
                fmePersonal.Enabled = False
                lblMsgPersonal.Visible = True
                fmeCodeTitle.Enabled = True
                lblMsgDept.Visible = False
                lblMsgPersonal.Visible = False
        End Select
        
        mSubLedgerType = True
    End Sub
    
'''    Private Sub HideFrame()
'''        fmePersonal.Visible = False
'''        fmeCodeTitle.Visible = False
'''        fmeSulekha.Visible = False
'''    End Sub
    
    Private Sub ShowTitle()
        lblTitle.Visible = True
        lblCode.Visible = True
        txtCode.Visible = True
        txtTitle.Visible = True
    End Sub

    Private Sub cmbSubLegerType_LostFocus()
'        If cmbSubLegerType.ListIndex > 0 Then
'            If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 10 Then 'Officials
'                If mSubLedgerType = True Then
'                    frmSearchEmplyees.Show vbModal
'                End If
'            End If
'        End If
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdClear_Click()
        Call FormInitialize
        If cmbSubLegerType.ListIndex > 0 Then
            If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 10 Then 'Officials
                'frmSearchEmplyees.Show vbModal
                Call FillGrid
            End If
        End If
    End Sub

    Private Sub cmdSave_Click()
        If SaveValidation Then
            If SaveSubLedger Then
                MsgBox "Subsidiary Account Head Saved Successfully", vbInformation
                Call FormInitialize
                Call FillGrid
            End If
        End If
    End Sub
    Private Sub Command1_Click()
        
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
            fmeSulekha.Left = 2468
            fmeSulekha.Top = 915
        WindowsXPC1.InitIDESubClassing
    End Sub

    Private Sub Form_Load()
        On Error GoTo Err:
            PopulateList cmbSubLegerType, "Select vchSubLedgerType,intSubLedgerTypeID From faSubLedgerTypes Where intSubLedgerTypeID in(1,2,3,4,6,7,8,10,12,13) ", , True, , True
            Call FillGrid
''            Call HideFrame
            lblMsgPersonal.Visible = False
            lblMsgDept.Visible = False
            mSubLedgerType = True
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Function SaveValidation() As Boolean
        On Error GoTo Err:
           
            If cmbSubLegerType.ListIndex = -1 Then
                MsgBox "Please Select the SubLedger Type", vbInformation
                cmbSubLegerType.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = (7 Or 8) Then
                If txtName.Text = "" Then
                    MsgBox "Please Enter the Name before Saving", vbInformation
                    txtName.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                       Added on 11/10/2011 By Poornima
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 6 Or cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 8 Then
                If Trim(txtName.Text) = "" Then
                    MsgBox "Please Enter the Name", vbInformation
                    txtName.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             '===================================================================================='
            '               Validation For linking to Subsidiary Cash Book                       '
            '===================================================================================='
            If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 10 Then     '   Officials'
                Dim objUser As New clsUser
                Dim mUserID As Variant
                
                mUserID = objUser.GetUserIDFromSthapanaEmpID(txtHouseName.Tag)
                If IsEmpty(mUserID) Then
                    MsgBox "This Official is not Present in Finance, Please Add the user and Continue the Process", vbInformation
                    SaveValidation = False
                    Exit Function
                Else
                    txtHouseName.Tag = mUserID
                End If
            End If
            '===================================================================================='
            Select Case cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex)
                Case 1:
                    If Trim(txtName.Text) = "" Then
                        MsgBox "Please Enter the Name!", vbInformation
                        txtName.SetFocus
                        SaveValidation = False
                        Exit Function
                    End If
                    If Trim(txtDDOCode.Text) = "" Then
                        MsgBox "Please Enter the DDO Code!", vbInformation
                        txtDDOCode.SetFocus
                        SaveValidation = False
                        Exit Function
                    End If
            End Select
            
            Select Case cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex)
                Case 1, 2, 3, 12:
                    If txtTitle.Text = "" Then
                        MsgBox "Please Enter Title", vbInformation
                        txtTitle.SetFocus
                        SaveValidation = False
                        Exit Function
                    End If
            End Select
            SaveValidation = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub FormInitialize()
        On Error GoTo Err:
            'cmbSubLegerType.ListIndex = -1
            txtDoorNo1.Text = ""
            txtDoorNo2.Text = ""
            txtHouseName.Text = ""
            txtLocalPlace.Text = ""
            txtMainPlace.Text = ""
            txtName.Text = ""
            txtStreet.Text = ""
            txtWardNo.Text = ""
            txtPhone.Text = ""
            txtName.Tag = ""
            txtTitle.Text = ""
            txtSubTitle.Text = ""
            txtOpeningBalance.Text = ""
            txtCode.Text = ""
            txtDDOCode.Text = ""
            txtDepartment.Text = ""
            txtDesignation.Text = ""
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Function SaveSubLedger() As Boolean
        On Error GoTo Err:
            Dim objDB As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim aryOut As Variant
            
            If objDB.SetConnection(mCnn) Then
                aryIn = Array(cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex), _
                    val(txtName.Tag), _
                     Trim(txtTitle.Text), _
                    Trim(txtSubTitle.Text), _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    val(txtOpeningBalance.Text), _
                    Null, _
                    txtHouseName.Tag, _
                    txtDDOCode.Text, _
                    Null, _
                    Trim(txtName.Text), _
                    Trim(txtHouseName), _
                    Trim(txtStreet), _
                    Trim(txtLocalPlace), _
                    Trim(txtMainPlace), _
                    Null, _
                    Null, _
                    Trim(txtPhone), _
                    val(txtWardNo.Text), _
                    val(Trim(txtDoorNo1)), _
                    Trim(txtDoorNo2), gbFinancialYearID, Trim(txtDesignation.Text), Trim(txtDepartment.Text), Null, Null, Null, Null, _
                    Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
                    
                objDB.ExecuteSP "spSaveSubSidiaryAccountHeads", aryIn, aryOut, , mCnn
                txtCode.Text = aryOut(0, 0)
            Else
                MsgBox "Connection to Finance doesnot Exist, Please contact your System Administrator", vbInformation
            End If
            SaveSubLedger = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Public Function FillGrid() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim mRowCnt As Integer
            Dim objDB As New clsDB
            
            If objDB.SetConnection(mCnn) Then
                mRowCnt = 1
                vsGrid.Rows = 2
                vsGrid.Clear 1, 1
                mSQL = "Select * from faSubSidiaryAccountHeads "
                mSQL = mSQL + " Inner Join faSubLedgerTypes On faSubSidiaryAccountHeads.intSubLedgerTypeID = faSubLedgerTypes.intSubLedgerTypeID "
                If cmbSubLegerType.ListIndex <> -1 Then
                    mSQL = mSQL + " Where faSubSidiaryAccountHeads.intSubLedgerTypeID = " & cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex)
                End If
                mSQL = mSQL + " Order By vchName "
                Rec.Open mSQL, mCnn
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    If Rec!intSubLedgerTypeID = 12 Then
                        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
                    End If
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSubLedgerType), "", Rec!vchSubLedgerType)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchHouseOrOffice), "", Rec!vchHouseOrOffice)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!numWardNo), "", Rec!numWardNo)
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                    vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                    vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!fltOpeningBalance), "", Rec!fltOpeningBalance)
                    vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!vchSubTitle), "", Rec!vchSubTitle)
                    If IsNull(Rec!vchName) Then
                        vsGrid.TextMatrix(mRowCnt, 13) = 1
                    Else
                        vsGrid.TextMatrix(mRowCnt, 13) = 2
                    End If
                    vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
                    vsGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!vchReferenceCode), "", Rec!vchReferenceCode)
                    vsGrid.TextMatrix(mRowCnt, 16) = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
                    vsGrid.TextMatrix(mRowCnt, 17) = IIf(IsNull(Rec!vchDepartment), "", Rec!vchDepartment)
                    vsGrid.TextMatrix(mRowCnt, 18) = IIf(IsNull(Rec!numEmpID), "", Rec!numEmpID)
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Finance doesnot Exist, Please contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function

    Private Sub lblExitSulekhaPanal_Click()
        fmeSulekha.Visible = False
    End Sub

    Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtOpeningBalance_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtOpeningBalance_LostFocus()
        txtOpeningBalance.Text = Format(val(txtOpeningBalance.Text), "0.00")
        If val(txtOpeningBalance.Text) < 0 Then
            txtOpeningBalance.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub txtPhone_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsGrid_Click()
        On Error GoTo Err:
            Dim mRowCnt As Integer
            
            Call FormInitialize
            
            If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
                mSubLedgerType = False
                cmbSubLegerType.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            Else
                cmbSubLegerType.ListIndex = -1
                Exit Sub
            End If
            
            vsGrid.Cell(flexcpBackColor, 1, 0, vsGrid.Rows - 1, 4) = vbWhite
            vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 4) = &HC0C0FF
            If vsGrid.TextMatrix(vsGrid.Row, 13) = 1 Then
                txtTitle.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            Else
                txtName.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            End If
            txtTitle.Text = vsGrid.TextMatrix(vsGrid.Row, 14)
            txtName.Tag = vsGrid.TextMatrix(vsGrid.Row, 0)
            txtCode.Text = txtName.Tag
            txtHouseName.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            txtLocalPlace.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
            txtStreet.Text = vsGrid.TextMatrix(vsGrid.Row, 5)
            txtMainPlace.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
            txtWardNo.Text = vsGrid.TextMatrix(vsGrid.Row, 7)
            txtDoorNo1.Text = vsGrid.TextMatrix(vsGrid.Row, 8)
            txtDoorNo2.Text = vsGrid.TextMatrix(vsGrid.Row, 9)
            txtPhone.Text = vsGrid.TextMatrix(vsGrid.Row, 10)
            'txtTitle.Text = txtName.Text
            txtSubTitle.Text = IIf(vsGrid.TextMatrix(vsGrid.Row, 12) = "", Trim(vsGrid.TextMatrix(vsGrid.Row, 0)), Trim(vsGrid.TextMatrix(vsGrid.Row, 12)))
            txtDDOCode.Text = vsGrid.TextMatrix(vsGrid.Row, 15)
            txtDepartment.Text = vsGrid.TextMatrix(vsGrid.Row, 17)
            txtDesignation.Text = vsGrid.TextMatrix(vsGrid.Row, 16)
            txtOpeningBalance.Text = val(vsGrid.TextMatrix(vsGrid.Row, 11))
            
            
            If fmeSulekha.Visible = True Then
                For mRowCnt = 1 To vsGrid2.Rows - 1
                    If vsGrid.TextMatrix(vsGrid.Row, 12) = vsGrid2.TextMatrix(mRowCnt, 2) Then
                        vsGrid2.Select mRowCnt, 0, mRowCnt, 2
                        vsGrid2.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 2) = &HC0C0FF
                    End If
                Next
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Function GetImplementingAgencies() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            Dim mRowCnt As Integer
            
            If objDB.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
                mSQL = "Select * from M_ImplAgency Where intImplAgencyTypeID <> 7 and intImplAgencyTypeID <> 6 Order By chvImplAgency"
                Rec.Open mSQL, mCnn
                vsGrid2.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid2.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intImplAgencyID), "", Rec!intImplAgencyID)
                    vsGrid2.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvEngImplAgency), "", Rec!chvImplAgency)
                    vsGrid2.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!chrImplAgencyCode), "", Rec!chrImplAgencyCode)
                    Rec.MoveNext
                    vsGrid2.Rows = vsGrid2.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetImplementingAgencies = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function GetAccreditedAgencies() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            Dim mRowCnt As Integer
            
            If objDB.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
                mSQL = "Select * from M_ImplAgency Where intImplAgencyTypeID = 7 Order By chvImplAgency"
                Rec.Open mSQL, mCnn
                vsGrid2.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid2.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intImplAgencyID), "", Rec!intImplAgencyID)
                    vsGrid2.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvEngImplAgency), "", Rec!chvImplAgency)
                    vsGrid2.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!chrImplAgencyCode), "", Rec!chrImplAgencyCode)
                    Rec.MoveNext
                    vsGrid2.Rows = vsGrid2.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetAccreditedAgencies = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function GetEmployees() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDB As New clsDB
            Dim mSQL As String
            Dim mRowCnt As Integer
            
            Dim mDepartment As Variant
            Dim mDesig As Variant
            
            If objDB.CreateNewConnection(mCnn, enuSourceString.Sthapana) Then
                
                If cmbDepartment.ListIndex = -1 Then
                    PopulateList cmbDepartment, "SELECT chvDeptName,intDeptID FROM TB_Department_Lcl_Mst Order By chvDeptName", , True, True, True, enuSourceString.Sthapana
                End If
                
                
                'mSQL = "Select chvEmpName,intEmpId,chvEmpId from TB_EmployeeDetails_Trn Order By chvEmpName"
                
                
                    
                If cmbDepartment.ListIndex < 1 Then
                    mDepartment = "%"
                Else
                    mDepartment = CStr(cmbDepartment.ItemData(cmbDepartment.ListIndex))
                End If
                
                If cmbDesignations.ListIndex < 1 Then
                    mDesig = "%"
                Else
                    mDesig = CStr(cmbDesignations.ItemData(cmbDesignations.ListIndex))
                End If
                    
                mSQL = "Select chvEmpName, intEmpId, chvEmpId from TB_EmployeeDetails_Trn where Convert(varchar(10),intDeptID) Like '" & mDepartment & "' And Convert(varchar(10),intDesigId) Like '" & mDesig & "'Order By Ltrim(chvEmpName)"
                
                Rec.Open mSQL, mCnn
                vsGrid2.Rows = 2
                mRowCnt = 1
                
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid2.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intEmpID), "", Rec!intEmpID)
                    vsGrid2.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvEmpName), "", Rec!chvEmpName)
                    vsGrid2.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!chvEmpID), "", Rec!chvEmpID)
                    Rec.MoveNext
                    vsGrid2.Rows = vsGrid2.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
                GetEmployees = True
            Else
                MsgBox "Connection To Sthapana does not exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function GetImpementingOfficer() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            Dim mRowCnt As Integer
            
            If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSQL = "Select * from suImplementingOfficer Where intLBTypeID = " & gbLBType & " Order By vchImplementingOfficer"
                Rec.Open mSQL, mCnn
                vsGrid2.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid2.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intImplementingOfficerID), "", Rec!intImplementingOfficerID)
                    vsGrid2.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchImplementingOfficer), "", Rec!vchImplementingOfficer)
                    vsGrid2.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchImplementingOfficerCode), "", Rec!vchImplementingOfficerCode)
                    Rec.MoveNext
                    vsGrid2.Rows = vsGrid2.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetImpementingOfficer = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function GetAuthorisedAgencies() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            Dim mRowCnt As Integer
            
            If objDB.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
                mSQL = "Select * from M_ImplAgency Where intImplAgencyTypeID = 6 order By chvImplAgency"
                Rec.Open mSQL, mCnn
                vsGrid2.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid2.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intImplAgencyID), "", Rec!intImplAgencyID)
                    vsGrid2.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvEngImplAgency), "", Rec!chvImplAgency)
                    vsGrid2.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!chrImplAgencyCode), "", Rec!chrImplAgencyCode)
                    Rec.MoveNext
                    vsGrid2.Rows = vsGrid2.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetAuthorisedAgencies = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function

    Private Sub vsGrid_DblClick()
        If cmbSubLegerType.ListIndex > 0 Then
            If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 10 Then 'Officials
                If vsGrid.TextMatrix(vsGrid.Row, 18) <> "" Then
                    frmEmployeeSubledger.EmployeeID = vsGrid.TextMatrix(vsGrid.Row, 18)
                    frmEmployeeSubledger.Show vbModal
                End If
            End If
        End If
    End Sub

    Private Sub vsGrid2_Click()
        On Error GoTo Err:
            Dim mRowCnt As Integer
        
            vsGrid2.Cell(flexcpBackColor, 1, 0, vsGrid2.Rows - 1, 2) = vbWhite
            vsGrid2.Cell(flexcpBackColor, vsGrid2.Row, 0, vsGrid2.Row, 2) = &HC0C0FF
    
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub vsGrid2_DblClick()
        On Error GoTo Err:
        
            Dim mRowCnt As Integer
            
            Select Case cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex)
                Case 1:
                    txtName.Text = "" 'vsGrid2.TextMatrix(vsGrid2.Row, 1)
                    txtTitle.Text = vsGrid2.TextMatrix(vsGrid2.Row, 1)
                Case 2:
                    txtTitle.Text = vsGrid2.TextMatrix(vsGrid2.Row, 1)
                Case 3:
                    txtTitle.Text = vsGrid2.TextMatrix(vsGrid2.Row, 1)
                Case 10:
                    txtName.Text = vsGrid2.TextMatrix(vsGrid2.Row, 1)
                    txtHouseName.Tag = vsGrid2.TextMatrix(vsGrid2.Row, 0)
            End Select
            txtSubTitle.Text = vsGrid2.TextMatrix(vsGrid2.Row, 2)
'''           --TO ALLOW MULTIPLE IMPO
'''            For mRowCnt = 1 To vsGrid.Rows - 1
'''                If vsGrid.TextMatrix(mRowCnt, 12) = Trim(txtSubTitle.Text) Then
'''                    txtOpeningBalance.Text = val(vsGrid.TextMatrix(mRowCnt, 11))
'''                    txtTitle.Text = Trim(vsGrid.TextMatrix(mRowCnt, 12))
'''                    txtName.Tag = val(vsGrid.TextMatrix(mRowCnt, 0))
'''                End If
'''            Next
            
            fmeSulekha.Visible = False
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
