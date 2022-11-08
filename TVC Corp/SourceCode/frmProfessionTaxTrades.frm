VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmPofessionTaxTrades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmPofessionTaxTrades"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
   LinkTopic       =   "frmProfessionTaxTrades"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopyToReceipt 
      Cancel          =   -1  'True
      Caption         =   "Copy To Receipt"
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
      Left            =   4020
      TabIndex        =   27
      Top             =   5910
      Width           =   1875
   End
   Begin VSFlex8LCtl.VSFlexGrid vsProfTaxDetails 
      Height          =   2565
      Left            =   30
      TabIndex        =   9
      Top             =   2940
      Width           =   10440
      _cx             =   18415
      _cy             =   4524
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
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProfessionTaxTrades.frx":0000
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6300
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   5970
      TabIndex        =   11
      Top             =   5910
      Width           =   1545
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
      Height          =   2400
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11760
      Begin VB.TextBox txtCategory 
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
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1650
         Width           =   2610
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
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1320
         Width           =   1080
      End
      Begin VB.TextBox txtInstNo 
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
         Left            =   1470
         TabIndex        =   24
         Top             =   2010
         Width           =   2715
      End
      Begin VB.ComboBox cmbWardYear 
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
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   23
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
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   960
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
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   20
         Top             =   975
         Width           =   600
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
         Left            =   2610
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1320
         Width           =   840
      End
      Begin VB.TextBox txtInstName 
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
         TabIndex        =   4
         Top             =   705
         Width           =   5085
      End
      Begin VB.TextBox txtAuthName 
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
         TabIndex        =   3
         Top             =   1065
         Width           =   5085
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
         TabIndex        =   2
         Top             =   600
         Width           =   3210
      End
      Begin VB.CommandButton cmdMaster 
         Caption         =   "..."
         Height          =   345
         Left            =   4200
         TabIndex        =   1
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Index           =   2
         Left            =   660
         TabIndex        =   28
         Top             =   1680
         Width           =   795
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
         Index           =   0
         Left            =   930
         TabIndex        =   22
         Top             =   1035
         Width           =   450
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
         Index           =   1
         Left            =   750
         TabIndex        =   19
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Left            =   1020
         TabIndex        =   17
         Top             =   300
         Width           =   390
      End
      Begin VB.Label lblShopName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Name"
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
         Left            =   4920
         TabIndex        =   8
         Top             =   735
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorised Person"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   1080
         Width           =   1365
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
         TabIndex        =   6
         Top             =   645
         Width           =   480
      End
      Begin VB.Label lblDemandname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institution No:"
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
         Left            =   30
         TabIndex        =   5
         Top             =   2040
         Width           =   1395
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2745
      Left            =   0
      TabIndex        =   26
      Top             =   2880
      Width           =   11610
      _cx             =   20479
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
      Rows            =   1
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProfessionTaxTrades.frx":00C0
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
   End
   Begin VB.Label lblFine 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10440
      TabIndex        =   31
      Top             =   6000
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fine"
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
      Left            =   9930
      TabIndex        =   30
      Top             =   6090
      Width           =   420
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
      Left            =   8640
      TabIndex        =   16
      Top             =   5745
      Width           =   375
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
      Left            =   9270
      TabIndex        =   15
      Top             =   6360
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
      Left            =   10440
      TabIndex        =   14
      Top             =   5700
      Width           =   1260
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
      Left            =   9150
      TabIndex        =   13
      Top             =   5700
      Width           =   1260
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Label2"
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   2490
      Visible         =   0   'False
      Width           =   11715
   End
End
Attribute VB_Name = "frmPofessionTaxTrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public mPTType As Integer  '''1-- traders ,2  Employees
    Dim dtUptoDate  As Date     ' Fine Upto Date
    Dim dtFromDate  As Date


    Private Sub cmbWard_Click()
        If cmbWard.ListIndex > -1 Then
            txtWardNo.Text = cmbWard.ItemData(cmbWard.ListIndex)
        End If
    End Sub
    Private Sub FillZone()
        Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Where intLBID = " & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster)
    End Sub
    Private Sub FillCat()
        'Call PopulateList(cmbCat, " Select 'Employees' cat ,2 catType  Union All Select 'Traders' cat , 1 catType ", "Traders", True, True, True, DBMaster)
    End Sub
    
    Private Sub FillWard()
        Dim mSql As String
        On Error Resume Next
        mSql = "SELECT chvWardNameEnglish,  intWardNo ,numWardID FROM GM_Ward"
        mSql = mSql + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID
        'mSQL = mSQL + " AND  intAsessmentYearID = " & cmbWardYear.ItemData(cmbWardYear.ListIndex)
        mSql = mSql + " AND numZoneID = " & cmbZone.ItemData(cmbZone.ListIndex)
        mSql = mSql + " Order By intWardNo ,chvWardNameEnglish"
        PopulateList cmbWard, mSql, , , , True, enuSourceString.DBMaster
    End Sub
    Private Function GetWardID(ByVal mWardNo As Integer) As Double
        Dim mSql As String
        Dim mCnn  As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim objdb As New clsDB
        mSql = "SELECT chvWardNameEnglish,  intWardNo ,numWardID FROM GM_Ward"
        mSql = mSql + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID
        mSql = mSql + " AND numZoneID = " & cmbZone.ItemData(cmbZone.ListIndex) & "And  intWardNo=" & mWardNo
        mSql = mSql + " Order By intWardNo ,chvWardNameEnglish"
        objdb.CreateNewConnection mCnn, enuSourceString.DBMaster
        Rec.Open mSql, mCnn
                If Not (Rec.BOF Or Rec.EOF) Then
                   GetWardID = IIf(IsNull(Rec!numWardId), 0, Rec!numWardId)
                Else
                   GetWardID = 0
                End If
     
    End Function
    Private Sub cmbZone_Change()
    Call FillWard
    End Sub

    Private Sub cmbZone_click()
    Call FillWard
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdCopyToReceipt_Click()
        Dim objAcc  As New clsAccounts
        Dim mLoop As Integer
        Dim mCount As Integer
        Dim mLoopChild As Integer
        Dim mInstID As Variant

        mInstID = txtInstNo.Text
        If mInstID > 0 Then
            frmReceiptsCounter.SubLedgerID = mInstID
            frmReceiptsCounter.cmbZone.Text = cmbZone.Text
            frmReceiptsCounter.cmbZone.Locked = True
            frmReceiptsCounter.txtWard.Text = cmbWard.Text
            frmReceiptsCounter.txtWard.Locked = True

            frmReceiptsCounter.txtWardNo.Text = val(txtWardNo)
            frmReceiptsCounter.txtBuildingNo = mInstID

            frmReceiptsCounter.txtDoorNo1.Text = Trim(txtHouseNo1.Text)
            frmReceiptsCounter.txtDoorNo2.Text = Trim(txtHouseNo2.Text)
            frmReceiptsCounter.txtName.Text = txtAuthName.Text
            frmReceiptsCounter.DemandBasedFlag = True
        End If

        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 0) <> "" Then
            frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 0) = vsGrid.TextMatrix(mLoop, 0)
                If mPTType = 1 Then
                     If (vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxTradersCurrent) Then
                    
                         frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 5) = vsGrid.TextMatrix(mLoop, 5)
                     ElseIf vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxTradersArrears Then
                         
                         frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 4) = vsGrid.TextMatrix(mLoop, 4)
                     End If
                ElseIf mPTType = 2 Then
                    If (vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxEmployees) Then
'
                         frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 5) = vsGrid.TextMatrix(mLoop, 5)
'                     ElseIf vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxTradersArrears Then
'
''                         frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 4) = vsGrid.TextMatrix(mLoop, 4)
                     End If
                End If
                frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 2) = vsGrid.TextMatrix(mLoop, 2)
                frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 3) = vsGrid.TextMatrix(mLoop, 8)
                frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 6) = vsGrid.TextMatrix(mLoop, 6)
                frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 7) = vsGrid.TextMatrix(mLoop, 7)
                frmReceiptsCounter.vsGrid.Cell(flexcpText, mLoop, 8) = CInt(vsGrid.Cell(flexcpText, mLoop, 8))
                frmReceiptsCounter.vsGrid.Cell(flexcpText, mLoop, 9) = vsGrid.TextMatrix(mLoop, 9)
                frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 10) = vsGrid.TextMatrix(mLoop, 10)
                frmReceiptsCounter.vsGrid.TextMatrix(mLoop, 11) = vsGrid.TextMatrix(mLoop, 11)
            End If
        Next

        frmReceiptsCounter.txtHouse.Text = txtInstName.Text
        frmReceiptsCounter.SubLedgerID = mInstID
        frmReceiptsCounter.Calculate
        frmReceiptsCounter.txtBuildingNo.Text = mInstID
        frmReceiptsCounter.txtDoorNo2.MaxLength = 15
        frmReceiptsCounter.txtDoorNo2.Locked = True
'        frmReceiptsCounter.txtStreet.Text = txtSubItemName.Text
        frmReceiptsCounter.txtDoorNo2.Tag = txtHouseNo2.Tag
        frmReceiptsCounter.vsGrid.Editable = flexEDNone
        frmReceiptsCounter.txtName.Enabled = False

        Unload Me
    End Sub

Private Sub cmdMaster_Click()
        Dim client              As New MSSOAPLib.SoapClient
        Dim objdb               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New Recordset
        Dim mRec                As New Recordset
        Dim objSOAP             As Variant
        Dim mUrl                As String
        Dim mArrOutChild        As String
        Dim mArrOutProf       As String
        Dim mDemand             As Variant
        Dim mArrIn              As Variant
        Dim mSql                As String
        Dim mxmlvalue          As Variant
        Dim mStatus                   As String
        Dim mLoop               As Integer
        Dim intWardYear   As Integer
        Dim numZoneID           As Double
        Dim intWardNo           As Double
        Dim mInstName           As Variant
        Dim intDoorNo1          As String
        Dim chvDoorNo2          As String
        Dim mCredencial As String
        Dim mOwner As String
        Dim mInstID As Variant
        Dim mAuthName As String
        Dim mCnt As Integer
        Dim mInstType As Integer
        vsGrid.Visible = False
        vsProfTaxDetails.Visible = True
        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        mUrl = gbDefaultUrl
        objSOAP.MSSoapInit (mUrl + "?WSDL")
        If err.Number = -2147352567 Then
                MsgBox "Please check"
        End If
        
        mLoop = 0
        mCredencial = "ikm@revenue@sanchaya"
        'If Trim(txtInstNo) <> "" Then mInstID = Trim(txtInstNo) Else mInstID = ""
        'mPTType   1 Traders  2-Employes
            vsProfTaxDetails.Clear 1, 0
            vsGrid.Clear 1, 0
            mInstType = mPTType
            
            If txtInstNo.Text <> "" Then
                mInstID = Trim(txtInstNo)
                mArrOutProf = (objSOAP.getprofessionsearchinst(gbLocalBodyID, mInstID, mCredencial))
                'mArrOutProf = (objSOAP.getprofessionsearchinst(gbLocalBodyID, mInstID, mInstType, mCredencial))
            Else
                
                If cmbZone.ListIndex > -1 Then numZoneID = cmbZone.ItemData(cmbZone.ListIndex) Else numZoneID = Null
        
                If cmbWardYear.ListIndex > -1 Then intWardYear = cmbWardYear.ItemData(cmbWardYear.ListIndex) Else intWardYear = Null
                If cmbWard.ListIndex > -1 Then
                    intWardNo = cmbWard.ItemData(cmbWard.ListIndex)  'GetWardID(cmbWard.ItemData(cmbWard.ListIndex))
                Else
                    MsgBox "Please Select Ward", vbApplicationModal
                    intWardNo = 0
                    Exit Sub
                End If
                If Trim(txtInstName) <> "" Then
                    mInstName = Trim(txtInstName)
                Else
                    'If mInstType = 1 Then
                        MsgBox "Please Enter Shop name", vbApplicationModal
                        mInstName = ""
                        Exit Sub
            
                    'End If
                End If
                If Trim(txtHouseNo1) <> "" Then intDoorNo1 = val(txtHouseNo1) Else intDoorNo1 = 0
                If Trim(txtHouseNo2) <> "" Then chvDoorNo2 = Trim(txtHouseNo2) Else chvDoorNo2 = ""
                
                If Trim(txtAuthName) <> "" Then mAuthName = Trim(txtAuthName) Else mAuthName = ""
                
                mInstID = "0"
        
             'mArrIN = Array(167, 2000, 4016701, 1016701005, 28, "0", "t", "a", mInstID, "ikm@revenue@sanchaya")
               ' mArrOutProf = (objSOAP.getprofessionsearch(gbLocalBodyID, intWardYear, numZoneID, intWardNo, intDoorNo1, chvDoorNo2, mInstName, mAuthName, mInstID, mInstType, mCredencial))
                 mArrOutProf = (objSOAP.getprofessionsearch(gbLocalBodyID, intWardYear, numZoneID, intWardNo, intDoorNo1, chvDoorNo2, mInstName, mAuthName, mInstID, mCredencial))
            End If
            
            mCnt = 0
            If mArrOutProf = "null" Or mArrOutProf = "" Then
                MsgBox "Institution details not Exists "
                
                Exit Sub
            End If
            mxmlvalue = convertJsonToVariantArray(mArrOutProf)
            mLoop = 0
            For mLoop = 0 To UBound(mxmlvalue)
                If mxmlvalue(mLoop, 1) <> "" Then
                mCnt = mCnt + 1
                vsProfTaxDetails.Rows = mCnt + 1
                    vsProfTaxDetails.TextMatrix(mCnt, 0) = mxmlvalue(mLoop, 0)  ''
                    vsProfTaxDetails.TextMatrix(mCnt, 1) = mxmlvalue(mLoop, 1)
                    vsProfTaxDetails.TextMatrix(mCnt, 2) = mxmlvalue(mLoop, 2) & "/" & mxmlvalue(mLoop, 3)
                    vsProfTaxDetails.TextMatrix(mCnt, 3) = mxmlvalue(mLoop, 4)
                    vsProfTaxDetails.TextMatrix(mCnt, 4) = mxmlvalue(mLoop, 5)
                End If
            Next

            

End Sub
Private Sub Initialize()
    txtHouseNo1.Text = ""
    txtHouseNo2.Text = ""
End Sub
Private Function convertJsonToVariantArray(ByRef jsonString As String) As Variant()
        Dim cleanedUpArray() As Variant
        Dim brokenUpRows As Variant
        Dim mLoop1 As Integer
        If jsonString <> "" Then
            'Remove the first and last square bracket in the string
            jsonString = Right$(jsonString, Len(jsonString) - 2)
            jsonString = Left$(jsonString, Len(jsonString) - 2)
            
            'Break up the string in an array
            brokenUpRows = Split(jsonString, "},{")
            Dim Counter As Integer
            Counter = 0
            Dim counter2 As Long
            Dim brokenUpCols As Variant
            Dim strDomain As Variant
            Dim arr As Variant
            Dim str(300, 300) As Variant
            Dim counter3 As Integer
            'Dim Cnt As Integer
            'Dim counter As Integer
            
            Counter = 0
            counter3 = 0
            ReDim linkArray(UBound(brokenUpRows)) As String
            
            For Counter = 0 To UBound(brokenUpRows)
                brokenUpCols = Split(brokenUpRows(Counter), ",")
                If Counter = 0 Then
                    ReDim cleanedUpArray(UBound(brokenUpRows), UBound(brokenUpCols)) As Variant
                End If
                For counter2 = 0 To UBound(brokenUpCols)
                    cleanedUpArray(Counter, counter2) = brokenUpCols(counter2)
                    'syalima
                    arr = Split(cleanedUpArray(Counter, counter2), ":")
    
                    If UBound(arr) > 0 Then
                        str(Counter, counter2) = Trim(Replace(arr(1), """", " "))
                    End If
                    'syalima
                Next
            Next
            convertJsonToVariantArray = str
        End If
    End Function

    Private Sub Form_Load()
        Call FillAssessmentYear
        Call FillZone
        cmbZone.Text = gbLocation
        Call FillWard
        Call FillCat
        If mPTType = 1 Then
            txtCategory.Text = "Traders"
            txtCategory.Tag = 1
        ElseIf mPTType = 2 Then
            txtCategory.Text = "Employees"
            txtCategory.Tag = 2
        End If
    End Sub
    Private Sub FillAssessmentYear()
        Dim mSql As String
        On Error Resume Next
        mSql = "SELECT DISTINCT intWardYear,intWardYear as ID From GM_Ward Where tnyWardType = 1 AND intLBID = " & gbLocalBodyID & " ORDER BY intWardYear DESC"
        Call PopulateList(cmbWardYear, mSql, , , , True, DBMaster)
        cmbWardYear.ListIndex = 0
    End Sub


    Private Sub txtHouseNo1_KeyPress(KeyAscii As Integer)
          If KeyAscii = 13 Then
                KeyAscii = 0
                PressTabKey
            End If
    End Sub
    
    Private Sub txtInstNo_KeyPress(KeyAscii As Integer)
          If KeyAscii = 13 Then
                KeyAscii = 0
                PressTabKey
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
                KeyAscii = 0
                PressTabKey
            End If
    End Sub
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
            'If vsGrid.Cell(flexcpChecked, mCount, 12) = vbChecked Then
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
            'End If
       Next
       lblTotalArrear.Caption = Format(mArrearAmt, "0.00")
       lblTotalCurrent.Caption = Format(mCurrentAmt, "0.00")
        
        If lblFine.Caption = "" Then
            mGrantTot = Format(mArrearAmt + mCurrentAmt, "0.00")
        Else
            mGrantTot = Format(mArrearAmt + mCurrentAmt + lblFine.Caption, "0.00")
        End If
        'txtGrandTotal.Text = Format(mGrantTot, "0.00")
        'lblGrandTotal.Caption = Format(mGrantTot, "0.00")
       
'        If mAdv > 0 Then
'            txtAdvance.Text = Format(mAdv, "0.00")
'            mNetAmt = mGrantTot - mAdv
'        Else
            'mNetAmt = txtGrandTotal.Text
        'End If
        'If mAdv < 0 Then
          '  cmdCopyToReceipt.Enabled = False
        'End If
        txtGrandTotal = Format(mGrantTot, "0.00")
    End Sub
    
Private Sub vsProfTaxDetails_DblClick()

    Dim client              As New MSSOAPLib.SoapClient
        Dim objdb               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim objSOAP             As Variant
        Dim mUrl                As String
        Dim mArrOutChild        As String
        Dim mArrOutProfDemand       As String
        Dim mDemand             As Variant
        Dim mArrIn              As Variant
        Dim mSql                As String
        Dim mxmlvalue          As Variant
        Dim mStatus                   As String
        Dim mLicenceDemandChild As Variant
        Dim mLoop               As Integer
        Dim intWardYear   As Integer
        Dim numZoneID           As Double
        Dim intWardNo           As Double
        Dim mInstName           As Variant
        Dim intDoorNo1          As String
        Dim chvDoorNo2          As String
        Dim mCredencial As String
        Dim mOwner As String
        Dim mInstID As Variant
        Dim mArrOutProfDemandChild As String
        Dim objAcc      As New clsAccounts
        Dim mInstType As Integer
        vsProfTaxDetails.Visible = False
        vsGrid.Visible = True
        Dim mAuthName As String
        'Set objSOAP = CreateObject("MSSOAP.SoapSerializer30")
         Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        mUrl = gbDefaultUrl
        objSOAP.MSSoapInit (mUrl + "?WSDL")
        If err.Number = -2147352567 Then
                MsgBox "Please check"
        End If
        mLoop = 0
        mCredencial = "ikm@revenue@sanchaya"
        mInstType = mPTType
       
        If vsProfTaxDetails.Row > 0 Then
            mInstID = vsProfTaxDetails.TextMatrix(vsProfTaxDetails.Row, 0) ' Hidden Field Institution ID
            If val(mInstID) > 0 Then
                'mArrOutProfDemand = (objSOAP.getprofessiondemand(mInstID, gbLocalBodyID, mInstType, mCredencial))
                mArrOutProfDemand = (objSOAP.getprofessiondemand(mInstID, gbLocalBodyID, mCredencial))
            End If
            Dim mCnt As Integer
            mCnt = 0
            mxmlvalue = convertJsonToVariantArray(mArrOutProfDemand)
            mLoop = 0
            For mLoop = 0 To UBound(mxmlvalue)
                If mxmlvalue(mLoop, 1) <> "" Then
                mCnt = mCnt + 1
               txtInstNo.Text = mxmlvalue(mLoop, 0)
               'cmbZone.Text = cmbZone.ItemData(cmbZone.ListIndex)
               'cmbZone.Te= (mxmlvalue(mLoop, 1))
               'cmbZone.ListIndex = 6 'mxmlvalue(mLoop, 1)
               'cmbWard.ListIndex = mxmlvalue(mLoop, 2)
               'cmbWardYear.ListIndex = mxmlvalue(mLoop, 3)
               txtWardNo.Text = mxmlvalue(mLoop, 4)
               txtInstName.Text = mxmlvalue(mLoop, 5)
               txtAuthName.Text = mxmlvalue(mLoop, 6)
               txtHouseNo1.Text = mxmlvalue(mLoop, 7)
               txtHouseNo2.Text = mxmlvalue(mLoop, 8)
                    
                    'vsProfTaxDetails.TextMatrix(mLoop + 1, 5) = mxmlvalue(mLoop, 6)
                   ' vsProfTaxDetails.Rows = vsProfTaxDetails.Rows + 1
                End If
            Next
            'mArrOutProfDemandChild = (objSOAP.getprofessiondemandchild(mInstID, gbLocalBodyID, mInstType, mCredencial, gbLBType))
            mArrOutProfDemandChild = (objSOAP.getprofessiondemandchild(mInstID, gbLocalBodyID, mCredencial, gbLBType))
            If mArrOutProfDemandChild = "" Then
                MsgBox "Demand details does not exists"
                Exit Sub
            End If
            'ReDim mCnt As Integer
            vsGrid.Clear 1, 0
            mCnt = 0
            mxmlvalue = convertJsonToVariantArray(mArrOutProfDemandChild)
            mLoop = 0
            For mLoop = 0 To UBound(mxmlvalue)
                If mxmlvalue(mLoop, 1) <> "" Then
                mCnt = mCnt + 1
                'vsGrid.Rows = mCnt + 1
                
                
'                vsGrid.Rows = vsGrid.Rows + 1
                        With vsGrid
                            .Rows = .Rows + 1
                             objAcc.SetAccounts (IIf(IsNull(mxmlvalue(mLoop, 10)), -1, mxmlvalue(mLoop, 10)))
'                             If (objAcc.AccountCode = gbAcHeadCodepro) Then
'                                MsgBox "Advance Head Exists with Sanchaya Demand", vbApplicationModal
'                                Exit Sub
'                             End If
                            .TextMatrix(mCnt, 0) = objAcc.AccountCode
                            .TextMatrix(mCnt, 1) = objAcc.AccountHead
                            .TextMatrix(mCnt, 2) = mxmlvalue(mLoop, 3) + "-" + CStr((CInt(mxmlvalue(mLoop, 3)) + 1))
                            Select Case mxmlvalue(mLoop, 4)
                                Case Is = 1: .Cell(flexcpText, mCnt, 3) = "Ist Half"
                                Case Is = 2: .Cell(flexcpText, mCnt, 3) = "IInd Half"
                                Case Is = 3: .Cell(flexcpText, mCnt, 3) = "Full Year"
                            End Select
                           
'                            If mInstType = 1 Then
                                If (objAcc.AccountCode = gbAcHeadCodeProfTaxTradersCurrent) Then
                                    .TextMatrix(mCnt, 5) = mxmlvalue(mLoop, 5)
                                     .TextMatrix(mCnt, 11) = mxmlvalue(mLoop, 5)
                                End If
                                If (objAcc.AccountCode = gbAcHeadCodeProfTaxTradersArrears) Then
                                    .TextMatrix(mCnt, 4) = mxmlvalue(mLoop, 5)
                                     .TextMatrix(mCnt, 11) = mxmlvalue(mLoop, 5)
                                End If
'                            ElseIf (mInstType = 2) Then
'                                If (objAcc.AccountCode = gbAcHeadCodeProfTaxEmployees) Then
'                                    .TextMatrix(mCnt, 5) = mxmlvalue(mLoop, 5)
'                                     .TextMatrix(mCnt, 11) = mxmlvalue(mLoop, 5)
'                                End If
'                            End If
                            
'                            If mInstType = 1 Then
'
'                            End If
                         
                            .TextMatrix(mCnt, 6) = objAcc.AccountHeadID
                            .TextMatrix(mCnt, 7) = mxmlvalue(mLoop, 3)
                            .Cell(flexcpText, mCnt, 8) = CInt(mxmlvalue(mLoop, 4))
                            .TextMatrix(mCnt, 9) = mxmlvalue(mLoop, 4)
                            .TextMatrix(mCnt, 10) = mxmlvalue(mLoop, 0)
                           .Cell(flexcpChecked, mCnt, 12) = vbChecked
'                            If mxmlvalue(mLoop, 4) - 10 < 4 Then
'                            .Cell(flexcpText, mCnt, 2) = str(mxmlvalue(mLoop, 3) - 1) & " - " & str(mxmlvalue(mLoop, 3))
'                            Else
'                            .Cell(flexcpText, mCnt, 2) = str(mxmlvalue(mLoop, 3) & " - " & str(mxmlvalue(mLoop, 3) + 1))
'                            End If
'                            .Cell(flexcpText, mCnt, 3) = IIf(IsNull(mxmlvalue(mLoop, 4)), "", mxmlvalue(mLoop, 4))
'                            If Rec!ArrearFlag = 0 Then
'                                .TextMatrix(mCnt, 4) = Format(IIf(IsNull(mxmlvalue(mLoop, 5)), "", mxmlvalue(mLoop, 5)), "0.00")
'                            Else
'                                .TextMatrix(mCnt, 4) = IIf(IsNull(mxmlvalue(mLoop, 5)), "", mxmlvalue(mLoop, 5))
'                            End If
'                            .TextMatrix(mCnt, 6) = objAcc.AccountHeadID
'
'                            .Cell(flexcpText, mCnt, 8) = IIf(IsNull(Rec!chvPeriodID), "", Rec!chvPeriodID)
'                            .Cell(flexcpText, mCnt, 9) = IIf(IsNull(Rec!ArrearFlag), "", Rec!ArrearFlag)
'                            .Cell(flexcpText, mCnt, 11) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'                            .Cell(flexcpChecked, mCnt, 12) = vbChecked
'                            Call Calculate
'                            .Cell(flexcpText, mCnt, 10) = IIf(IsNull(Rec!intKeyID), "", Rec!intKeyID) 'txtDeedRegNo.Tag 'Rec!intKeyID 'Rec!numDemandID
'                            If .TextMatrix(mCnt, 0) = gbAcHeadCodeAdvanceBuilding Or .TextMatrix(mRows, 0) = gbAcHeadCodeAdvanceLand Then
'                                .Cell(flexcpText, mCnt, 14) = 1  'To identify Advance
'                            End If
'                            .Cell(flexcpText, mCnt, 15) = mDueDay
                            CalculateFineforProf
                        End With
                   Call Calculate
                End If
            Next
       End If
        
    
End Sub
    Private Function CalculateFineforProf() As Double
          Dim mLoop As Integer
        Dim mLoopCrl As Integer ' Act as a Static Variable
        Dim mFineAmt As Double  ' Total Fine Amount
        Dim mPTax    As Double  ' Total Arrear Property Tax
        Dim mLC      As Double
        Dim mCess    As Double
        Dim mPartAmt As Double  ' Total Ptax+LC+Cess after adjusting Advance Amount
'        Dim dtUptoDate As Date  ' Fine Upto Date
        Dim dtDemandDate As Date
        Dim mFine    As Double
'        Dim dtFromDate  As Date
'        mAdvAmt = 0
        dtUptoDate = gbTransactionDate
'        mAnyAdvanceFlag = True
        mLoopCrl = 1
'        mAdvCheckedRow = 0
        mPartAmt = 0
        mFineAmt = 0
        
        For mLoop = mLoopCrl To vsGrid.Rows - 1
               
'               'Note:- Geting Advance if any :: Seting dtUptoDate
'
'                If mAnyAdvanceFlag Then
'                    If mAdvAmt <= 0 Then
'                        Call GetAdvanceAmt
'                    End If
'                End If
'
               'Note: Not Arrear Property Tax Or Row is not selecte
               '      in both this case it skips the loop body
                If ((vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxEmployees Or _
                     vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxTradersArrears Or _
                     vsGrid.TextMatrix(mLoop, 0) = gbAcHeadCodeProfTaxTradersCurrent _
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
                 
                 If dtUptoDate >= dtDemandDate Then ' [1] dtUptoDate > dtDemandDate
                        
'                        'Note:- Finding Property Tax/LC/Cess
                         mPTax = val(vsGrid.TextMatrix(mLoop, 11))
'                         'Note:- In Next two rows expecting LC and Cess
'                         If vsGrid.Rows - 1 >= mLoop + 1 Then
'                         If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeLibraryCess Or vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodePoorHomeCess Then
'                             If vsGrid.TextMatrix(mLoop + 1, 0) = gbAcHeadCodeLibraryCess Then
'                                 mLC = val(vsGrid.TextMatrix(mLoop + 1, 11))
'                             Else
'                                 mCess = val(vsGrid.TextMatrix(mLoop + 1, 11))
'                             End If
'                         End If
'                         End If
'
'                         'Note:- Cess
'                         If vsGrid.Rows - 1 >= mLoop + 2 Then
'                         If vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodeLibraryCess Or vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodePoorHomeCess Then
'                             If vsGrid.TextMatrix(mLoop + 2, 0) = gbAcHeadCodeLibraryCess Then
'                                 mLC = val(vsGrid.TextMatrix(mLoop + 2, 11))
'                             Else
'                                 mCess = val(vsGrid.TextMatrix(mLoop + 2, 11))
'                             End If
'                         End If
'                         End If
                        'Note:- End of Block: Finding Property Tax/LC/Cess
'FindNextAdvance:
                         'Note:- Find If any Advance And Set new dtUptoDate
'                         If mAdvAmt <= 0 Then
'                             If mAnyAdvanceFlag Then Call GetAdvanceAmt
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
                 
'                        'Note:- Setting of Advance
'                         While (mFine > 0 And mAdvAmt > 0)
'                             If mFine <= mAdvAmt Then
'                                 mAdvAmt = mAdvAmt - mFine
'                                 mFine = 0
'                             Else
'                                 mFine = mFine - mAdvAmt
'                                 mAdvAmt = 0
'                                 If mAnyAdvanceFlag Then Call GetAdvanceAmt '::: Gets Any other Advance Exists also sets UptoDate
'                             End If
'                         Wend
                         
'                         If mAdvAmt > 0 Then ' [IF-3]
'                             If (mPTax + mLC + mCess) <= mAdvAmt Then
'                                 'No need to calculate Fine
'                                 mAdvAmt = mAdvAmt - (mPTax + mLC + mCess)
'                                 mPTax = 0
'                                 mLC = 0
'                                 mCess = 0
'                             Else '(mPTax + mLC + mCess) > mAdvAmt
'                                 mPartAmt = 0
'                                 If mCess > 0 Then
'                                     'NOTE:- Not completed!!! - Aiby/Dated:16-Sep-2009
'                                     'This part should be changed further to seperate to find Cess and LC
'                                     mPartAmt = (mPTax + mLC) - mAdvAmt
'                                     mPTax = Format(mPartAmt * 100 / 105, "0.00")
'                                     mLC = mPartAmt - mPTax
'                                 Else
'                                     mPartAmt = (mPTax + mLC) - mAdvAmt
'                                     mPTax = Format(mPartAmt * 100 / 105, "0.00")
'                                     mLC = mPartAmt - mPTax
'                                 End If
'                                 mAdvAmt = 0
'                                 If mPartAmt > 0 Then
'                                    'dtFromDate = dtUptoDate
'                                    GoTo FindNextAdvance:
'                                End If
'                             End If
'                         End If ' [IF-3]
                 
GoNext:
        Next
        lblFine.Caption = mFineAmt
        CalculateFineforProf = mFineAmt
        
    End Function
  
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
'
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
