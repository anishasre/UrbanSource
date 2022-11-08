VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchRentOnLandBuildings 
   BackColor       =   &H00FFFFF7&
   Caption         =   "Search Rent on Land & Building"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   ClipControls    =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearFilters 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Clear Filters"
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
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1755
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1755
   End
   Begin VB.PictureBox WindowsXPC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   -3630
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   6555
      Width           =   1200
   End
   Begin VB.Frame framParty 
      BackColor       =   &H00FFFFF7&
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
      Height          =   2565
      Left            =   210
      TabIndex        =   4
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtRoomNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1590
         TabIndex        =   20
         Top             =   1800
         Width           =   870
      End
      Begin VB.TextBox txtShopName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4950
         TabIndex        =   18
         Top             =   510
         Width           =   3465
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4950
         TabIndex        =   16
         Top             =   1365
         Width           =   3465
      End
      Begin VB.TextBox txtLocalPlace 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4950
         TabIndex        =   14
         Top             =   930
         Width           =   3465
      End
      Begin VB.TextBox txtDoorNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1590
         TabIndex        =   12
         Top             =   1395
         Width           =   870
      End
      Begin VB.TextBox txtWardNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1590
         TabIndex        =   9
         Top             =   990
         Width           =   870
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7095
         TabIndex        =   8
         Top             =   1875
         Width           =   1335
      End
      Begin VB.ComboBox cmbAssetType 
         Height          =   360
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   1740
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Room No"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   870
         TabIndex        =   19
         Top             =   1860
         Width           =   630
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Shop Name"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4110
         TabIndex        =   17
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lblMainPlace 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4125
         TabIndex        =   15
         Top             =   1425
         Width           =   840
      End
      Begin VB.Label lblLocalPlace 
         BackStyle       =   0  'Transparent
         Caption         =   "Lolcal Place"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4065
         TabIndex        =   13
         Top             =   990
         Width           =   900
      End
      Begin VB.Label lblDoorNo 
         BackStyle       =   0  'Transparent
         Caption         =   "DoorNo"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   960
         TabIndex        =   11
         Top             =   1470
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rend on Land && Buildings"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   270
         Left            =   3210
         TabIndex        =   0
         Top             =   150
         Width           =   2370
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1110
         TabIndex        =   3
         Top             =   1065
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Type"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   735
         TabIndex        =   1
         Top             =   585
         Width           =   810
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4110
      Left            =   165
      TabIndex        =   10
      Top             =   2700
      Width           =   9060
      _cx             =   15981
      _cy             =   7250
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
      BackColorFixed  =   -2147483638
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
      Rows            =   14
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchRentOnLandBuildings.frx":0000
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
Attribute VB_Name = "frmSearchRentOnLandBuildings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearFilters_Click()
    Call FormInitialize
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub FormInitialize()
    cmbAssetType.ListIndex = -1
    txtDoorNo.Text = ""
    txtLocalPlace.Text = ""
    txtMainPlace.Text = ""
    txtShopName.Text = ""
    txtWardNo.Text = ""
    txtRoomNo.Text = ""
    vsGrid.Clear 1, 1
End Sub
Private Sub cmdSearch_Click()
    Dim objDb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    Dim mRowCount As Integer
    Dim mAssetType As Variant
    
    Dim mLicenseFee As Variant
    Dim mServiceCharge As Variant
    
    If cmbAssetType.ListIndex = -1 Then
        mAssetType = "%"
    Else
        mAssetType = cmbAssetType.ItemData(cmbAssetType.ListIndex)
    End If
    
        objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        
    mSQL = " SELECT     Right(snRentAssetDetails.numWardID,2) as WardNO, snRentAssetDetails.chvAssetName, snRentAssetDetails.intCategory, snRentAssetDetails.chvLocation, "
    mSQL = mSQL + "  snRentAssetDetails.chvMainPlace, snRentAssetDetails.chvRemarks, snRentDeedDetails.fltLicenseFee, snRentDeedDetails.fltSurcharge, "
    mSQL = mSQL + " snRentDeedDetails.tnyStatus, snRentDeedDetails.chvShopName, snRentDeedDetails.numDeedRegNo, snRentAssetDetails.numRegNo, "
    mSQL = mSQL + " snRentAssetSubItems.numSubItemsID, snRentAssetSubItems.chvDoorNo, snRentAssetSubItems.chvSubItems, snRentAssetSubItems.chvDescription,  snRentAssetSubItems.intStatus as OccupyStatus, "
    mSQL = mSQL + " smAddressBook.vchName , smAddressBook.vchPlace, smAddressBook.intDistrictID, smAddressBook.vchPin, smAddressBook.vchPhone "
    mSQL = mSQL + " FROM         snRentAssetDetails INNER JOIN "
    mSQL = mSQL + " snRentAssetSubItems ON snRentAssetDetails.numRegNo = snRentAssetSubItems.numRegNo INNER JOIN "
    mSQL = mSQL + " snRentDeedDetails ON snRentAssetSubItems.numSubItemsID = snRentDeedDetails.chvSubUnit "
    mSQL = mSQL + " INNER JOIN snRentOwnersDetails "
    mSQL = mSQL + " ON snRentOwnersDetails.numDeedRegNo = snRentDeedDetails.numDeedRegNo "
    mSQL = mSQL + " INNER JOIN smAddressBook "
    mSQL = mSQL + " ON snRentOwnersDetails.numAddressId = smAddressBook.intAddressID   WHERE  "

    If cmbAssetType.ListIndex <> -1 Then
        mSQL = mSQL + " snRentAssetDetails.intCategory = " & cmbAssetType.ItemData(cmbAssetType.ListIndex)
    Else
        mSQL = mSQL + " snRentAssetDetails.intCategory > '0' "
    End If
    
    If txtLocalPlace.Text <> "" Then
        mSQL = mSQL + " AND snRentAssetDetails.chvLocation LIKE '%" & txtLocalPlace.Text & "%'"
    End If
    
    If txtMainPlace.Text <> "" Then
        mSQL = mSQL + " AND snRentAssetDetails.chvMainPlace LIKE '%" & txtMainPlace.Text & "%'"
    End If
    
    If txtWardNo.Text <> "" Then
        mSQL = mSQL + " AND Right(snRentAssetDetails.numWardID,2) = " & Val(txtWardNo.Text)
    End If
    
    If txtDoorNo.Text <> "" Then
        mSQL = mSQL + " AND snRentAssetSubItems.chvDoorNo LIKE ' " & txtDoorNo.Text & "%'"
    End If
    
    If txtShopName.Text <> "" Then
        mSQL = mSQL + " AND snRentDeedDetails.chvShopName LIKE '%" & txtShopName.Text & "%'"
    End If
    
    If txtRoomNo.Text <> "" Then
        mSQL = mSQL + " AND snRentAssetSubItems.chvSubItems LIKE '%" & txtRoomNo.Text & "%'"
    End If
    
    Rec.Open mSQL, mCnn
    
    vsGrid.Rows = 2
    mRowCount = 1
    If Not Rec.EOF Or Not Rec.BOF Then
        While Not (Rec.EOF Or Rec.BOF)
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!WardNO), "", Rec!WardNO)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!chvShopName), "", Rec!chvShopName)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            mLicenseFee = IIf(IsNull(Rec!fltLicenseFee), "", Rec!fltLicenseFee)
            mServiceCharge = IIf(IsNull(Rec!fltSurcharge), "", Rec!fltSurcharge)
            vsGrid.TextMatrix(mRowCount, 4) = mLicenseFee + mServiceCharge
            Rec.MoveNext
            mRowCount = mRowCount + 1
            vsGrid.Rows = vsGrid.Rows + 1
        Wend
    Else
        vsGrid.Clear 1, 1
    End If
End Sub

Private Sub Form_Load()
    Dim mCnn As New ADODB.Connection
    Dim objDb As New clsDB
    Dim mSQL As String
    Call FormInitialize
    objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    mSQL = "Select chvCategoryDetails,intCategoryID from snMstCategoryRent"
    PopulateList cmbAssetType, mSQL, , True, , True, enuSourceString.iSaankhyaMasters
End Sub

Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
If Not (KeyAscii <= Asc("9") Or KeyAscii >= Asc("0")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsGrid_DblClick()
    Dim Rec As New ADODB.Recordset
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    Dim mRowCount As Integer
    Dim mGet As String
    Dim mCurrent As Double
    Dim mArrear As Double
    objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        
        mCurrent = 0
        mArrear = 0
        mSQL = "set dateformat dmy SELECT * FROM snRentDeedDetails INNER JOIN smRentDemand ON snRentDeedDetails.numDeedRegNo = smRentDemand.numDeedRegNo"
        mSQL = mSQL + " INNER JOIN smRentDemandChild ON smRentDemand.intDemandID = smRentDemandChild.intDemandID"
         mSQL = mSQL + " INNER JOIN DB_Finance..faAccountHeads A on smRentDemandChild.intAccountHeadID=A.intAccountHeadID "
        mSQL = mSQL + " Where snRentDeedDetails.chvShopName= '" & vsGrid.Text & "'  And smRentDemandChild.tnyStatus = 0 And dtDemandDate < '" & Date & "'"
        
        frmRentOnLandBuildings.Show
        frmSearchRentOnLandBuildings.Visible = False
        mRowCount = 1
        Rec.Open mSQL, mCnn
        
    While Not (Rec.EOF)
        frmRentOnLandBuildings.txtShopName.Text = Rec!chvShopName
        frmRentOnLandBuildings.vsGrid.TextMatrix(mRowCount, 0) = Rec!vchAccountHeadCode
        frmRentOnLandBuildings.vsGrid.TextMatrix(mRowCount, 1) = Rec!vchAccountHead
        frmRentOnLandBuildings.vsGrid.TextMatrix(mRowCount, 2) = Rec!dtDemandDate
        frmRentOnLandBuildings.vsGrid.TextMatrix(mRowCount, 5) = Rec!fltTotalAmount
        frmRentOnLandBuildings.vsGrid.Cell(flexcpChecked, mRowCount, 12) = 1
          mArrear = mArrear + CDbl(Val(frmRentOnLandBuildings.vsGrid.TextMatrix(mRowCount, 4)))
        mCurrent = mCurrent + CDbl(Val(frmRentOnLandBuildings.vsGrid.TextMatrix(mRowCount, 5)))
        
        frmRentOnLandBuildings.vsGrid.Rows = frmRentOnLandBuildings.vsGrid.Rows + 1
        
        mRowCount = mRowCount + 1
        Rec.MoveNext
    Wend
        frmRentOnLandBuildings.lblTotalCurrent.Caption = mCurrent
        frmRentOnLandBuildings.lblTotalArrear.Caption = mArrear
        Rec.Close
    Set mCnn = Nothing
        objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        mGet = "Select Right (numWardID,2) From snRentDeedDetails"
        'Rec.Open mGet, mCnn
        PopulateList frmRentOnLandBuildings.cmbWard, mGet, , True, , , enuSourceString.iSaankhyaMasters
        'Rec.Close
End Sub
