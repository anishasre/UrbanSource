VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchBuildingDetails 
   BackColor       =   &H00C7F6EF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDemand 
      BackColor       =   &H00C7F6EF&
      Caption         =   "Search in Demand List"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9540
      TabIndex        =   14
      Top             =   690
      Width           =   2040
   End
   Begin VB.CheckBox chkPreviousOwner 
      BackColor       =   &H00C7F6EF&
      Caption         =   "Previous Owner"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9540
      TabIndex        =   13
      Top             =   390
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ComboBox cmbWard 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   555
      Width           =   2580
   End
   Begin VB.TextBox txtDoorNo1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5985
      TabIndex        =   9
      Top             =   555
      Width           =   1485
   End
   Begin VB.TextBox txtDoorNo2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7485
      TabIndex        =   8
      Top             =   555
      Width           =   1080
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7485
      TabIndex        =   7
      Top             =   960
      Width           =   1065
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2370
      TabIndex        =   6
      Top             =   975
      Width           =   5115
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4620
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   11820
      _cx             =   20849
      _cy             =   8149
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
      BackColorFixed  =   13100007
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchBuildingDetails.frx":0000
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
   Begin VB.ComboBox cmbZone 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5985
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2580
   End
   Begin VB.ComboBox cmbAssessmentYear 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2580
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3390
      Top             =   8610
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ward"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1905
      TabIndex        =   12
      Top             =   615
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Door No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5265
      TabIndex        =   11
      Top             =   615
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Owner"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1035
      TabIndex        =   5
      Top             =   1065
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zone"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   2
      Top             =   195
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assessment Year"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   0
      Top             =   210
      Width           =   1365
   End
End
Attribute VB_Name = "frmSearchBuildingDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    
    Option Explicit
    Private Sub FormInitialize()
        On Error Resume Next
        cmbAssessmentYear.ListIndex = 0
        cmbZone.ListIndex = 0
        cmbWard.ListIndex = -1
        chkDemand.Value = 1
        If cmbWard.ListCount > 0 Then
            cmbWard.ListIndex = 0
        End If
        txtDoorNo1.Text = ""
        txtDoorNo2.Text = ""
        txtName.Text = ""
        vsGrid.Clear 0, 0
        vsGrid.Cell(flexcpText, 0, 0) = "Ward"
        vsGrid.Cell(flexcpText, 0, 1) = "Door No"
        vsGrid.Cell(flexcpText, 0, 2) = "Name of Owner"
        'On Error GoTo 0
    End Sub
    
    Private Sub SearchBuilding()
        Dim mSQL As String
        Dim mWhere As String
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        Dim mBuildingID         As Variant
        Dim mZoneID             As Variant
        Dim mAssessmentYearID   As Variant
        Dim mWardNo             As Variant
        Dim mDoorNo1            As Variant
        Dim mDoorNo2 As Variant
        Dim mName As Variant
        Dim mResHName As Variant
        
        mBuildingID = Null
        If cmbZone.ListIndex > -1 Then
            mZoneID = cmbZone.ItemData(cmbZone.ListIndex)
        End If
        If cmbWard.ListIndex > -1 Then
            mWardNo = cmbWard.ItemData(cmbWard.ListIndex)
        Else
            MsgBox "Please select the Ward !", vbInformation
            Exit Sub
        End If
        If Len(Trim(txtDoorNo1)) Then
            mDoorNo1 = Val(txtDoorNo1)
        End If
        If Len(Trim(txtDoorNo2)) Then
            mDoorNo2 = Val(txtDoorNo2)
        End If
        If Len(Trim(txtName.Text)) Then
            mName = Trim(txtName)
        End If
        mResHName = Null
        
        'mSQL = mSQL + " SELECT snGeneralAssessment.chvOwners, "
        'mSQL = mSQL + "        snGeneralAssessment.chvHouseName, "
        'mSQL = mSQL + "        snGeneralAssessment.chvResidenceAssNo, "
        'mSQL = mSQL + "        snGeneralAssessment.chvLocalPlace, "
        'mSQL = mSQL + "        snGeneralAssessment.fltPTax1+ snGeneralAssessment.fltPTax2 as AnnualPTax, "
        'mSQL = mSQL + "        snGeneralAssessment.fltLC1 + snGeneralAssessment.fltLC2 as AnnualLC, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.numBuildingId, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.intAssessmentYear, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.numZoneId, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.numWardId, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.intDoorNo1, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.chvDoorNo2, "
        'mSQL = mSQL + "        snGeneralAssessmentDoorNo.intWardNo, "
        'mSQL = mSQL + "        snGeneralAssessment.fltPTax1, "
        'mSQL = mSQL + "        snGeneralAssessment.fltPTax2, "
        'mSQL = mSQL + "        snGeneralAssessment.fltLC1, "
        'mSQL = mSQL + "        snGeneralAssessment.fltLC2 "
        'mSQL = mSQL + " FROM   snGeneralAssessmentDoorNo INNER JOIN "
        'mSQL = mSQL + "        snGeneralAssessment ON snGeneralAssessmentDoorNo.numBuildingId = snGeneralAssessment.numBuildingID"
        'mSQL = mSQL + " Where  "
        
        mSQL = " SELECT DISTINCT snGeneralAssessmentDoorNo.intWardNo, "
        mSQL = mSQL + " IsNull(LTRIM(Str(snGeneralAssessmentDoorNo.intDoorNo1) +'/'+  snGeneralAssessmentDoorNo.chvDoorNo2),LTRIM(Str( snGeneralAssessmentDoorNo.intDoorNo1))),"
        mSQL = mSQL + " snGeneralAssessment.chvOwners, "
        If Month(gbTransactionDate) - 4 < 6 Then
            mSQL = mSQL + " snGeneralAssessment.fltPTax1, snGeneralAssessment.fltLC1, snGeneralAssessment.fltPTax1 + snGeneralAssessment.fltLC1 As Total1, "
        Else
            mSQL = mSQL + " snGeneralAssessment.fltPTax2, snGeneralAssessment.fltLC2, snGeneralAssessment.fltPTax2 + snGeneralAssessment.fltLC2 As Total2, "
        End If
        mSQL = mSQL + " snGeneralAssessmentDoorNo.numBuildingId "
        mSQL = mSQL + " FROM snGeneralAssessmentDoorNo INNER JOIN"
        mSQL = mSQL + " snGeneralAssessment ON snGeneralAssessmentDoorNo.numBuildingId = snGeneralAssessment.numBuildingID"
        
        If chkDemand.Value = 1 Then
            mSQL = mSQL + " Inner Join snDemandTbl ON snDemandTbl.numBuildingID = snGeneralAssessment.numBuildingID "
        Else
                
        End If
        mSQL = mSQL + " Where "
       
        If Not (mBuildingID = Null) Then
            mSQL = mSQL + " snGeneralAssessmentDoorNo.numBuildingID = & mBuildingID  AND"
        End If
        If Not mZoneID = Null Then
            mSQL = mSQL + " snGeneralAssessmentDoorNo.numZoneID = " & mZoneID & " AND'"
        End If
        If Not mAssessmentYearID = Null Then
            mSQL = mSQL + " snGeneralAssessmentDoorNo.intAssessmentYear = " & mAssessmentYearID & " AND "
        End If
        If Not IsNull(mWardNo) Then
            mSQL = mSQL + " snGeneralAssessmentDoorNo.intWardNo = " & mWardNo & " AND "
        Else
            MsgBox vbCrLf & "Please select a ward" & vbCrLf, vbInformation
            cmbWard.SetFocus
            Exit Sub
        End If
        If Not mDoorNo1 = Null Then
            mSQL = mSQL + " snGeneralAssessmentDoorNo.intDoorNo1 = " & mDoorNo1 & " AND "
        End If
        If Not mDoorNo2 = Null Then
            mSQL = mSQL + " snGeneralAssessmentDoorNo.chvDoorNo2 LIKE '" & mDoorNo2 & "' And "
        End If
        If Not mName = Null Then
            mSQL = mSQL + " snGeneralAssessment.chvOwners LIKE '" & mName & "%' AND"
        End If
        If Not mResHName = Null Then
            mSQL = mSQL + " snGeneralAssessment.chvHouseName LIKE '" & mResHName & "%' AND"
        End If
        mSQL = Left(mSQL, Len(mSQL) - 4)
        
        Debug.Print mSQL
        
        
        Me.MousePointer = vbHourglass
        cmdSearch.Enabled = False
        vsGrid.Rows = 1
        objDb.CreateNewConnection mCnn, enuSourceString.SanchayaLite
        Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        If Not (Rec.BOF And Rec.EOF) Then
            vsGrid.Rows = Rec.RecordCount + 1
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColSel = 6
            vsGrid.RowSel = vsGrid.Rows - 1
            mSQL = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = mSQL
        End If
        Rec.Close
        
        If vsGrid.Rows > 1 Then
            vsGrid.Row = 1
            vsGrid.Col = 0
        End If
        Me.MousePointer = vbDefault
        cmdSearch.Enabled = True
        
        'mSQL = "        Select RIGHT(Cast( numWardID As varChar(20)), 2), "
        'mSQL = mSQL + " Cast(intDoorNo1 As varChar(10)) +' ' + Isnull('/' + chvDoorNo2,''), "
        'mSQL = mSQL + " snGeneralAssessmentAddress.chvSPFullName , "
        'mSQL = mSQL + " fltPTax1, fltLC1 ,(fltPTax1 + fltLC1) Total ,snGeneralAssessment.numBuildingID "
        'mSQL = mSQL + " FROM Db_SanchayaLite.dbo.snGeneralAssessment snGeneralAssessment "
        '
        'If chkDemand.Value = 1 Then
        'mSQL = mSQL + " INNER JOIN snDemandTbl snDemandTbl ON snDemandTbl.numBuildingID = snGeneralAssessment.numBuildingID "
        'End If
        '
        'mSQL = mSQL + " LEFT JOIN Db_SanchayaLite.dbo.snGeneralAssessmentDoorNo snGeneralAssessmentDoorNo ON snGeneralAssessment.numBuildingID = snGeneralAssessmentDoorNo.numBuildingId INNER JOIN"
        'mSQL = mSQL + "     Db_SanchayaLite.dbo.snGeneralAssessmentOwner snGeneralAssessmentOwner ON snGeneralAssessment.numBuildingID = snGeneralAssessmentOwner.numBuildingID LEFT JOIN"
        'mSQL = mSQL + "     Db_SanchayaLite.dbo.snGeneralAssessmentAddress snGeneralAssessmentAddress ON snGeneralAssessmentAddress.numAddressId = snGeneralAssessmentOwner.numAddressID AND snGeneralAssessmentOwner.sintOwnerSlNo = 1"
        '
        'mWhere = " Where "
        'If cmbWard.ListIndex > -1 Then
        '    mWhere = mWhere + " numWardID = " & cmbWard.ItemData(cmbWard.ListIndex)
        '    If Trim(txtDoorNo1) <> "" Then
        '        If Len(mWhere) > 7 Then mWhere = mWhere + " And"
        '        mWhere = mWhere + " intDoorNo1 = " & Val(txtDoorNo1)
        '        If Trim(txtDoorNo2) <> "" Then
        '            mWhere = mWhere + " And chvDoorNo2 Like '" & Trim(txtDoorNo2) & "'"
        '        End If
        '    End If
        '    If Trim(txtName) <> "" Then
        '        If Len(mWhere) <> 7 Then mWhere = mWhere + " And "
        '        mWhere = mWhere + " chvSPFullName Like '%" & Trim(txtName) & "%'"
        '    End If
        '    mSQL = mSQL + mWhere
        '    Me.MousePointer = vbHourglass
        '
        '    cmdSearch.Enabled = False
        '    vsGrid.Rows = 1
        '    objDB.CreateNewConnection mCnn, enuSourceString.SanchayaLite
        '    Rec.CursorLocation = adUseClient
        '    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        '    If Not (Rec.BOF And Rec.EOF) Then
        '        vsGrid.Rows = Rec.RecordCount + 1
        '        vsGrid.Col = 0
        '        vsGrid.Row = 1
        '        vsGrid.ColSel = 6
        '        vsGrid.RowSel = vsGrid.Rows - 1
        '        mSQL = Rec.GetString(, , vbTab, Chr(13))
        '        vsGrid.Clip = mSQL
        '    End If
        '    If vsGrid.Rows > 1 Then
        '        vsGrid.Row = 1
        '        vsGrid.Col = 0
        '    End If
        '    Me.MousePointer = vbDefault
        '    cmdSearch.Enabled = True
        'Else
        '    MsgBox vbCrLf & "Please select a ward" & vbCrLf, vbInformation
        '    cmbWard.SetFocus
        'End If
        
        
        
    End Sub
    Private Sub FillWard()
        Dim mSQL As String
        If cmbAssessmentYear.ListIndex > -1 Then
            mSQL = "SELECT chvWardNameEnglish, intWardNo, numWardID FROM GM_Ward "
            mSQL = mSQL + " WHERE tnyWardType = 1 AND intLBID = " & gbLocalBodyID & " AND numZoneID = " & cmbZone.ItemData(cmbZone.ListIndex)
            mSQL = mSQL + " AND intWardYear = " & cmbAssessmentYear.ItemData(cmbAssessmentYear.ListIndex)
            mSQL = mSQL + " Order By intWardNo "
            PopulateList cmbWard, mSQL, , , , True, enuSourceString.DBMaster
        End If
    End Sub
    Private Sub cmbZone_click()
        Call FillWard
    End Sub
    Private Sub cmdSearch_Click()
        Call SearchBuilding
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbEnter Then
            Call SearchBuilding
        End If
    End Sub
    
    Private Sub Form_Load()
        Dim mSQL As String
        vsGrid.Cell(flexcpText, 0, 0) = "Ward"
        vsGrid.Cell(flexcpText, 0, 1) = "Door No"
        vsGrid.Cell(flexcpText, 0, 2) = "Owner"
        vsGrid.Cell(flexcpText, 0, 3) = "Tax"
        vsGrid.Cell(flexcpText, 0, 4) = "L.C"
        vsGrid.Cell(flexcpText, 0, 5) = "Total"
        
        
        mSQL = "SELECT intAssessmentID, intAssessmentID From snMstAssessment"
        PopulateList cmbAssessmentYear, mSQL, , , , True, enuSourceString.SanchayaLite
        
        Call FormInitialize
        mSQL = "SELECT  chvZoneNameEnglish  , numZoneID  From GM_Zone  Where intLBId = " & gbLocalBodyID & " Order By chvZoneNameEnglish"
        PopulateList cmbZone, mSQL, "Main Office", True, , True, enuSourceString.DBMaster
        vsGrid.SelectionMode = flexSelectionByRow
        vsGrid.AutoSearch = flexSearchFromTop
        
    End Sub
    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            Call vsGrid_KeyDown(13, 0)
        End If
    End Sub
    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then
            If vsGrid.Row > 0 Then
                If vsGrid.TextMatrix(vsGrid.Row, 6) <> "" Then
                   frmPropertyTax.BuildingID = Val(vsGrid.TextMatrix(vsGrid.Row, 6))
                   Unload Me
                   frmPropertyTax.Visible = True
                   frmPropertyTax.ZOrder (0)
                End If
            End If
        End If
    End Sub
