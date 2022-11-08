VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmProfessionTaxSearch 
   BackColor       =   &H80000018&
   Caption         =   "Profession Tax Search"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   8070
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2265
      Left            =   60
      TabIndex        =   8
      Top             =   1680
      Width           =   7905
      _cx             =   13944
      _cy             =   3995
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmProfessionTaxSearch.frx":0000
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
   Begin VB.TextBox txtAddress 
      Height          =   1575
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmProfessionTaxSearch.frx":01FA
      Top             =   30
      Width           =   4125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   1665
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   3795
      Begin VB.TextBox txtInstitution 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1020
         Width           =   2355
      End
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2355
      End
      Begin VB.ComboBox cmbMainPlace 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Institution"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Location"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Main Place"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmProfessionTaxSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmbLocation_click()
    FillGrid cmbMainPlace.Text, cmbLocation.Text, txtInstitution.Text
End Sub

Private Sub cmbMainPlace_Click()
    FillGrid cmbMainPlace.Text, cmbLocation.Text, txtInstitution.Text
End Sub

Private Sub Form_Load()
    txtInstitution.Text = ""
    txtAddress.Text = ""
    Me.ZOrder (0)
    FillGrid
    FillCombos
End Sub
Private Sub FillGrid(Optional mMainPlace As String, Optional mLocation As String, Optional mInstitutionName As String)
    vsGrid.Rows = 1
    Dim str As String
    Dim mCon As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    str = " "
    str = str & _
            "SELECT     Traders.intInstID," & _
            " chvInstitutionName," & _
            " Traders.intRevenueWard," & _
            " Traders.intDoorNo," & _
            " Traders.chvDoorNo," & _
            " Traders.chvLocalPlace," & _
            " Traders.chvMainPlace," & _
            " Traders.chvPostOffice," & _
            " Traders.intPincode," & _
            " Traders.intApplicableRate," & _
            " Traders.fltHalfYearIncome," & _
            " Traders.fltTaxRate," & _
            " Traders.chvPayeeName," & _
            " Traders.intSerialNo , Traders.chvMultiDoorNo" & _
            " FROM  DB_SaankhyaMasters..TB_ProfessionTaxInstitution_MST" & _
            " Traders INNER JOIN DB_SaankhyaMasters..TB_WardinLB_MST Ward" & _
            " ON Traders.intRevenueWard=Ward.intWardID   "
    If (Not (mMainPlace = "" And mLocation = "" And mInstitutionName = "")) Then
        str = str & " Where"
    End If
    If Not mMainPlace = "" Then
        str = str & " chvMainPlace='" & mMainPlace & "'  AND"
    End If
    
    If Not mLocation = "" Then
        str = str & " chvLocalPlace='" & mLocation & "' AND"
    End If
    If Not mInstitutionName = "" Then
        str = str & " chvInstitutionName like '" & mInstitutionName & "%' AND"
    End If
    str = mID(str, 1, Len(str) - 3)
    If objDB.SetExtDBConnection(mCon, objDB.GetConnectionString(enuSourceString.SaankhyaMasters)) Then
        Set Rec = objDB.ExecuteSP(str, , , , mCon, adCmdText)
        If Not Rec.EOF Then
            Me.vsGrid.LoadArray Rec.GetRows
        End If
    End If
    
End Sub
Private Sub FillCombos()
    PopulateList cmbMainPlace, "SELECT distinct chvMainPlace FROM TB_ProfessionTaxInstitution_MST order by chvMainPlace", , True, True, False, SaankhyaMasters
    PopulateList cmbLocation, "SELECT distinct chvLocalPlace  FROM TB_ProfessionTaxInstitution_MST order by chvLocalPlace", , True, True, False, SaankhyaMasters
End Sub


Private Sub txtInstitution_LostFocus()
    FillGrid cmbMainPlace.Text, cmbLocation.Text, txtInstitution.Text
End Sub

Private Sub vsGrid_Click()
    If vsGrid.Row > 0 Then
        txtAddress.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
        txtAddress.Text = txtAddress.Text & ", " & vsGrid.TextMatrix(vsGrid.Row, 5)
        txtAddress.Text = txtAddress.Text & ", " & vsGrid.TextMatrix(vsGrid.Row, 6)
        txtAddress.Text = txtAddress.Text & ", " & vsGrid.TextMatrix(vsGrid.Row, 7)
        txtAddress.Text = txtAddress.Text & ", " & vsGrid.TextMatrix(vsGrid.Row, 8)
    End If
End Sub

Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            frmProfessionalTax.numSubLedgerID = Val(vsGrid.TextMatrix(vsGrid.Row, 0))
            frmProfessionalTax.txtInstitution.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            frmProfessionalTax.txtLocation.Text = vsGrid.TextMatrix(vsGrid.Row, 5)
            frmProfessionalTax.txtMainPlace.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
            frmProfessionalTax.txtAddress.Text = txtAddress.Text
            frmProfessionalTax.FillGrid Val(vsGrid.TextMatrix(vsGrid.Row, 0))
            Unload Me
        End If
End Sub
