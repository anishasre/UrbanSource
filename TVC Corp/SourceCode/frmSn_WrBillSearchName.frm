VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSn_WrBillSearchName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "frmSn_WrBillSearchName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8LCtl.VSFlexGrid fgSearch 
      Height          =   4545
      Left            =   15
      TabIndex        =   0
      Top             =   345
      Width           =   5400
      _cx             =   9525
      _cy             =   8017
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSn_WrBillSearchName.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
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
   Begin VB.Label lblHeading 
      BackColor       =   &H00C0C0FF&
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
      TabIndex        =   1
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "frmSn_WrBillSearchName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim mCnn As New ADODB.Connection
    '*********************************************************************************************'
    '              Common form to search Water Bill Offices, Caretakers etc                       '
    '*********************************************************************************************'
    Private Sub fgSearch_DblClick()
        If fgSearch.row > 0 Then
            If val(fgSearch.TextMatrix(fgSearch.row, 1)) <> 0 Then
                Call fgSearch_KeyPress(13)
            End If
        End If
    End Sub

    Private Sub fgSearch_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
    End Sub
    
    Private Sub fgSearch_KeyPress(KeyAscii As Integer)
        If fgSearch.row > 0 Then
            If KeyAscii = 13 Then
                If fgSearch.Rows <> 0 Then
                    Select Case intWrBillSearchID
                        Case 1
                            frmSn_WrBillConnections.txtOffice = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                            frmSn_WrBillConnections.txtOffice.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                        Case 2
                            frmSn_WrBillConnections.txtcareTakersName.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                            frmSn_WrBillConnections.txtcareTakersName = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                            frmSn_WrBillConnections.txtDesignation = fgSearch.Cell(flexcpText, fgSearch.row, 3)
                        Case 3
                            frmSn_WrBillConnectionList.txtOfficeInst = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                            frmSn_WrBillConnectionList.txtOfficeInst.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                        Case 4
                            frmSn_WrBillConnectionList.txtCaretaker.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                            frmSn_WrBillConnectionList.txtCaretaker = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                        Case 5
                            gbSearchID = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                            gbSearchStr = fgSearch.Cell(flexcpText, fgSearch.row, 2)
'                            frmSn_WrBillListOfTransactionDetails.txtCaretaker.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
'                            frmSn_WrBillListOfTransactionDetails.txtCaretaker.Text = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                        Case 6
                            frmSn_WrBillDetails.txtCaretaker.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                            frmSn_WrBillDetails.txtCaretaker.Text = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                        Case 7
                            frmSn_WrBillDetails.txtOfficeInst.Text = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                            frmSn_WrBillDetails.txtOfficeInst.Tag = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                        Case Else
                            gbSearchID = fgSearch.Cell(flexcpText, fgSearch.row, 1)
                            gbSearchStr = fgSearch.Cell(flexcpText, fgSearch.row, 2)
                    End Select
                End If
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        Dim i       As Integer
        Dim objDb   As New clsDB
        
        'CenterForm Me
        fgSearch.Clear 1
        fgSearch.Rows = 2
        'Set conSanchaya = gFunSetConnection(Dsn.Sanchaya)
        objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        
        Select Case intWrBillSearchID
            Case 1
                PopulateOffices
            Case 2
                PopulateCaretakers
            Case 3
                PopulateOffices
            Case 4
                PopulateCaretakers
            Case 5
                PopulateCaretakers
            Case 6
                PopulateCaretakers
            Case 7
                PopulateOffices
            Case 8
                PopulateCircles
            Case 9
                PopulateDivisions
            Case 10
                PopulateSubDivisions
            Case 11
                PopulateSections
            Case 12
                PopulateOffices
        End Select
    End Sub
    
    Public Sub PopulateOffices()
        Dim vParamIn(0)     As Variant
        Dim vParamOut       As Variant
        Dim objDb           As New clsDB
        Dim mSql            As String
        Dim Rec             As New ADODB.Recordset
        
        fgSearch.Cols = 3
        Me.Caption = "Search Offices/Institutions"
        lblHeading.Caption = "   Offices/Insitutions"
        fgSearch.Cell(flexcpText, 0, 0) = "Sl No"
        fgSearch.Cell(flexcpText, 0, 1) = "OfficeId"
        fgSearch.Cell(flexcpText, 0, 2) = "Office/Institution"
        fgSearch.Cols = 3
        vParamIn(0) = lngWrBillSearchZoneID ' intCareTakerID
        'vParamIn(1) = intWrBillSearchWardID
        'ExecuteSP "spsnWrBillMastersOfficeInstitution_S", rselect, adCmdStoredProc, vParamIn, vParamOut, conSanchaya
'        mSQL = "Select snWrBillMastersOfficeInstitution.intID,chvName From snWrBillCareTakerChild"
'        mSQL = mSQL + " Inner Join snWrBillMastersOfficeInstitution On snWrBillCareTakerChild.intOfficeInstitutionID = snWrBillMastersOfficeInstitution.intID"
'        If intWrBillCaretakerID <> 0 Then
'            mSQL = mSQL + " Where snWrBillCareTakerChild.intCareTakerID =" & lngWrBillSearchZoneID
'        End If
        mSql = "Select snWrBillMastersOfficeInstitution.intID,chvName From snWrBillMastersOfficeInstitution"
        If intWrBillCaretakerID <> 0 Then
        mSql = mSql + " Inner Join snWrBillCareTakerChild On snWrBillMastersOfficeInstitution.intID = snWrBillCareTakerChild.intOfficeInstitutionID"
        mSql = mSql + " Where snWrBillCareTakerChild.intCareTakerID =" & intWrBillCaretakerID
        End If
        objDb.ExecuteSP mSql, , vParamOut, , mCnn, adCmdText
        If IsArray(vParamOut) Then
            For i = 0 To UBound(vParamOut, 2)
                fgSearch.Rows = fgSearch.Rows + 1
                fgSearch.Cell(flexcpText, i + 1, 0) = i + 1 'Sl No
                fgSearch.Cell(flexcpText, i + 1, 1) = vParamOut(0, i) 'OfficeId
                fgSearch.Cell(flexcpText, i + 1, 2) = vParamOut(1, i) 'Office Name
            Next i
            fgSearch.Rows = fgSearch.Rows - 1
        End If
    End Sub
    
    Public Sub PopulateCaretakers()
        Dim vParamIn(1)     As Variant
        Dim vParamOut       As Variant
        Dim objDb           As New clsDB
        
        fgSearch.Cols = 4
        Me.Caption = "Search Caretakers"
        lblHeading.Caption = "   Caretakers"
        fgSearch.Cell(flexcpText, 0, 0) = "Sl No"
        fgSearch.Cell(flexcpText, 0, 1) = "CareTakerId"
        fgSearch.Cell(flexcpText, 0, 2) = "CareTaker Name"
        fgSearch.Cell(flexcpText, 0, 3) = "Designation"
        vParamIn(0) = lngWrBillSearchZoneID
        vParamIn(1) = intWrBillSearchWardID
        'ExecuteSP "spsnWrBillMastersCareTakers_S", rselect, adCmdStoredProc, vParamIn, vParamOut, conSanchaya
        objDb.ExecuteSP "spsnWrBillMastersCareTakers_S", vParamIn, vParamOut, , mCnn, adCmdStoredProc
        If IsArray(vParamOut) Then
            For i = 0 To UBound(vParamOut, 2)
                fgSearch.Rows = fgSearch.Rows + 1
                fgSearch.Cell(flexcpText, i + 1, 0) = i + 1 'Sl No
                fgSearch.Cell(flexcpText, i + 1, 1) = vParamOut(0, i) 'CareTakerId
                fgSearch.Cell(flexcpText, i + 1, 2) = vParamOut(1, i) 'CareTaker Name
                fgSearch.Cell(flexcpText, i + 1, 3) = vParamOut(2, i) 'Designation
            Next i
            fgSearch.Rows = fgSearch.Rows - 1
        End If
    End Sub
    
    Public Sub PopulateCircles()
        Dim vParamOut       As Variant
        Dim objDb           As New clsDB
        
        fgSearch.Cols = 4
        Me.Caption = "Search Circles"
        lblHeading.Caption = "   Circles"
        fgSearch.Cell(flexcpText, 0, 0) = "Sl No"
        fgSearch.Cell(flexcpText, 0, 1) = "CircleID"
        fgSearch.Cell(flexcpText, 0, 2) = "Circle"
        fgSearch.Cols = 3
        'ExecuteSP "spsnWrBillMastersCareTakers_S", rselect, adCmdStoredProc, vParamIn, vParamOut, conSanchaya
        objDb.ExecuteSP "snWrBillMastersCircle_S", , vParamOut, , mCnn, adCmdStoredProc
        If IsArray(vParamOut) Then
            For i = 0 To UBound(vParamOut, 2)
                fgSearch.Rows = fgSearch.Rows + 1
                fgSearch.Cell(flexcpText, i + 1, 0) = i + 1 'Sl No
                fgSearch.Cell(flexcpText, i + 1, 1) = vParamOut(0, i) 'CareTakerId
                fgSearch.Cell(flexcpText, i + 1, 2) = vParamOut(1, i) 'CareTaker Name
            Next i
            fgSearch.Rows = fgSearch.Rows - 1
        End If
    End Sub
    
    Public Sub PopulateDivisions()
        Dim vParamIn(1)     As Variant
        Dim vParamOut       As Variant
        Dim objDb           As New clsDB
        Dim mSql            As String
        
        fgSearch.Cols = 4
        Me.Caption = "Search Divisions"
        lblHeading.Caption = "   Divisions"
        fgSearch.Cell(flexcpText, 0, 0) = "Sl No"
        fgSearch.Cell(flexcpText, 0, 1) = "DivisionID"
        fgSearch.Cell(flexcpText, 0, 2) = "Division"
        fgSearch.Cols = 3
        If intWrBillCircleID <> 0 Then
            mSql = "SELECT intId,chvName From snWrBillMastersDivision Where tnyCircle = " & intWrBillCircleID
            'ExecuteSP "spsnWrBillMastersCareTakers_S", rselect, adCmdStoredProc, vParamIn, vParamOut, conSanchaya
        Else
            mSql = "SELECT intId,chvName From snWrBillMastersDivision"
        End If
        objDb.ExecuteSP mSql, , vParamOut, , mCnn, adCmdText
        If IsArray(vParamOut) Then
            For i = 0 To UBound(vParamOut, 2)
                fgSearch.Rows = fgSearch.Rows + 1
                fgSearch.Cell(flexcpText, i + 1, 0) = i + 1 'Sl No
                fgSearch.Cell(flexcpText, i + 1, 1) = vParamOut(0, i) 'CareTakerId
                fgSearch.Cell(flexcpText, i + 1, 2) = vParamOut(1, i) 'CareTaker Name
            Next i
            fgSearch.Rows = fgSearch.Rows - 1
        End If
    End Sub
    
    Public Sub PopulateSubDivisions()
        Dim vParamIn(1)     As Variant
        Dim vParamOut       As Variant
        Dim objDb           As New clsDB
        Dim mSql            As String
        
        fgSearch.Cols = 4
        Me.Caption = "Search Sub Divisions"
        lblHeading.Caption = "   Sub Divisions"
        fgSearch.Cell(flexcpText, 0, 0) = "Sl No"
        fgSearch.Cell(flexcpText, 0, 1) = "SubDivisionID"
        fgSearch.Cell(flexcpText, 0, 2) = "SubDivision"
        fgSearch.Cols = 3
'        vParamIn(0) = lngWrBillSearchZoneID
'        vParamIn(1) = intWrBillSearchWardID
        mSql = "SELECT intId,chvName From snWrBillMastersSubDivision "
        If intWrBillCircleID <> 0 And intWrBillDivisionID <> 0 Then
            mSql = mSql + " Where tnyCircle = " & intWrBillCircleID
            mSql = mSql + " And tnyDivision = " & intWrBillDivisionID
        ElseIf intWrBillCircleID <> 0 Then
            mSql = mSql + " Where tnyCircle = " & intWrBillCircleID
        ElseIf intWrBillDivisionID <> 0 Then
             mSql = mSql + " Where tnyDivision = " & intWrBillDivisionID
        End If
        objDb.ExecuteSP mSql, , vParamOut, , mCnn, adCmdText
        'ExecuteSP "spsnWrBillMastersCareTakers_S", rselect, adCmdStoredProc, vParamIn, vParamOut, conSanchaya
        
        If IsArray(vParamOut) Then
            For i = 0 To UBound(vParamOut, 2)
                fgSearch.Rows = fgSearch.Rows + 1
                fgSearch.Cell(flexcpText, i + 1, 0) = i + 1 'Sl No
                fgSearch.Cell(flexcpText, i + 1, 1) = vParamOut(0, i) 'CareTakerId
                fgSearch.Cell(flexcpText, i + 1, 2) = vParamOut(1, i) 'CareTaker Name
            Next i
            fgSearch.Rows = fgSearch.Rows - 1
        End If
    End Sub
    
    Public Sub PopulateSections()
        Dim vParamIn(1)     As Variant
        Dim vParamOut       As Variant
        Dim objDb           As New clsDB
        Dim mSql            As String
        
        fgSearch.Cols = 4
        Me.Caption = "Search Sections"
        lblHeading.Caption = "   Sections"
        fgSearch.Cell(flexcpText, 0, 0) = "Sl No"
        fgSearch.Cell(flexcpText, 0, 1) = "SectionID"
        fgSearch.Cell(flexcpText, 0, 2) = "Section"
        fgSearch.Cols = 3
        mSql = "SELECT intId,chvName FROM snWrBillMastersSection"
        If intWrBillCircleID <> 0 And intWrBillDivisionID <> 0 And intWrBillSubDivisionID <> 0 Then
            mSql = mSql + " Where tnyCircle = " & intWrBillCircleID
            mSql = mSql + " And tnyDivision = " & intWrBillDivisionID
            mSql = mSql + " And tnySubDivision = " & intWrBillSubDivisionID
        ElseIf intWrBillCircleID <> 0 Then
            mSql = mSql + " Where tnyCircle = " & intWrBillCircleID
        ElseIf intWrBillDivisionID <> 0 Then
             mSql = mSql + " Where tnyDivision = " & intWrBillDivisionID
        ElseIf intWrBillSubDivisionID <> 0 Then
            mSql = mSql + " Where tnySubDivision = " & intWrBillSubDivisionID
        End If
        'ExecuteSP "spsnWrBillMastersCareTakers_S", rselect, adCmdStoredProc, vParamIn, vParamOut, conSanchaya
        objDb.ExecuteSP mSql, , vParamOut, , mCnn, adCmdText
        If IsArray(vParamOut) Then
            For i = 0 To UBound(vParamOut, 2)
                fgSearch.Rows = fgSearch.Rows + 1
                fgSearch.Cell(flexcpText, i + 1, 0) = i + 1 'Sl No
                fgSearch.Cell(flexcpText, i + 1, 1) = vParamOut(0, i) 'CareTakerId
                fgSearch.Cell(flexcpText, i + 1, 2) = vParamOut(1, i) 'CareTaker Name
            Next i
            fgSearch.Rows = fgSearch.Rows - 1
        End If
        
    End Sub
