VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchSubsidiaryAccountHeads 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchSubsidiaryAccountHeads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   270
      Left            =   5940
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4185
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Width           =   7695
      _cx             =   13573
      _cy             =   7382
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
      Rows            =   13
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchSubsidiaryAccountHeads.frx":1CCA
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
   Begin VB.ComboBox cmbSubLegerType 
      Height          =   390
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   4485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SubLedger Type"
      Height          =   270
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Searching Subsidiary Account Heads"
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
      TabIndex        =   0
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "frmSearchSubsidiaryAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private mvarSubLedgerTypeID As Integer
    Private mcheckIMPO As Integer
    Public Property Let SubLedgerType(mData As Long)
        'For Selecting SubLedgerType
        mvarSubLedgerTypeID = mData
    End Property
    Public Property Let checkIMPO(mData As Long)
        mcheckIMPO = mData
    End Property
    Private Function GetImpementingOfficer() As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objDb As New clsDB
            Dim mRowCnt As Integer
            
            If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSql = "Select * from suImplementingOfficer Where intLBTypeID = " & gbLBType & " Order By vchImplementingOfficer"
                Rec.Open mSql, mCnn
                vsGrid.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intImplementingOfficerID), "", Rec!intImplementingOfficerID)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchImplementingOfficer), "", Rec!vchImplementingOfficer)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchImplementingOfficerCode), "", Rec!vchImplementingOfficerCode)
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetImpementingOfficer = True
        Exit Function
err:
        MsgBox (Error$)
    End Function

        
    Private Function GetAuthorisedAgencies() As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objDb As New clsDB
            Dim mRowCnt As Integer
            
            If objDb.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
                mSql = "Select * from M_ImplAgency Where intImplAgencyTypeID = 6 order By chvImplAgency"
                Rec.Open mSql, mCnn
                vsGrid.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intImplAgencyID), "", Rec!intImplAgencyID)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvEngImplAgency), "", Rec!chvImplAgency)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!chrImplAgencyCode), "", Rec!chrImplAgencyCode)
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetAuthorisedAgencies = True
        Exit Function
err:
        MsgBox (Error$)
    End Function

'''    Private Function GetContractors() As Boolean
'''        On Error GoTo Err:
'''            Dim mCnn As New ADODB.Connection
'''            Dim Rec As New ADODB.Recordset
'''            Dim mSql As String
'''            Dim objDb As New clsDB
'''            Dim mRowCnt As Integer
'''
'''            If objDb.SetConnection(mCnn) Then
'''                mSql = "Select * from faSubSidiaryAccountHeads Where intSubLedgerTypeID = 7 Order By vchName"
'''                Rec.Open mSql, mCnn
'''                vsGrid.Rows = 2
'''                mRowCnt = 1
'''                While Not (Rec.EOF Or Rec.BOF)
'''                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
'''                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
'''                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
'''                    Rec.MoveNext
'''                    vsGrid.Rows = vsGrid.Rows + 1
'''                    mRowCnt = mRowCnt + 1
'''                Wend
'''            Else
'''                MsgBox "Connection to Finance does not Exist, Please Contact your System Operator", vbInformation
'''            End If
'''            GetContractors = True
'''        Exit Function
'''Err:
'''        MsgBox (Error$)
'''    End Function
'''
'''    Private Function GetSuppliers() As Boolean
'''        On Error GoTo Err:
'''            Dim mCnn As New ADODB.Connection
'''            Dim Rec As New ADODB.Recordset
'''            Dim mSql As String
'''            Dim objDb As New clsDB
'''            Dim mRowCnt As Integer
'''
'''            If objDb.SetConnection(mCnn) Then
'''                mSql = "Select * from faSubSidiaryAccountHeads Where intSubLedgerTypeID = 8 Order By vchName"
'''                Rec.Open mSql, mCnn
'''                vsGrid.Rows = 2
'''                mRowCnt = 1
'''                While Not (Rec.EOF Or Rec.BOF)
'''                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
'''                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
'''                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
'''                    Rec.MoveNext
'''                    vsGrid.Rows = vsGrid.Rows + 1
'''                    mRowCnt = mRowCnt + 1
'''                Wend
'''            Else
'''                MsgBox "Connection to Finance does not Exist, Please Contact your System Operator", vbInformation
'''            End If
'''            GetSuppliers = True
'''        Exit Function
'''Err:
'''        MsgBox (Error$)
'''    End Function
'''
'''     Private Function GetSubsidiaryCashBook() As Boolean
'''        On Error GoTo Err:
'''            Dim mCnn As New ADODB.Connection
'''            Dim Rec As New ADODB.Recordset
'''            Dim mSql As String
'''            Dim objDb As New clsDB
'''            Dim mRowCnt As Integer
'''
'''            If objDb.SetConnection(mCnn) Then
'''                mSql = "Select * from faSubSidiaryAccountHeads Where intSubLedgerTypeID = 12 Order By vchName"
'''                Rec.Open mSql, mCnn
'''                vsGrid.Rows = 2
'''                mRowCnt = 1
'''                While Not (Rec.EOF Or Rec.BOF)
'''                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
'''                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
'''                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
'''                    Rec.MoveNext
'''                    vsGrid.Rows = vsGrid.Rows + 1
'''                    mRowCnt = mRowCnt + 1
'''                Wend
'''            Else
'''                MsgBox "Connection to Finance does not Exist, Please Contact your System Operator", vbInformation
'''            End If
'''            GetSubsidiaryCashBook = True
'''        Exit Function
'''Err:
'''        MsgBox (Error$)
'''    End Function
'''
'''    Private Function GetEmployees() As Boolean
'''        On Error GoTo Err:
'''            Dim mCnn As New ADODB.Connection
'''            Dim Rec As New ADODB.Recordset
'''            Dim objDb As New clsDB
'''            Dim mSql As String
'''            Dim mRowCnt As Integer
'''
'''            If objDb.CreateNewConnection(mCnn, enuSourceString.Sthapana) Then
'''                mSql = "Select chvEmpName,intEmpId from TB_EmployeeDetails_Trn Order By chvEmpName"
'''                Rec.Open mSql, mCnn
'''                vsGrid.Rows = 2
'''                mRowCnt = 1
'''                While Not (Rec.EOF Or Rec.BOF)
'''                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intEmpId), "", Rec!intEmpId)
'''                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!chvEmpName), "", Rec!chvEmpName)
'''                    Rec.MoveNext
'''                    vsGrid.Rows = vsGrid.Rows + 1
'''                    mRowCnt = mRowCnt + 1
'''                Wend
'''                GetEmployees = True
'''            Else
'''                MsgBox "Connection To Sthapana does not exist, Please Contact your System Administrator", vbInformation
'''            End If
'''        Exit Function
'''Err:
'''        MsgBox (Error$)
'''    End Function
    
    
    Private Function GetSubLedgersFromFin(ByVal intSubLedgerID As Long) As Boolean
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objDb As New clsDB
            Dim mRowCnt As Integer

            If objDb.SetConnection(mCnn) Then
                mSql = "Select * from faSubSidiaryAccountHeads Where intSubLedgerTypeID = " & intSubLedgerID & " Order By vchName"
                Rec.Open mSql, mCnn
                vsGrid.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
                    If Not IsNull(Rec!vchName) Then
                        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    ElseIf Not IsNull(Rec!vchTitle) Then
                        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
                    End If
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchSubTitle), "", Rec!vchSubTitle)
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact your System Operator", vbInformation
            End If
            GetSubLedgersFromFin = True
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Sub cmbSubLegerType_Click()
         On Error GoTo err:
                vsGrid.Clear 1, 1
'''                If cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 4 Then
'''                    Call GetAuthorisedAgencies
'''                ElseIf cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex) = 1 Then
'''                    Call GetImpementingOfficer
'''                Else
'''                    Call GetSubLedgersFromFin(cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex))
'''                End If
                If mcheckIMPO = 1 Then
                    Call GetImpementingOfficer
                Else
                    Call GetSubLedgersFromFin(cmbSubLegerType.ItemData(cmbSubLegerType.ListIndex))
                End If
            Exit Sub
err:
            MsgBox (Error$)
    End Sub

    Private Sub Command1_Click()
'            Dim objDb As New clsDB
'            Dim mCnn As New ADODB.Connection
'            frmSearchMasters.SQLQry = "Select intSubLedgerTypeID,vchSubLedgerType from faSubLedgerTypes"
'            frmSearchMasters.LetConnection = enuSourceString.Saankhya
'            frmSearchMasters.Show vbModal, Me
'            MsgBox (gbSearchStr)
'        frmSearchTransactionType.Show vbModal
'        MsgBox gbSearchStr
'        frmSearchFunction.Show vbModal
    End Sub

    Private Sub Form_Load()
        On Error GoTo err:
            gbSearchID = -1
            gbSearchCode = ""
            gbSearchStr = ""
            If mvarSubLedgerTypeID > 0 Then
                PopulateList cmbSubLegerType, "Select vchSubLedgerType,intSubLedgerTypeID From faSubLedgerTypes Where intSubLedgerTypeID=" & mvarSubLedgerTypeID, , True, , True
                cmbSubLegerType.ListIndex = 1
            Else
                PopulateList cmbSubLegerType, "Select vchSubLedgerType,intSubLedgerTypeID From faSubLedgerTypes ", , True, , True
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        mvarSubLedgerTypeID = -1
        mcheckIMPO = -1
    End Sub

    Private Sub vsGrid_Click()
        vsGrid.Cell(flexcpBackColor, 1, 0, vsGrid.Rows - 1, 3) = vbWhite
        vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 3) = &HC0C0FF
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
            gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 0))
            gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 1))
            gbSearchCode = Trim(vsGrid.TextMatrix(vsGrid.Row, 2))
            Unload Me
        End If
    End Sub

    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call vsGrid_Click
            Call vsGrid_DblClick
        End If
    End Sub
