VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRequisitionAuthorisation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authorization"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   Icon            =   "frmRequisitionAuthorisation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Requisition Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   11220
      Begin VB.TextBox txtPurpose 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4365
         TabIndex        =   27
         Top             =   3195
         Width           =   4605
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   26
         Top             =   2790
         Width           =   3480
      End
      Begin VB.TextBox txtIMPO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   25
         Top             =   2385
         Width           =   3480
      End
      Begin VB.TextBox txtMicroSector 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   24
         Top             =   1980
         Width           =   3480
      End
      Begin VB.TextBox txtSubSector 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   23
         Top             =   1575
         Width           =   3480
      End
      Begin VB.TextBox txtProjectName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   22
         Top             =   1170
         Width           =   3480
      End
      Begin VB.TextBox txtProjNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   21
         Top             =   1170
         Width           =   1050
      End
      Begin VB.TextBox txtRequisitionDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   20
         Top             =   765
         Width           =   2220
      End
      Begin VB.TextBox txtRequisitionNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4365
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   19
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label Label13 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   18
         Top             =   2835
         Width           =   150
      End
      Begin VB.Label Label12 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   17
         Top             =   3195
         Width           =   150
      End
      Begin VB.Label Label11 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   16
         Top             =   810
         Width           =   150
      End
      Begin VB.Label Label8 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   15
         Top             =   1215
         Width           =   150
      End
      Begin VB.Label Label7 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   14
         Top             =   1575
         Width           =   150
      End
      Begin VB.Label Label6 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   13
         Top             =   2025
         Width           =   150
      End
      Begin VB.Label Label2 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4185
         TabIndex        =   12
         Top             =   405
         Width           =   150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Purpose"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   11
         Top             =   3195
         Width           =   2265
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Implementing Officer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   9
         Top             =   2385
         Width           =   2265
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Requisition No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Allotment Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   7
         Top             =   765
         Width           =   2265
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Requisition Amount "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   6
         Top             =   2790
         Width           =   2265
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "SubSector"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   1530
         Width           =   2265
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "MicroSector"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   1980
         Width           =   2265
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Project No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         Top             =   1170
         Width           =   2265
      End
   End
   Begin VB.CommandButton cmdAuthorize 
      Caption         =   "AUTHORIZE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4230
      TabIndex        =   1
      Top             =   4185
      Width           =   1635
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   4860
      Width           =   11265
      _cx             =   19870
      _cy             =   5424
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRequisitionAuthorisation.frx":1CCA
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
Attribute VB_Name = "frmRequisitionAuthorisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private Sub cmdAuthorize_Click()
        Dim mCnn             As New ADODB.Connection
        Dim objDB            As New clsDB
        Dim Rec              As New ADODB.Recordset
        Dim mArrIN           As Variant
        Dim mArrOut          As Variant
        Dim mAuthorizationNo As Variant
        Dim mSQL             As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya

        If txtRequisitionNo.Text <> "" Then
            mArrIN = Array(val(txtRequisitionNo.Tag), _
                            Null, _
                            gbTransactionDate, _
                            gbUserID, _
                            1, _
                            gbFinancialYearID _
                            )
            objDB.ExecuteSP "spSaveAuthorizeReqInbox", mArrIN, mArrOut, , mCnn, adCmdStoredProc
            mAuthorizationNo = mArrOut(0, 0)
        'Else
        End If

        mSQL = "Update faAllotments set vchAuthorizationNo=" & mAuthorizationNo & " ,"
        mSQL = mSQL + "  dtAuthorizationDate ='" & DdMmmYy(gbTransactionDate) & "',"
        mSQL = mSQL + "  fltAuthorizedAmt =" & txtAmount.Text & ", intAuthorizedByUserID =" & gbUserID & ", tnyStage=2 WHERE vchRequisitionNO= " & txtRequisitionNo.Text & ""
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        cmdAuthorize.Enabled = False
        
        Call GetProjectDetails
        Call SynTOWEB(mAuthorizationNo)
    End Sub
    
    Private Sub SynTOWEB(mAuthorizationNo As Variant) '******************UPDATE intAuthorizedUserID ,dtAuthorizedDate ,intAuthorizationNo  TO DB_FinanceWEB******************
        Dim mArrIN As Variant
        Dim mArrOut As Variant
        Dim mUrl   As String
        Dim client1 As New MSSOAPLib.SoapClient
        Dim objSOAP As Variant
        Dim clnt As New SoapClient30
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL  As String
        Dim mLoop As Integer
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
        '--------------'
        ' Web Service  '
        '--------------'
    
        mUrl = gbDefaultUrlForRequisition
        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        objSOAP.mssoapinit mUrl + "?WSDL"

        'For mLoop = 1 To vsGrid.Rows - 1
            'If vsGrid.TextMatrix(mLoop, 1) <> "" Then
              mArrOut = objSOAP.UpdateAuthorizationNoToWEBReqInbox(val(txtRequisitionNo.Tag), gbLBID, gbUserID, gbTransactionDate, mAuthorizationNo)
            
              mSQL = "UPDATE faAllotments SET   vchAuthorizationNo=" & mAuthorizationNo & ", dtAuthorizationDate = " & gbTransactionDate & " , "
              mSQL = mSQL + " fltAuthorizedAmt=" & txtAmount.Text & ", intAuthorizedByUserID=" & gbUserID & ",tnyStage=2 "
              mSQL = mSQL + " WHERE vchRequisitionNo=" & txtRequisitionNo.Text & "  "
              objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            'End If
        'Next mLoop
    End Sub
    Private Sub Form_Activate()
        'Me.Left = 0
        'Me.Top = 0
    End Sub

    Private Sub GetProjectDetails()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mAccHeadId As Integer
        Dim mRowCnt As Integer
        Dim mLoop As Integer
        
        If objDB.SetConnection(mCnn) Then
            mSQL = " SELECT * FROM faAllotments"
            mSQL = mSQL + "  INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faAllotments.intSourceID"
            mSQL = mSQL + "  INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faAllotments.intFundCategoryID"
            mSQL = mSQL + "  WHERE numProjectID = " & txtProjNo.Tag & " "
            mSQL = mSQL + "  AND intFinancialYearID= " & gbFinancialYearID & " "
            
            Rec.CursorLocation = adUseClient
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            mRowCnt = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
                If Rec!dtRequisitionDate <> "" Then
                    vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
                Else
                    vsGrid.TextMatrix(mRowCnt, 1) = ""
                End If
                vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!fltRequestedAmt), "", Rec!fltRequestedAmt)
                vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchAuthorizationNo), "", Rec!vchAuthorizationNo)
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                
                If Rec!tnyStage = 0 And Rec!tnyStatus = 0 Then
                    vsGrid.TextMatrix(mRowCnt, 8) = "Request Waiting"
                ElseIf Rec!tnyStage = 1 And Rec!tnyStatus = 0 Then
                    vsGrid.TextMatrix(mRowCnt, 8) = "Approved"
                ElseIf Rec!tnyStage = 2 And Rec!tnyStatus = 0 Then
                    vsGrid.TextMatrix(mRowCnt, 8) = "Authorized"
                End If
                
                If Rec!tnyStatus = 2 Then  'IF REQ IS CANCELLED
                    For mLoop = 0 To vsGrid.Cols - 1
                        vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HC0E0FF
                    Next mLoop
                      
                End If
                
                Rec.MoveNext
                mRowCnt = mRowCnt + 1
            Wend
            Rec.Close
        End If
    End Sub
    
    Private Sub Form_Load()
        'frmProj.Enabled = False
       
    End Sub
    
    Public Sub GetProjectAllotmentDetails()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mAccHeadId As Integer
           
        If objDB.SetConnection(mCnn) Then
            mSQL = " SELECT * FROM faRequisitionInbox"
            mSQL = mSQL + " LEFT JOIN suImplementingOfficer ON suImplementingOfficer.intImplementingOfficerID=faRequisitionInbox.intImplementingOfficersID"
            mSQL = mSQL + " INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faRequisitionInbox.intSourceID"
            mSQL = mSQL + " INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faRequisitionInbox.intFundCategoryID"
            mSQL = mSQL + " LEFT JOIN suProjectDetails ON faRequisitionInbox.numProjectID=suProjectDetails.decProjectID AND intYearID=2014"
            mSQL = mSQL + " Left JOIN faSubSector On faSubSector.intSubSecID=faRequisitionInbox.intSubSecID"
            mSQL = mSQL + " Left JOIN faMicroSectorHeads On faMicroSectorHeads.intMircoSectorID=faRequisitionInbox.intMircoSectorID"
            
            mSQL = mSQL + " WHERE intID=" & txtRequisitionNo.Tag
            
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtRequisitionNo.Text = IIf(Rec!intRequisitionNo = 0, "", Rec!intRequisitionNo)
                If Rec!dtRequestDate <> "" Then
                    txtRequisitionDate.Text = DdMmmYy(Rec!dtRequestDate)
                Else
                    txtRequisitionDate.Text = ""
                End If
                txtProjNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
                txtProjNo.Tag = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
                txtProjectName.Text = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
                txtSubSector.Text = IIf(IsNull(Rec!vchSubSectorEng), "", Rec!vchSubSectorEng)
                txtMicroSector.Text = IIf(IsNull(Rec!vchMicroSector), "", Rec!vchMicroSector)
                txtIMPO.Text = IIf(IsNull(Rec!vchImplementingOfficer), "", Rec!vchImplementingOfficer)
                txtAmount.Text = IIf(IsNull(Rec!fltRequestedAmt), "", Rec!fltRequestedAmt)
                txtPurpose.Text = IIf(IsNull(Rec!nvchDescription), "", Rec!nvchDescription)
            End If
            Rec.Close
        End If
        Call GetProjectDetails
    End Sub


