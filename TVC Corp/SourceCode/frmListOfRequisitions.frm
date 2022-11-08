VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfRequisitions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Requisitions For Fund From Implementing Officers"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListOfRequisitions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmListOfRequisitions.frx":1CCA
   ScaleHeight     =   7740
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTreasuryBill 
      Caption         =   "&Treasury Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5580
      TabIndex        =   13
      Top             =   7215
      Width           =   1905
   End
   Begin VB.CommandButton cmdProceedings 
      Caption         =   "&Proceedings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3660
      TabIndex        =   12
      Top             =   7215
      Width           =   1905
   End
   Begin VB.CommandButton cmdLetterOfAllotment 
      Caption         =   "&Letter of Allotment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1740
      TabIndex        =   11
      Top             =   7215
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13860
      TabIndex        =   15
      Top             =   0
      Width           =   13860
      Begin VB.Label lblMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   4650
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   8000
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F4FAFA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   13800
      TabIndex        =   14
      Top             =   7155
      Width           =   13860
      Begin VB.CommandButton cmdAppropriationControlRegister 
         Caption         =   "Appropriation Control Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10440
         TabIndex        =   19
         Top             =   30
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   10
         Top             =   30
         Width           =   1395
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   3420
         Top             =   270
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   765
      Width           =   13860
      _cx             =   24447
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      BackColorFixed  =   16055034
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   14540253
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
      Rows            =   50
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfRequisitions.frx":200C
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
      WordWrap        =   -1  'True
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4FAFA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   0
      TabIndex        =   16
      Top             =   5850
      Width           =   13845
      Begin VB.CommandButton cmdSearchRequisitions 
         Caption         =   "search"
         Height          =   345
         Left            =   6510
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkAuthorized 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4FAFA&
         Caption         =   " AUTHORIZED"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11700
         TabIndex        =   7
         Top             =   315
         Width           =   1860
      End
      Begin VB.CommandButton cmdSearchSource 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5805
         TabIndex        =   9
         Top             =   675
         Width           =   345
      End
      Begin VB.TextBox txtSource 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   690
         Width           =   3600
      End
      Begin VB.TextBox txtIMPO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   263
         Width           =   5085
      End
      Begin VB.CommandButton cmdSearchIMPO 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   2
         Top             =   285
         Width           =   255
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8100
         TabIndex        =   3
         Top             =   263
         Width           =   1215
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9825
         TabIndex        =   5
         Top             =   255
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpkrFromDate 
         Height          =   315
         Left            =   9330
         TabIndex        =   4
         Top             =   270
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66387969
         CurrentDate     =   40134
      End
      Begin MSComCtl2.DTPicker dtpkrToDate 
         Height          =   315
         Left            =   11040
         TabIndex        =   6
         Top             =   255
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66387969
         CurrentDate     =   40134
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Of Fund"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   630
         TabIndex        =   21
         Top             =   765
         Width           =   1470
      End
      Begin VB.Label lblImpOfficer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Implementing Officer"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   30
         TabIndex        =   18
         Top             =   300
         Width           =   2100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7605
         TabIndex        =   17
         Top             =   300
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   9660
         X2              =   9735
         Y1              =   435
         Y2              =   420
      End
   End
End
Attribute VB_Name = "frmListOfRequisitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Option Explicit
    Dim mImpOff     As Boolean
    Private intIMPO As Integer ' 1=Implementing Officer, 2=Implementing Agency,3=Accredited Agency
    Dim mPreviousYear As Variant
    
    Dim mLoadMode       As Integer      '10-For UNAUTHORIZED DRAWAL
    '*********************************************************************************************'
    '                               Form to list all the Requisitions                             '
    '*********************************************************************************************'
    Private Sub Fillgrid()
        
        Dim mCnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim msql  As String
        
        
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If mLoadMode = 10 Then 'And mPreviousYear = 1
            lblMsg.Visible = True
            lblMsg.Caption = "UNAUTHORIZED DRAWAL"
        ElseIf mPreviousYear = 1 Then
            lblMsg.Visible = True
            lblMsg.Caption = "PREVIOUS YEAR PENDING REQUISITIONS"
        Else
            lblMsg.Visible = False
        End If
        
        
        msql = " SELECT Distinct vchRequisitionNo,dtRequisitionDate,vchNameofIMPO,vchDesignation,fltRequestedAmt,vchProjectNo,vchSourceFundName,vchTransactionCategory," & vbNewLine
        msql = msql + " CASE   WHEN Isnull(tnyStage,0)=1 and Isnull(fltTotalAltIssued,0)=0 THEN 'Waiting for Approval'    " & vbNewLine
     
        If chkAuthorized.Value = vbChecked Then
            msql = msql + " WHEN Isnull(tnyStage,0)=2 and Isnull(faAllotments.tnyStatus,0)=1  THEN 'Authorized' " & vbNewLine
        Else
            msql = msql + " WHEN Isnull(tnyStage,0)=2 and Isnull(faAllotments.tnyStatus,0)=0   THEN 'Approved'" & vbNewLine
        End If
        
        msql = msql + " END Status,"
        msql = msql + " intID , intAuthorizedByUserID, intSourceID, intFundCategoryID, " & vbNewLine
        msql = msql + " fltAuthorizedAmt,ISNULL(faAllotments.tnyStage,0) tnyStage,ISNULL(faAllotments.tnyStatus,0) tnyStatus, ISNULL(faAllotments.intTreasuryID,0) intTreasuryID  From faAllotments " & vbNewLine
        msql = msql + " Left Join suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID" & vbNewLine
        msql = msql + " Left Join faTransactionCategory On faAllotments.intFundCategoryID = faTransactionCategory.intCategoryID" & vbNewLine
        If chkAuthorized.Value = vbUnchecked Then
        If mPreviousYear = 1 Then
            msql = msql + " Inner Join faPendingTaskRequest On faPendingTaskRequest.intKeyID=faAllotments.intID"
        End If
        End If
        
        msql = msql + " Where  " & vbNewLine
        msql = msql + " faAllotments.tnyStatus <> 2 " & vbNewLine
       
        
        If mPreviousYear = 1 Then
            msql = msql + " And faAllotments.intFinancialYearID = " & gbFinancialYearID - 1 & "" & vbNewLine
        Else
            msql = msql + " And faAllotments.intFinancialYearID = " & gbFinancialYearID & "" & vbNewLine
        End If
        
        If chkAuthorized.Value = vbChecked Then
            msql = msql + "  AND Isnull(tnyStage,0) =2  and Isnull(faAllotments.tnyStatus,0)=1 " & vbNewLine
        Else
            If mPreviousYear = 1 And mLoadMode = 10 Then
                msql = msql + " And faPendingTaskRequest.intTaskID = 16 "
            ElseIf mPreviousYear = 1 Then
                msql = msql + " And faPendingTaskRequest.intTaskID in (3 ,13)"
            End If
            msql = msql + "  AND Isnull(tnyStage,0) in (1,2)  and Isnull(faAllotments.tnyStatus,0)=0 " & vbNewLine
        End If
        
        If mImpOff = True Then
            msql = msql + " And faAllotments.intImplementingOfficersID = " & txtImpo.Tag
        End If
        If txtSource.Tag <> "" Then
            msql = msql + " And faAllotments.intSourceID = " & txtSource.Tag
        End If
        If txtFromDate.Text <> "" And txtToDate.Text <> "" Then
            msql = msql + " And faAllotments.dtRequisitionDate between  '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' " & vbNewLine
        End If
        
        
        '***********FOR UNAUTHORIZED DRAWAL*************************
        If mLoadMode = 10 Then
            msql = msql + "  And IsNull(faAllotments.tnyTypeID,0)=3"
            
        Else
            msql = msql + "  And IsNull(faAllotments.tnyTypeID,0)=0"
        End If
        '***********END*********************************************
        
        msql = msql + " Order by dtRequisitionDate desc,vchRequisitionNo desc"
        Rec.CursorLocation = adUseClient
        Rec.Open msql, mCnn, adOpenDynamic, adLockOptimistic
        vsGrid.Clear 1, 0
        vsGrid.Rows = 1
        If Not (Rec.BOF And Rec.EOF) Then
            vsGrid.Rows = Rec.RecordCount + 1
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColSel = 16
            vsGrid.RowSel = vsGrid.Rows - 1
            msql = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = msql
        End If
        Rec.Close
        mImpOff = False
        'vsGrid.RowSel = 0
        
        vsGrid.Row = vsGrid.Rows - 1
        'vsGrid.Col = 0
       
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub FormInitialize()
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            ElseIf TypeOf ctrl Is OptionButton Then
                ctrl.Value = False
            ElseIf TypeOf ctrl Is ComboBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                ctrl.Tag = ""
            End If
        Next
        vsGrid.Clear 1, 1
        'txtFromDate.Text = DdMmmYy(gbStartingDate) '(DateAdd("d", -30, gbTransactionDate))
        'txtToDate.Text = DdMmmYy(gbTransactionDate)
    End Sub

    Private Sub chkAuthorized_Click()
        Call Fillgrid
    End Sub
    
Private Sub cmdSearchRequisitions_Click()
    frmSearchRequisitions.Show vbModal
End Sub

    Private Sub Form_Unload(Cancel As Integer)
        'mPreviousYear = 0
        mLoadMode = 0
    End Sub

    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
         If CDate(txtFromDate.Text) < CDate(gbStartingDate) Then
            If CDate(txtFromDate.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                txtFromDate.Text = DateAdd("yyyy", -1, gbStartingDate)
                txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
            End If
            txtToDate.Text = DateAdd("yyyy", -1, gbEndingDate)
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
            
            If IsDate(txtToDate.Text) Then
                Call PreviousYearRequisitions
                Call Fillgrid
            End If
            
            
            
            
            SaveSetting "SaankhyaDE", "App", "ListReqFromDate", txtFromDate.Text
            
            
         End If
    End Sub
    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        If txtFromDate.Text <> "" Then
            Call PreviousYearRequisitions
            Call Fillgrid
            If mPreviousYear = 1 Then
                txtToDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            End If
        End If
        If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
            MsgBox "Please Enter a Date Less than Or equal to ToDate", vbInformation
            txtFromDate.Text = ""
            txtFromDate.SetFocus
            Exit Sub
        End If
        
        SaveSetting "SaankhyaDE", "App", "ListReqToDate", txtToDate.Text
        
    End Sub

    Private Sub cmdAppropriationControlRegister_Click()
        Dim mCnn     As New ADODB.Connection
        Dim objDb    As New clsDB
        Dim Rec      As New ADODB.Recordset
        Dim msql     As String
        Dim mAuthAmt As Variant
        
        '*********************************************************************************************'
        '                   Procedure to view the Appropriation Control Register                      '
        '*********************************************************************************************'
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
               
        mAuthAmt = ""
        msql = "Select fltAuthorizedAmt From faAllotments Where intID = " & val(vsGrid.TextMatrix(vsGrid.Row, 9))
        Rec.Open msql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mAuthAmt = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
        End If
        Rec.Close
        If mAuthAmt = "" Then
            MsgBox "Please authorize the Allotment", vbInformation
            Exit Sub
        Else
            If vsGrid.TextMatrix(vsGrid.Row, 11) = 1 Then
                frmViewAllotmentLetter.ArrayIn = Array(val(vsGrid.TextMatrix(vsGrid.Row, 11)), val(vsGrid.TextMatrix(vsGrid.Row, 12)))
            Else
                frmViewAllotmentLetter.ArrayIn = Array(val(vsGrid.TextMatrix(vsGrid.Row, 11)), 0)
            End If
            frmViewAllotmentLetter.Mode = 7
            frmViewAllotmentLetter.Show vbModal
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdLetterOfAllotment_Click()
        Dim mCnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim msql            As String
        Dim mAltReceived    As Variant
        Dim mAuthAmt        As Variant
        Dim mYearID         As Variant
        Dim mSourceID       As Integer
        '*********************************************************************************************'
        '                           Procedure to view the Letter of Allotment                         '
        '*********************************************************************************************'
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        '============================================================='
        'BLOCKED BY AIBY on 27-JULY-2014
        'If vsGrid.Row > 0 Then
        '    If vsGrid.Row <> vsGrid.Rows - 1 Then
        '        If vsGrid.TextMatrix(vsGrid.Row + 1, 15) = 0 Then
        '            MsgBox "Verify the Preceding Requisition"
        '            Exit Sub
        '        End If
        '    End If
        'End If
        'END OF CODE
        '============================================================='
        
        
        If mPreviousYear = 1 Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        If mLoadMode = 10 Then
            frmIssueLetterOfAllotment.LoadMode = 10
        End If
               
        mAuthAmt = ""
        mAltReceived = ""
        msql = "Select fltAuthorizedAmt,intSourceID From faAllotments Where intID = " & val(vsGrid.TextMatrix(vsGrid.Row, 9))
        Rec.Open msql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mAuthAmt = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
            mSourceID = IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID)
        End If
        Rec.Close
        If mAuthAmt = "" Then
            MsgBox "Please authorize the Allotment", vbInformation
            Exit Sub
        Else
            msql = "Select fltTotalAltReceived From faAllotments Where intID = " & val(vsGrid.TextMatrix(vsGrid.Row, 9))
            Rec.Open msql, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                mAltReceived = IIf(IsNull(Rec!fltTotalAltReceived), "", Rec!fltTotalAltReceived)
            End If
            If mAltReceived = "" Then
                '**
                
                
                If gbSeatGroupID = gbSeatGroupSecretary Or gbSeatGroupID = gbSeatGroupChairPerson Then
PanchayatLog:
                    If vsGrid.TextMatrix(vsGrid.Row, 9) <> "" Then
                        frmIssueLetterOfAllotment.RequisitionID = val(vsGrid.TextMatrix(vsGrid.Row, 9))
                        frmIssueLetterOfAllotment.SourceID = mSourceID
                        Unload Me
                        frmIssueLetterOfAllotment.Show vbModal
                    Else
                        MsgBox "Please select a Requisition", vbInformation
                        Exit Sub
                    End If
                Else
                    '** ------------------------------------------------------------ **'
                    '**  Modified by Aiby ::  Dated: 12-Oct-2011                     **'
                    '**  if its Pachayat Login Previleges mush be checked Accoringly **'
                    '** ------------------------------------------------------------ **'
                    If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
                        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupChairPerson Then
                            GoTo PanchayatLog:
                        End If
                    End If
                    '** ------------------------------------------------------------ **'
                    '** End of Code Block Added                                      **'
                    '** ------------------------------------------------------------ **'
                    MsgBox "Letter Of Allotment Not authorized Yet..", vbInformation, "Saaankhya"
                    Exit Sub
                End If
            Else
                frmViewAllotmentLetter.ArrayIn = Array(CStr(val(vsGrid.TextMatrix(vsGrid.Row, 9))), mYearID)
                frmViewAllotmentLetter.Mode = 3
                 'Unload Me
                frmViewAllotmentLetter.Show vbModal
               
            End If
            Rec.Close
        End If
        'Call FillGrid
        Exit Sub
  
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdNew_Click()
        
        frmRequisition.RequisitionID = ""
        frmRequisition.PreviousYearMode = 0
        
        '**************UNAUTHORISED DRAWAL***************
        If mLoadMode = 10 Then
            frmRequisition.LoadMode = 10
        Else
            frmRequisition.LoadMode = 0
        End If
        '*************END********************************
        
        Unload Me
        frmRequisition.Show vbModal
        Fillgrid
    End Sub
    Private Sub ProceedingNumber()
        Dim mCnn        As New ADODB.Connection
        Dim msql        As String
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        
        gbSearchID = -1
        gbSearchStr = ""
        frmProceedings.chkEdit.Value = 0
        frmProceedings.Module = 130
        frmProceedings.Show vbModal
        If gbSearchID > 0 Then
            Dim objProceedings As New clsProceedings
            With objProceedings
                .ProceedingsID = gbSearchID
                .getProceedingsByID
                If .Used > 0 Then
                    MsgBox "This Proceedings already used", vbInformation
                    .ProceedingsID = -1
                Else
                    msql = "UPDATE  faProceedings SET tnyUsed=1 , intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 9))
                    msql = msql + " Where intProceedingsID= " & gbSearchID
                    objDb.ExecuteSP msql, , , , mCnn, adCmdText
                End If
            End With
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub
    
    Private Sub cmdProceedings_Click()
        
        Dim mCnn        As New ADODB.Connection
        Dim msql        As String
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mAuthAmt    As Variant
        Dim mProceedingsID As Integer
        Dim mYearID     As Variant
        '*********************************************************************************************'
        '                               Procedure to view the Proceedings                             '
        '*********************************************************************************************'
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If mPreviousYear = 1 Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        '-------------TO CHECK WHETHER PROCEEDINGS NO IS ALREADY LINKED WITH REQUISITIONS----------------------
        msql = "SELECT * FROM faProceedings WHERE intModuleID=130 AND intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 9))
        Rec.Open msql, mCnn
        If (Rec.EOF And Rec.BOF) Then
            If MsgBox(" Do you want to Link ProceedingsNo With Requisition?", vbYesNo, "Saankhya") = vbYes Then
                Call ProceedingNumber
            End If
        End If
        Rec.Close
        
        mAuthAmt = ""
        msql = "Select fltAuthorizedAmt From faAllotments Where intID = " & val(vsGrid.TextMatrix(vsGrid.Row, 9))
        Rec.Open msql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mAuthAmt = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
        End If
        Rec.Close
        If mAuthAmt = "" Then
            MsgBox "Please authorize the Allotment", vbInformation
            Exit Sub
        Else
            frmViewAllotmentLetter.ArrayIn = Array(CStr(val(vsGrid.TextMatrix(vsGrid.Row, 9))), CStr(mYearID))
            frmViewAllotmentLetter.Mode = 4
            frmViewAllotmentLetter.Show vbModal
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub CheckValidateProceedingNo()
    
    End Sub
    Private Sub cmdSearchIMPO_Click()
        frmSearchSubsidiaryAccountHeads.SubLedgerType = 1 ' 1=Implementing Officer
        frmSearchSubsidiaryAccountHeads.Show vbModal
        txtImpo.Text = Trim(gbSearchStr)
        txtImpo.Tag = gbSearchID
        If txtImpo.Tag <> -1 Then
            mImpOff = True
        End If
        Call Fillgrid
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    
    Private Sub cmdSearchSource_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        txtSource.SetFocus
    End Sub

    Private Sub cmdSearchSource_GotFocus()
        If gbSearchID > -1 And Trim(gbSearchStr) <> "" Then
            txtSource.Text = gbSearchStr
            txtSource.Tag = gbSearchID
            gbSearchID = -1
            gbSearchCode = ""
            gbSearchStr = ""
        Else
            txtSource.Text = ""
            txtSource.Tag = ""
        End If
        Call Fillgrid
    End Sub

    Private Sub cmdTreasuryBill_Click()
        Dim mCnn        As New ADODB.Connection
        Dim msql        As String
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mAuthAmt    As Variant
        Dim mYearID     As Variant
        '*********************************************************************************************'
        '                               Procedure to view the Treassury Bill                          '
        '*********************************************************************************************'
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If mPreviousYear = 1 Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        mAuthAmt = ""
        msql = "Select fltAuthorizedAmt From faAllotments Where intID = " & val(vsGrid.TextMatrix(vsGrid.Row, 9))
        Rec.Open msql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mAuthAmt = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
        End If
        Rec.Close
        If mAuthAmt = "" Then
            MsgBox "Please authorize the Allotment", vbInformation
            Exit Sub
        Else
            frmViewAllotmentLetter.ArrayIn = Array(CStr(val(vsGrid.TextMatrix(vsGrid.Row, 9))), CStr(mYearID)) 'Array(vsGrid.TextMatrix(vsGrid.Row, 9))
            If val(vsGrid.TextMatrix(vsGrid.Row, 11)) = 3 Then
                frmViewAllotmentLetter.Mode = 5 ' B-FUND
            Else
                If val(vsGrid.TextMatrix(vsGrid.Row, 16)) = 1 Then
                    frmViewAllotmentLetter.Mode = 9 ' :: TR59(C)
                Else
                    frmViewAllotmentLetter.Mode = 6 '::TR59(B)_59(C)
                End If
            End If
            frmViewAllotmentLetter.Show vbModal
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub dtpkrFromDate_CloseUp()
        txtFromDate.Text = CheckDateInMMM(dtPkrFromDate.Value)
    End Sub
    Private Sub dtpkrToDate_CloseUp()
        txtToDate.Text = CheckDateInMMM(dtPkrToDate.Value)
    End Sub
    Private Sub SynProjectMaster()
        
        Dim mCn As New ADODB.Connection
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        
        Dim Rec As New ADODB.Recordset
        Dim RecProj As New ADODB.Recordset
        
        Dim msql As String
        
        objDb.CreateNewConnection mCn, enuSourceString.Sulekha
        If mCn.State = 1 Then
            'mSQL = " SELECT  decProjectID, intYearID, chvDPCOrderNo, dtDPCOrderDate, intImplementingOfficerID, intMicroSectorID, vchApproverFullName, "
            'mSQL = mSQL + " vchApproverDesignation From ProjectDetails WHERE intYearID  = 2013"
          
            msql = " SELECT ProjectDetails.decProjectID,ProjectDetails.intYearID, nchApprovalNo chvDPCOrderNo, dtApprovaldate dtDPCOrderDate,"
            msql = msql + " intImplOfficerID intImplementingOfficerID, chvFullName vchApproverFullName, ProjectDetails.intYearID, chvDesignation vchApproverDesignation"
            msql = msql + " From ProjectDetails INNER JOIN SubjectCheckList ON SubjectCheckList.decProjectID = ProjectDetails.decProjectID"
            msql = msql + " Where ProjectDetails.intYearID = 2013 And SubjectCheckList.intYearID = 2013"

            
            Rec.Open msql, mCn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (Rec.BOF And Rec.EOF) Then
                msql = " SELECT  decProjectID, intYearID, chvDPCOrderNo, dtDPCOrderDate, intImplementingOfficerID, intMicroSectorID, vchApproverFullName, "
                msql = msql + " vchApproverDesignation From suProjectDetails WHERE intYearID  = 2013"
            
                objDb.SetConnection mCnn
                RecProj.Open msql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                If Not (RecProj.BOF And RecProj.EOF) Then ' CHECKING WHETHER ANY RECORDS IN DB_Finance.ProjectMaster
                    While Not Rec.EOF
                        'Debug.Print Rec!decProjectID
                        RecProj.Find "decProjectID =" & Rec!decProjectID, , , 1
                        If Not RecProj.EOF Then
                            
                            RecProj!chvDPCOrderNo = Rec!chvDPCOrderNo
                            RecProj!dtDPCOrderDate = Rec!dtDPCOrderDate
                            RecProj!intImplementingOfficerID = Rec!intImplementingOfficerID
                            'RecProj!intMicroSectorID = Rec!intMicroSectorID
                            RecProj!vchApproverFullName = Rec!vchApproverFullName
                            RecProj!vchApproverDesignation = Rec!vchApproverDesignation
                            
                            RecProj.Update
                        End If
                        Rec.MoveNext
                    Wend
                End If
            End If
            Rec.Close
        End If
        
        
    
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Call PreviousYearRequisitions
        Call Fillgrid
        
    End Sub
    Private Sub Form_Load()
        Call FormInitialize

        'Call FillGrid
        vsGrid.SelectionMode = flexSelectionByRow
        'Note:-For Only Operator Type User Can Access New Command
        If gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
            cmdNew.Enabled = True
        Else
            cmdNew.Enabled = False
        End If
        txtFromDate.Text = DdMmmYy(gbStartingDate) '(DateAdd("d", -30, gbTransactionDate))
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        If gbUserID <> 4 Then
            'cmdNew.Visible = False
        End If
        
        
        '============================================================================================'
        ' SYNC PROJECT MASTER FROM SULEKHA TO SAANKHYA - MODIFED ON FEB,2014
        '============================================================================================'
        
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim msql As String
        Dim objDb As New clsDB
        
        If gbLBPanchayat Then
            msql = "SELECT  ISNULL(CONVERT(INTEGER, vcbSignature),0) intFlag FROM faDBSubVersions WHERE  vchDBSubVersionKey = '0.13'"
        Else
            msql = "SELECT  ISNULL(CONVERT(INTEGER, vcbSignature),0) intFlag FROM faDBSubVersions WHERE  vchDBSubVersionKey = '0.15'"
        End If
        objDb.SetConnection mCnn
        If mCnn.State = 1 Then
            Rec.Open msql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!intFlag = 0 Then
                    Call SynProjectMaster
                    If gbLBPanchayat Then
                        msql = "Update faDBSubVersions SET vcbSignature = CONVERT(varbinary,1) WHERE  vchDBSubVersionKey = '0.13'"
                    Else
                        msql = "Update faDBSubVersions SET vcbSignature = CONVERT(varbinary,1) WHERE  vchDBSubVersionKey = '0.15'"
                    End If
                    mCnn.Execute msql
                End If
            End If
            Rec.Close
        End If
        
        '============================================================================================'
        ' END OF BLOCK - SYNC PROJECT MASTER
        '============================================================================================'
        
        Dim mLastDate As Variant
        Dim mFromDate As Variant
        Dim mToDate As Variant
        mLastDate = GetSetting("SaankhyaDE", "App", "ListReqLastDate")
      
        If mLastDate <> gbTransactionDate Then
            'SaveSetting "SaankhyaDE", "App", "Path", CStr(App.Path)
            SaveSetting "SaankhyaDE", "App", "ListReqLastDate", gbTransactionDate
        Else
            mFromDate = GetSetting("SaankhyaDE", "App", "ListReqFromDate")
            mToDate = GetSetting("SaankhyaDE", "App", "ListReqToDate")
            If mPreviousYear Then
                If IsDate(mFromDate) Then
                    txtFromDate.Text = DdMmmYy(CDate(mFromDate))
                Else
                    txtFromDate.Text = DdMmmYy(gbStartingDate)
                End If
                If IsDate(mToDate) Then
                    txtToDate.Text = DdMmmYy(CDate(mToDate))
                Else
                    txtToDate.Text = DdMmmYy(gbTransactionDate)
                End If
            Else
                If Not (mFromDate >= gbStartingDate And mFromDate <= gbEndingDate) Then
                    txtFromDate.Text = DdMmmYy(gbStartingDate)
                End If
                If Not (mToDate >= gbStartingDate And mToDate <= gbEndingDate) Then
                    txtToDate.Text = DdMmmYy(gbTransactionDate)
                End If
            End If
                
        End If
        
    End Sub
    Public Property Let IMPO(mData As Integer)
        intIMPO = mData
    End Property
    Public Property Get IMPO() As Integer
        IMPO = intIMPO
    End Property
  
    Private Sub vsGrid_DblClick()
        Dim mYearID       As Variant
        
        If vsGrid.Row > 0 Then
            If mPreviousYear = 1 Then
                mYearID = gbFinancialYearID - 1
            Else
                mYearID = gbFinancialYearID
            End If
            
            If mLoadMode = 10 Then
                frmViewAllotmentLetter.LoadMode = 10
            End If
            
            If vsGrid.TextMatrix(vsGrid.Row, 9) <> "" Then
                frmViewAllotmentLetter.ArrayIn = Array(CStr(val(vsGrid.TextMatrix(vsGrid.Row, 9))), CStr(mYearID))
                If vsGrid.TextMatrix(vsGrid.Row, 10) <> "" Then
                    frmViewAllotmentLetter.Mode = 2
                Else
                    frmViewAllotmentLetter.Mode = 1
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        frmViewAllotmentLetter.cmdVerify.Visible = True
                    Else
                        If vsGrid.TextMatrix(vsGrid.Row, 10) <> "" Then
                            frmViewAllotmentLetter.cmdVerify.Visible = False
                        Else
                            frmViewAllotmentLetter.cmdVerify.Visible = True
                        End If
                    End If
                End If
                Unload Me
                'Me.Hide
                frmViewAllotmentLetter.Show vbModal
            End If
            Call Fillgrid
        End If
    End Sub

    Private Sub PreviousYearRequisitions()
        Dim msql        As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mStartDate  As Date
        Dim mEndDate    As Date
        Dim mYearID     As Variant
        
            ' BLOCKED BY AIBY on 10 APRIL
            'If objDb.CreateNewConnection(mcnn, enuSourceString.Saankhya) Then
            '            mSql = "Select * From faFinancialYear Where tinCurrentFinancialYearFlag=1"
            '            Set Rec = objDb.ExecuteSP(mSql, , , , mcnn, adCmdText)
            '            If Not (Rec.EOF Or Rec.BOF) Then
            '                mYearID = Rec!intFinancialYear
            '                mStartdate = Rec!dtStartingDate
            '                mEnddate = Rec!dtEndingDate
            '            End If
            '            Rec.Close
            
            mYearID = gbCurrentPeriodID
            mStartDate = gbStartingDate
            mEndDate = gbEndingDate

            If CDate(txtFromDate.Text) < CDate(mStartDate) And CDate(txtToDate.Text) < CDate(mEndDate) Then
                mPreviousYear = 1
            Else
                mPreviousYear = 0
            End If
            'End If
    End Sub
    Public Property Let PreviousYear(ByVal mData As Integer)
        mPreviousYear = mData
    End Property
    Public Property Let LoadMode(mData As Integer)
        mLoadMode = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = mLoadMode
    End Property
