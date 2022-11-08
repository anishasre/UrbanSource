VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmListOfInterruptedEditRequests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Interrupted Edit Request"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReject 
      Caption         =   "&Reject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelRequest 
      Caption         =   "Cancel Request"
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
      Left            =   1455
      TabIndex        =   3
      Top             =   6150
      Width           =   1395
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "&Approve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
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
      Left            =   45
      TabIndex        =   1
      Top             =   6150
      Width           =   1395
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6030
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   11730
      _cx             =   20690
      _cy             =   10636
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfInterruptedEditRequests.frx":0000
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
Attribute VB_Name = "frmListOfInterruptedEditRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mCount  As Integer
    '*********************************************************************************************'
    '               Form to list all the Interrupt Receipt Edit requests                          '
    '*********************************************************************************************'
    
    Private Function CheckInterruptReceiptRequestStatus() As Boolean
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mStatus = ""
        mSql = "Select tnyStatus From faInterruptedRequests"
        'If gbUserTypeID = 3 Then
        mSql = mSql + " Where numUserID =" & gbUserID
        mSql = mSql + " And intCounterID =" & gbCounterID
        mSql = mSql + " And intTypeID = 1"
        'End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
        End If
        Rec.Close
        mCnn.Close
        If mStatus <> "" Then
            If mStatus = 2 Then
                CheckInterruptReceiptRequestStatus = True
            Else
                CheckInterruptReceiptRequestStatus = False
            End If
        Else
            CheckInterruptReceiptRequestStatus = False
        End If
        Exit Function
err:
        MsgBox err.Description
    End Function
    Private Function GetUserName(mUserID As Variant) As String
        Dim mCnnUserName    As New ADODB.Connection
        Dim RecUserName     As New ADODB.Recordset
        Dim objUserName     As New clsDB
        Dim mSQLUserName    As String
        
        '*********************************************************************************************'
        '                       Function to get User Name from DB_Masters                             '
        '*********************************************************************************************'
        On Error GoTo err
        objUserName.CreateNewConnection mCnnUserName, enuSourceString.DBMaster
        
        mSQLUserName = "Select * From GM_User"
        mSQLUserName = mSQLUserName + " Where numUserID = " & mUserID
        RecUserName.Open mSQLUserName, mCnnUserName
        If Not (RecUserName.EOF And RecUserName.BOF) Then
            GetUserName = IIf(IsNull(RecUserName!vchEmpName), "", RecUserName!vchEmpName)
        End If
        RecUserName.Close
        Exit Function
err:
        MsgBox err.Description
    End Function
    
    Private Sub FillvsGrid()
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim mRowCount   As Integer
        
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        mCount = 0
        mSql = "Select *,faInterruptedRequests.tnyStatus As Status From faInterruptedRequests"
        mSql = mSql + " Inner Join faVouchers On faInterruptedRequests.intVoucherID = faVouchers.intVoucherID"
        mSql = mSql + " Inner Join faCounters On faInterruptedRequests.intCounterID = faCounters.intCounterID"
        mSql = mSql + " Inner Join faCancelReason On faInterruptedRequests.intReasonID = faCancelReason.intCancelID"
        mSql = mSql + " Where intTypeID = 3"
        If gbUserTypeID = 3 Then
            mSql = mSql + " And numUserID =" & gbUserID
        End If
        'If gbUserTypeID <> 3 Then
        mSql = mSql + " And faInterruptedRequests.tnyStatus <> 0"
        'End If
        mSql = mSql + " Order By dtRequestDate"
        Rec.Open mSql, mCnn
        While Not Rec.EOF
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.TextMatrix(mRowCount, 0) = GetUserName(Rec!numUserID)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate)
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo) & IIf(IsNull(Rec!vchDoorNoP3), "", "-" & Rec!vchDoorNoP3)
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!vchCancelReason), "", Rec!vchCancelReason)
            If Rec!Status = 2 Then
                vsGrid.Cell(flexcpChecked, mRowCount, 5) = True
            Else
                vsGrid.Cell(flexcpChecked, mRowCount, 5) = False
            End If
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intCounterID), "", Rec!intCounterID)
            vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intReasonID), "", Rec!intReasonID)
            mRowCount = mRowCount + 1
            mCount = mCount + 1
            Rec.MoveNext
        Wend
        Rec.Close
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdCancelRequest_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
        
        '*********************************************************************************************'
        '               Procedure to Cancel the request for Interrupt Receipt Edit                    '
        '*********************************************************************************************'
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If vsGrid.Row > 0 Then
            If vsGrid.TextMatrix(vsGrid.Row, 3) <> "" Then
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 5) = 2 Then
                    mSql = "Update faInterruptedRequests"
                    mSql = mSql + " Set tnyStatus = 0"
                    mSql = mSql + " Where numUserID = " & vsGrid.TextMatrix(vsGrid.Row, 6)   '
                    mSql = mSql + " And intCounterID = " & vsGrid.TextMatrix(vsGrid.Row, 7)  '
                    'mSql = mSql + " And intVoucherNo = " & vsGrid.TextMatrix(vsGrid.Row, 3) 'CHANGED BY AIBY ON 10-NOV-2011
                    mSql = mSql + " And intVoucherID = " & val(vsGrid.TextMatrix(vsGrid.Row, 8))   '
                    mCnn.Execute mSql
                    MsgBox "Request cancelled successfully", vbInformation
                Else
                    MsgBox "Can't cancel the Request(Already approved)"
                End If
            End If
        Else
            MsgBox "Please select a Request", vbInformation
            Exit Sub
        End If
        Call FillvsGrid
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdNew_Click()
        frmInterruptedEditRequest.Show vbModal
        Call FillvsGrid
    End Sub

'''    Private Sub cmdReject_Click()
'''        Dim mRowCount   As Integer
'''
'''        For mRowCount = 1 To mCount
'''            If vsGrid.Cell(flexcpChecked, mRowCount, 5) = vbChecked Then
'''                frmReject.Mode = 5
'''                frmReject.RequestTypeID = vsGrid.TextMatrix(mRowCount, 8)
'''                frmReject.Show vbModal
'''                cmdReject.Enabled = False
'''                cmdVerify.Enabled = False
'''            End If
'''        Next
'''    End Sub

    Private Sub cmdVerify_Click()
        Dim mCnn        As New ADODB.Connection
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mSanction   As Integer
        Dim mRowCount   As Integer
        
        '*********************************************************************************************'
        '                       Procedure to approve the Interrupt Edit Request                       '
        '*********************************************************************************************'
        'On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If vsGrid.Rows > 1 Then
            If gbUserTypeID <> 3 Then
                For mRowCount = 1 To mCount
                    If vsGrid.Cell(flexcpChecked, mRowCount, 5) = vbChecked Then
                        mSanction = 2
                    Else
                        mSanction = 1
                    End If
                    mSql = "Update faInterruptedRequests"
                    mSql = mSql + " Set tnyStatus =" & mSanction
                    mSql = mSql + " Where numUserID=" & vsGrid.TextMatrix(mRowCount, 6)
                    mSql = mSql + " And intCounterID = " & vsGrid.TextMatrix(mRowCount, 7)
                    'mSql = mSql + " And intVoucherNo = " & vsGrid.TextMatrix(mRowCount, 3)
                    mSql = mSql + " And intVoucherID = " & val(vsGrid.TextMatrix(mRowCount, 8))
                    mSql = mSql + " And intTypeID = 3"
                    mCnn.Execute mSql
                Next
            End If
            MsgBox "Successfully Saved", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Me.Width = 11940
        Me.Height = 7155
    End Sub

    Private Sub Form_Load()
        If gbUserTypeID = 3 Then
            cmdNew.Visible = True
            cmdVerify.Visible = False
'            cmdReject.Visible = False
            cmdCancelRequest.Visible = True
        Else
            cmdNew.Visible = False
            cmdVerify.Visible = True
'            cmdReject.Visible = True
            cmdCancelRequest.Visible = False
        End If
        Call FillvsGrid
    End Sub
    
    Private Sub vsGrid_Click()
        If gbUserTypeID <> 3 Then
            If vsGrid.Col = 5 Then
                vsGrid.Editable = flexEDKbdMouse
            Else
                vsGrid.Editable = flexEDNone
            End If
        End If
    End Sub

    Private Sub vsGrid_DblClick()
        Dim aryIn As Variant
        
        On Error GoTo err
        If gbUserTypeID = 3 Then
            If vsGrid.Row > 0 Then
                If vsGrid.TextMatrix(vsGrid.Row, 3) <> "" Then
                    If vsGrid.Cell(flexcpChecked, vsGrid.Row, 5) = 1 Then
                        If CheckInterruptReceiptRequestStatus Then
                            frmReceiptsCounter.InterruptEditMode = True
                            frmReceiptsCounter.DisplayReceiptDetails (vsGrid.TextMatrix(vsGrid.Row, 8))
                    Exit Sub
                        Else
                            MsgBox "You are not in Interrupt Receipt Mode", vbInformation
                            Exit Sub
                        End If
                    Else
                        MsgBox "Request pending for Approval", vbInformation
                    End If
                End If
            End If
        Else
            If vsGrid.Row > 0 Then
                If vsGrid.TextMatrix(vsGrid.Row, 8) <> "" Then
                    aryIn = Array(vsGrid.TextMatrix(vsGrid.Row, 8))
                    frmViewVoucher.ArrayIn = aryIn
                    frmViewVoucher.FormName = "frmInterruptReceipt"
                    frmViewVoucher.Show vbModal
                End If
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
