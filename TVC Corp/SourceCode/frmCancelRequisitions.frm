VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCancelRequisitions 
   Caption         =   "Cancel Requisitions"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   Icon            =   "frmCancelRequisitions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   15
      TabIndex        =   11
      Top             =   3480
      Width           =   7935
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel Request"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4755
         TabIndex        =   14
         Top             =   225
         Visible         =   0   'False
         Width           =   1575
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   7080
         Top             =   1080
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton cmdRequest 
         Caption         =   "Request"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdApprove 
         Caption         =   "Approve"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2970
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   9375
      Begin VB.CommandButton cmdSearchReason 
         Caption         =   "..."
         Height          =   285
         Left            =   5280
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdSearchRequisitionNo 
         Caption         =   "..."
         Height          =   285
         Left            =   5280
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtReason 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtRequisitionNo 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason *"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Requisition No *"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   9345
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "This form records the details of Requisitions to be Cancelled"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   9375
      End
   End
End
Attribute VB_Name = "frmCancelRequisitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mPreviousYearMode As Integer
    Dim mPreviousYearRequestID As Variant
    
    Private Sub GetPreviousYearCancelRequestDetails()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
             

        If mPreviousYearRequestID > 0 Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                mSql = "Select * from faPendingTaskRequest Where intRequestID= " & mPreviousYearRequestID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtRequisitionNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    txtRequisitionNo.Tag = IIf(IsNull(Rec!intKeyID), "", Rec!intKeyID)
                    txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                End If
                Rec.Close
            End If
            txtRequisitionNo.Enabled = False
            txtRemarks.Enabled = False
        End If
    End Sub
    
    Private Sub cmdApprove_Click()
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mCnnSulekha    As New ADODB.Connection
        Dim mAllotmentNo As Variant
            
            If Not (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                MsgBox "Connection To Plan[Sulekha] Module not found", vbCritical
                Exit Sub
            End If
            
            If CheckPayOrderExist(val(txtRequisitionNo.Tag)) = True Then
                MsgBox "Payment Order Exists for the Requisition.Please Cancel the Payment Order", vbInformation
                Exit Sub
            End If
            
            
            If objdb.SetConnection(mCnn) Then
                mSql = "Update faRequisitionRequest set tnyStatus = 1 where intRequisitionID = " & txtRequisitionNo.Tag & "  "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = "Update faAllotments set tnyStatus = 2 where intID = " & txtRequisitionNo.Tag & "  "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = "Update faPendingTaskRequest set tnyStatus = 8 where intTaskID = 10 AND intKeyID = " & txtRequisitionNo.Tag & "  "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = "Select vchAllotmentNo from faAllotments  Where intID=" & txtRequisitionNo.Tag & "  "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    mAllotmentNo = IIf(IsNull(Rec!vchAllotmentNo), 0, Rec!vchAllotmentNo)
                End If
                Rec.Close
            End If
            
            If mCnnSulekha.State Then mCnnSulekha.Close
            If (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                mSql = "Update RequisitionDetails set tnyCancel = 2,tnyTransfer=0 where  intReqID = " & val(txtRequisitionNo.Tag) & "  "
                objdb.ExecuteSP mSql, , , , mCnnSulekha, adCmdText
            Else
                MsgBox "Connection to Sulekha doesnot exists", vbInformation, "Saankhya"
                Exit Sub
            End If
            cmdCancel.Enabled = False
            cmdApprove.Enabled = False
            
    End Sub

Private Sub cmdCancel_Click()
    If val(cmdCancel.Tag) = 1 Then
        MsgBox "This request is already Approved!", vbInformation
        Exit Sub
    End If
    If MsgBox("Do you want to remove the Cancellation Request ?!", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
        Dim mID As Integer
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        mID = IIf(txtRequisitionNo.Tag = "", -1, val(txtRequisitionNo.Tag))
        If mID > 0 Then
            mSql = "Delete FROM  faRequisitionRequest Where intRequisitionID = " & mID
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            cmdRequest.Enabled = False
            cmdCancel.Enabled = False
            MsgBox "Request for cancellation is removed!", vbInformation
            Unload Me
        End If
    End If
End Sub

    Private Sub cmdRequest_Click()
        Dim mID     As Variant
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim arrInput    As Variant
        Dim mStatus     As Integer
        Dim Reqn        As uRequisition
        Dim mSql As String
        
        If txtRequisitionNo.Text = "" Then
            MsgBox "Please select a Requsition ", vbInformation
            Exit Sub
        End If
        If txtReason.Text = "" Then
            MsgBox "Please select a Reason ", vbInformation
            Exit Sub
        End If
        If txtRemarks.Text = "" Then
            MsgBox "Please give Remarks", vbInformation
            Exit Sub
        End If
        
        If CheckPayOrderExist(val(txtRequisitionNo.Tag)) = True Then
            MsgBox "Payment Order Exists for the Requisition.Please Cancel the Payment Order", vbInformation
            txtRequisitionNo.Text = ""
            txtRequisitionNo.Tag = ""
            txtReason.Text = ""
            txtRemarks.Text = ""
            Exit Sub
        End If
        
        mID = IIf(txtRequisitionNo.Tag = "", -1, val(txtRequisitionNo.Tag))
        mStatus = CheckRequisitionRequestExist(val(txtRequisitionNo.Tag))
        If mStatus = 1 Then
            MsgBox "Request Already exists,waiting for Approval", vbInformation
            Exit Sub
        ElseIf mStatus = 2 Then
            MsgBox "Cancellation Request is Already Approved", vbInformation
            Exit Sub
        '------------------------------'
        'Blocked By Aiby: 10-Nov-2011
            'Else
            ' mID = -1
        '------------------------------'
        End If
        
        arrInput = Array(mID, _
                        Trim(txtRequisitionNo.Text), _
                        val(txtReason.Tag), _
                        Trim(txtRemarks.Text) _
                        )
        'Changed By Aiby : 10-Nov-2011
        'objdb.ExecuteSP "spSaveRequisitionRequest", arrInput, , , mCnn, adCmdStoredProc
        objdb.ExecuteSP "spSaveCancelRequisition", arrInput, , , mCnn, adCmdStoredProc
       
        '--------------------------------------------------------------------------------------'
        'Blocked by Aiby on 10-Nov-2011
        'mSql = "Update faAllotments set tnyStatus=1 where intID = " & txtRequisitionNo.Tag & "  "
        'objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        '--------------------------------------------------------------------------------------'
        MsgBox "Request saved Successfully", vbInformation
        Unload Me
               
    End Sub

'''    Private Sub cmdSearchForwardSeat_Click()
'''        Dim objDB   As New clsDB
'''        Dim mCnn    As New ADODB.Connection
'''        Dim Rec     As New ADODB.Recordset
'''        Dim mCnt    As Integer
'''        Dim msql    As String
'''
'''        txtForwardSeat.Tag = ""
'''        txtForwardSeat.Text = ""
'''        msql = "Select chvSeatTitle, numSeatID From GL_Seats Where intGroupID in (5,6) And intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
'''        frmSearchSeat.SQLString = msql
'''        frmSearchSeat.Show vbModal
'''        If gbSearchID > -1 Then
'''            txtForwardSeat.Tag = gbSearchID
'''            txtForwardSeat.Text = gbSearchStr
'''        Else
'''            gbSearchID = -1
'''            gbSearchStr = ""
'''        End If
'''    End Sub

    Private Sub cmdSearchReason_Click()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objdb As New clsDB
            
        On Error GoTo err:
            If txtRequisitionNo.Text = "" Then
                MsgBox "Please Select a Requisition Before Giving Reason", vbInformation
                Exit Sub
            End If
            frmSearchMasters.SQLQry = "Select intReasonID, vchReason From faReasons Where intType=90 "
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Show vbModal
            txtReason.Text = gbSearchStr
            txtReason.Tag = gbSearchID
'''            If objDB.SetConnection(mCnn) Then
'''                mSQL = " Select intCategory from faReasons Where intReasonID=" & gbSearchID
'''                Rec.Open mSQL, mCnn
'''                If Not (Rec.EOF Or Rec.BOF) Then
'''                    cmdSearchReason.Tag = Rec!intCategory
'''                End If
'''            End If
            gbSearchID = -1
            gbSearchStr = ""
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Sub cmdSearchRequisitionNo_Click()
        frmSearchRequesition.PreviousYearMode = 0
        frmSearchRequesition.Show vbModal
        txtRequisitionNo.Text = gbSearchStr
        txtRequisitionNo.Tag = gbSearchID
        
        If CheckPayOrderExist(val(txtRequisitionNo.Tag)) = True Then
            MsgBox "Payment Order Exists for the Requisition.Please Cancel the Payment Order", vbInformation
            txtRequisitionNo.Text = ""
            txtRequisitionNo.Tag = ""
        End If
        Call CheckTransferCredit(val(txtRequisitionNo.Tag))
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Sub CheckTransferCredit(ByVal mReqID As Double)
        Dim mSql                As String
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim objdb               As New clsDB
        Dim mSqlChild           As String
        Dim RecChild            As New ADODB.Recordset
        Dim mAccountHeadID      As Variant
        
        mSql = " SELECT ISNULL(numProjectNo,0) numProjectID,intKeyID1 intAccountHeadID"
        mSql = mSql + " From faPayOrder"
        mSql = mSql + " INNER JOIN faVouchers ON faVouchers.intVoucherID=faPayOrder.intVoucherID"
        mSql = mSql + " WHERE intAllotmentID = " & mReqID
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "The Connection to Saankhya not Present", vbCritical
            Exit Sub
        End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            If Rec!numProjectID <> 0 Then
                mAccountHeadID = Rec!intAccountHeadID
                If gbLBPanchayat = 1 Then
                    mSqlChild = "SELECT intKeyID1 FROM faVouchers "
                    mSqlChild = mSqlChild + " INNER JOIN faVoucherChild ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
                    mSqlChild = mSqlChild + " WHERE intTransactionTypeID=4010"
                    
                Else
                    mSqlChild = "SELECT intKeyID1 FROM faVouchers "
                    mSqlChild = mSqlChild + " INNER JOIN faVoucherChild ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
                    mSqlChild = mSqlChild + " WHERE intTransactionTypeID=4006"
                End If
                RecChild.Open mSqlChild, mCnn
                If Not (RecChild.EOF And RecChild.BOF) Then
                    While Not RecChild.EOF
                        If RecChild!intKeyID1 = mAccountHeadID Then
                            MsgBox "Transfer Credit Already Done.No Cancellation is possible", vbInformation
                            txtRequisitionNo.Text = ""
                            txtRequisitionNo.Tag = ""
                            Exit Sub
                        End If
                        RecChild.MoveNext
                    Wend
                    
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = (Screen.Width - Me.Width) / 2
    End Sub
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
             cmdApprove.Enabled = False
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdRequest.Enabled = False
        End If
        If mPreviousYearMode Then   'FOR PREVIOUS YEAR CANCEL REQUISITIONS
            Call GetPreviousYearCancelRequestDetails
        End If

    End Sub
    Private Function CheckPayOrderExist(ReqID As Long) As Boolean
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objdb As New clsDB
        If objdb.SetConnection(mCnn) Then
            mSql = "SELECT * FROM faPayOrder WHERE intAllotmentID=" & ReqID & " AND ISNULL(tnyCancelled,0)<>1 "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                CheckPayOrderExist = True
            Else
                CheckPayOrderExist = False
            End If
         End If
    End Function
    
    Private Function CheckRequisitionRequestExist(ByVal RqID As Double) As Integer
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            If objdb.SetConnection(mCnn) Then
            
                ' Changed by Aiby : 10-Nov-2011
                '   mSql = " Select tnyStatus from faAllotments where  intID= " & RqID & " "
                '   mSql = mSql + " And tnyStatus in (1,2)"
                '    Rec.Open mSql, mCnn
                '    If Not (Rec.EOF Or Rec.BOF) Then
                '        If Rec!tnyStatus = 1 Then      'Requested
                '            CheckRequisitionRequestExist = 1
                '        ElseIf (Rec!tnyStatus = 2) Then  ' Approved/Cancelled
                '            CheckRequisitionRequestExist = 2
                '    '''                    Else
                '    '''                        CheckRequisitionRequestExist = 0
                '            Exit Function
                '        End If
                '    Else
                '        CheckRequisitionRequestExist = 0  'NOT EXISTS IN THE TABLE
                '    End If
                '    Rec.Close
                
                mSql = " Select * From faRequisitionRequest WHERE intRequisitionID = " & RqID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyStatus = 0 Then      'Requested
                        CheckRequisitionRequestExist = 1
                    ElseIf (Rec!tnyStatus = 1) Then  ' Approved
                        CheckRequisitionRequestExist = 2
                    End If
                Else
                    CheckRequisitionRequestExist = 0  'NOT EXISTS IN THE TABLE
                End If
                Rec.Close
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact Your System Administrator"
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function

    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearRequestID(mData As Integer)
        mPreviousYearRequestID = mData
    End Property

