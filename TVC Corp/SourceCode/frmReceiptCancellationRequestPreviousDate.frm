VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReceiptCancellationRequestPreviousDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Canceallation Request - (Previous Date)"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11040
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11040
      TabIndex        =   23
      Top             =   0
      Width           =   11040
   End
   Begin VB.Frame frme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1.11000e5
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3210
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2370
         Width           =   4965
      End
      Begin VB.CommandButton cmdSearchReceipt 
         Caption         =   "..."
         Height          =   315
         Left            =   5415
         TabIndex        =   36
         Top             =   1125
         Width           =   300
      End
      Begin VB.TextBox txtReceiptsNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3225
         TabIndex        =   35
         Top             =   1125
         Width           =   2175
      End
      Begin VB.CommandButton cmdProceedings 
         Caption         =   "..."
         Height          =   330
         Left            =   5415
         TabIndex        =   34
         Top             =   285
         Width           =   300
      End
      Begin VB.TextBox txtProceedingsNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3225
         MaxLength       =   50
         TabIndex        =   1
         Top             =   300
         Width           =   2175
      End
      Begin VB.TextBox txtProceedingsDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3225
         TabIndex        =   2
         Top             =   660
         Width           =   2175
      End
      Begin VB.TextBox txtRequestedBy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3225
         TabIndex        =   7
         Top             =   3915
         Width           =   4965
      End
      Begin VB.TextBox txtRequestedByDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3225
         TabIndex        =   8
         Top             =   4260
         Width           =   2175
      End
      Begin VB.ComboBox cmbReason 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3225
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2010
         Width           =   4965
      End
      Begin VB.TextBox txtStationaryNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3210
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3315
         Width           =   2175
      End
      Begin VB.TextBox txtStationaryCount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3210
         TabIndex        =   5
         Top             =   2970
         Width           =   2175
      End
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   11490
         Top             =   6735
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   30
         TabIndex        =   30
         Top             =   5265
         Width           =   12120
      End
      Begin VB.CommandButton cmdRequest 
         Caption         =   "&Request"
         Height          =   465
         Left            =   3210
         TabIndex        =   9
         Top             =   4635
         Width           =   2190
      End
      Begin VB.Frame Frame2 
         Caption         =   "Secretary"
         Height          =   1680
         Left            =   5355
         TabIndex        =   27
         Top             =   5475
         Width           =   5160
         Begin VB.CommandButton cmdSecondtAuthorize 
            Caption         =   "Authorize"
            Height          =   435
            Left            =   1515
            TabIndex        =   15
            Top             =   1110
            Width           =   3345
         End
         Begin VB.TextBox txtSecondtAuthorizedBy 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1530
            TabIndex        =   13
            Top             =   450
            Width           =   3390
         End
         Begin VB.TextBox txtSecondtAuthorizedDate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1530
            TabIndex        =   14
            Top             =   765
            Width           =   1560
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Authorized By   :"
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
            Left            =   135
            TabIndex        =   29
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Date   :"
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
            Left            =   855
            TabIndex        =   28
            Top             =   900
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Accounts Officer"
         Height          =   1680
         Left            =   165
         TabIndex        =   24
         Top             =   5475
         Width           =   5160
         Begin VB.CommandButton cmdFirstAuthorize 
            Caption         =   "Authorize"
            Height          =   435
            Left            =   1575
            TabIndex        =   12
            Top             =   1110
            Width           =   3345
         End
         Begin VB.TextBox txtFirstAuthorizedBy 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1575
            TabIndex        =   10
            Top             =   420
            Width           =   3390
         End
         Begin VB.TextBox txtFirstAuthorizedDate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1575
            TabIndex        =   11
            Top             =   735
            Width           =   1680
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Authorized By  :"
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
            Left            =   180
            TabIndex        =   26
            Top             =   465
            Width           =   1200
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Date   : "
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
            Left            =   780
            TabIndex        =   25
            Top             =   810
            Width           =   600
         End
      End
      Begin VB.TextBox txtReceiptsDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3225
         TabIndex        =   3
         Top             =   1485
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   690
         X2              =   8340
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2385
         TabIndex        =   33
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stationary Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1650
         TabIndex        =   32
         Top             =   3345
         Width           =   1530
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stationary Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1815
         TabIndex        =   31
         Top             =   3030
         Width           =   1365
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2820
         TabIndex        =   22
         Top             =   4290
         Width           =   360
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   21
         Top             =   3825
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2565
         TabIndex        =   20
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   19
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2310
         TabIndex        =   18
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proceedings Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1725
         TabIndex        =   17
         Top             =   705
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proceedings No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1875
         TabIndex        =   16
         Top             =   360
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmReceiptCancellationRequestPreviousDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Public mVoucherID     As Long
    Public crtStatus      As Integer
     Public mTransactionTypeID As Variant
    Private Sub cmbReason_Click()
        If cmbReason.ListIndex < 0 Then
            Exit Sub
        Else
            txtStationaryCount.Text = 1
            txtStationaryCount.Enabled = False
            txtStationaryNo.Enabled = True
        End If
    End Sub
    Private Sub cmdFirstAuthorize_Click()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
       
        If objdb.SetConnection(mCnn) Then
            mSql = "Update faReceiptCancellationRequest Set tnyStatusAO = 1, numAuthorisedByAO = " & gbUserID & " ,dtAuthorisationDateAO= '" & CheckDateInMMM(txtFirstAuthorizedDate.Text) & "' Where intID = " & val(txtProceedingsNo.Tag)
            Rec.Open mSql, mCnn
            If gbLBPanchayat = 1 Then
                mSql = "Update faReceiptCancellationRequest Set tnyStatusSec = 1, numAuthorisedBySec = " & gbUserID & " ,dtAuthorisationDateSec= '" & CheckDateInMMM(txtSecondtAuthorizedDate.Text) & "' Where intID = " & val(txtProceedingsNo.Tag)
                Rec.Open mSql, mCnn
                mSql = "Update faVouchers Set tnysync=Null,tnyStatus = 4, tnyCancelFlag = 1 Where intVoucherID = " & mVoucherID & " "
                mCnn.Execute mSql
                mSql = "Update faTransactions Set tnysync=Null,tnyStatus = 4 Where intVoucherID = " & mVoucherID & ""
                mCnn.Execute mSql
                mSql = "Update faCancelledVouchers Set tnyApproveStatus = 1, tinType=9 Where intVoucherID = " & mVoucherID & " "
                mCnn.Execute mSql
            End If
            cmdFirstAuthorize.Enabled = False
        Else
            MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
        End If
        If CancelReceipt = True Then
            If mTransactionTypeID = gbTransactionTypePTax Then
                If gbFetchDemandFromWeb = 1 Then
                        PTaxWebDemand (mVoucherID)
                Else
                   Call CancelPropertyTax(CDbl(txtReceiptsNo.Text), mVoucherID)
                End If
            ElseIf mTransactionTypeID = gbTransactionTypeDandO Then
                If gbLinkWithDandOWeb = 1 Then
                        CancelDAndODemand (mVoucherID)
                End If
            End If
            MsgBox " The request for Receipt Cancellation saved successfully", vbInformation
        
        End If
        Call FormInitialize
        Me.Hide
        frmListOfReceiptCancellationRequest.Show
    End Sub
    'Cancel D&O Demand in Sanchaya
'Created by Syalima S On Jan 2018
Private Function CancelDAndODemand(ByVal mVoucherID As Double) As Boolean
        On Error GoTo err:
        Dim objSOAP             As Variant
        Dim mArrOutDemandRecpt As Variant
        Dim flagSankhya As Integer
        Dim mUrl As String
        Dim arrInput As String
        Dim mArrOut As Variant
        Dim Rec As New ADODB.Recordset
        Dim mDemandID As Variant
        
           Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
            If Not (Rec.EOF Or Rec.BOF) Then
                mDemandID = Rec!numDemandID
            End If
            
           Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
           
            mUrl = gbDefaultUrl
            On Error Resume Next
            objSOAP.MSSoapInit (mUrl + "?WSDL")
           
            'arrInput = Array(mDemandID, gbLocalBodyID, vsGrid.TextMatrix(8, 1))
            arrInput = mDemandID & "#" & gbLocalBodyID & "#" & mVoucherID
            mArrOut = (objSOAP.savecancellreceipt(arrInput))
            CancelDAndODemand = True
err:
        MsgBox (Error$)
    End Function
    Private Sub cmdFirstAuthorize_GotFocus()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        If objdb.SetConnection(mCnn) Then
            mSql = " SELECT  faReceiptCancellationRequest.intID FROM faReceiptCancellationRequest INNER JOIN"
            mSql = mSql + "  faVouchers ON faReceiptCancellationRequest.intVoucherID = faVouchers.intVoucherID"
            mSql = mSql + " WHERE faVouchers.intVoucherNo =  " & txtReceiptsNo.Text & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!intID <> "" Then
                     txtProceedingsNo.Tag = (Rec!intID)
                End If
            End If
            Rec.Close
        Else
            MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
        End If
    End Sub

    Private Sub cmdProceedings_Click()
    Dim mSql    As String
    Dim objdb   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
        gbSearchID = -1
        gbSearchStr = ""
        frmProceedings.chkEdit.Value = 0
        frmProceedings.Module = 95
        frmProceedings.Show vbModal
        If gbSearchID > 0 Then
            Dim objProceedings As New clsProceedings
            With objProceedings
                .ProceedingsID = gbSearchID
                .getProceedingsByID
                'txtProceddingsNo.Tag = .ProceedingsID
                txtProceedingsNo.Text = .ProceedingsNo
                txtProceedingsDate.Text = CheckDateInMMM(.ProceedingsDate)
                
            End With
        End If
        gbSearchID = -1
        gbSearchStr = ""
        If objdb.SetConnection(mCnn) Then
            If txtProceedingsNo.Text <> "" Then
                mSql = " SELECT * From faProceedings WHERE (tnyUsed = 1)and vchProceedingsNo = '" & txtProceedingsNo.Text & "'  "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                        MsgBox "Proceedings Number Already Requested", vbInformation, "Saankhya"
                        txtProceedingsNo.Text = ""
                        txtProceedingsDate.Text = ""
                        Exit Sub
                End If
                Rec.Close
            End If
        End If
    End Sub

    Private Sub cmdRequest_Click()
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim mSql                As String
        Dim objdb               As New clsDB
        Dim mArrIn              As Variant
        Dim mArrOut             As Variant
        Dim mintID              As Variant
        Dim mProceedingsNo      As Variant
        Dim mProceedingsDate    As Date
        Dim mVoucherNo          As Variant
        Dim mReason             As Variant
        Dim mRequestedby        As Variant
        Dim mRequestedDate      As Date
        Dim mAOAuthorise        As Variant
        Dim mAOAuthoriseDate    As Date
        Dim mSecAuthorise       As Variant
        Dim mSecAuthoriseDate   As Variant
        
                
        If SaveValidation = False Then Exit Sub
                
        If objdb.SetConnection(mCnn) Then
            mProceedingsNo = txtProceedingsNo.Text
            
            mintID = IIf(txtProceedingsNo.Tag = "", -1, val(txtProceedingsNo.Tag))
            mProceedingsDate = Trim(txtProceedingsDate.Text)
            'mVoucherNo = Trim(txtReceiptsNo.Text)
            mReason = cmbReason.Text
            mRequestedby = gbUserID
            mRequestedDate = Trim(txtRequestedByDate.Text)
            mArrIn = Array(mintID, mProceedingsNo, _
                            mProceedingsDate, _
                            mVoucherID, _
                            mReason, _
                            mRequestedby, _
                            mRequestedDate, Null, _
                            Null, _
                            Null, _
                            Null, _
                            gbLocalBodyID, _
                            0, _
                            0, _
                            Trim(txtStationaryNo.Text), _
                            Trim(txtRemarks.Text) _
                        )
            objdb.ExecuteSP "spSaveReceiptCancellationRequest", mArrIn, , , mCnn, adCmdStoredProc
            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
            cmdRequest.Enabled = False
            mSql = "UPDATE   faProceedings Set tnyUsed = 1 WHERE vchProceedingsNo = '" & txtProceedingsNo.Text & " ' "
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
        End If
        Call FormInitialize
        Me.Hide
        frmListOfReceiptCancellationRequest.Show
    End Sub
    Private Sub cmdSearchReceipt_Click()
        Dim mCnn  As New ADODB.Connection
        Dim Rec   As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql  As String
        frmSearchVouchers.CheckMode = 10
        frmSearchVouchers.chkPayment.Enabled = False
        frmSearchVouchers.chkContra.Enabled = False
        frmSearchVouchers.chkJournal.Enabled = False
        frmSearchVouchers.chkInterrupted.Visible = False 'True
        'frmSearchVouchers.chkInterrupted.value = 1
        If objdb.SetConnection(mCnn) Then
            mSql = " SELECT MAX(dtDate)newdate From faVouchers WHERE tnyVoucherTypeID = 10 "
            Rec.Open mSql, mCnn
            If Rec!newdate = gbTransactionDate Then
                mSql = mSql + " And (dtDate NOT IN (SELECT MAX(dtDate) FROM  faVouchers WHERE tnyVoucherTypeID = 10 ))"
            End If
            Rec.Close
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtReceiptsDate.Tag = DdMmmYy((Rec!newdate))
                frmSearchVouchers.txtFromDate = DdMmmYy((Rec!newdate))
                frmSearchVouchers.txtToDate = DdMmmYy((Rec!newdate))
                frmSearchVouchers.txtFromDate.Enabled = False
                frmSearchVouchers.txtToDate.Enabled = False
            End If
            Rec.Close
        End If
        frmSearchVouchers.PreviousYearMode = 0
        frmSearchVouchers.Show vbModal
        If gbSearchCode <> "" Then
            txtReceiptsNo.Text = gbSearchCode
            txtReceiptsDate.Text = CheckDateInMMM(CStr(gbReceiptDate))
            If IsDate(txtReceiptsDate) Then
                If txtReceiptsDate.Text <> txtReceiptsDate.Tag Then
                    txtReceiptsNo.Text = ""
                    txtReceiptsDate.Text = ""
                    MsgBox "You can only cancel Receipt from previous Transaction Date", vbInformation
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
        If gbSearchCode <> "" Then
            If objdb.SetConnection(mCnn) Then
                mSql = " SELECT  intVoucherID From faVouchers Where intVoucherNo = " & gbSearchCode & " "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If Rec!intVoucherID <> "" Then
                        mVoucherID = (Rec!intVoucherID)
                    End If
                End If
                Rec.Close
             Else
                MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
            End If
            
           
             ''''---------Added On 26 Mar 2015 By Anisha C to block Ajdusted Receipt From Cancellation------------------------------------------
            If objdb.SetConnection(mCnn) Then
                mSql = "Select * from faVouchers Where tnyVoucherGroupID=2 and numLinkKeyID = " & gbSearchCode
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    MsgBox "This Receipt is done Adjustment Journal. So you are not allowed to Cancel again.", vbInformation
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
            End If

            
            
            gbSearchCode = ""
            gbSearchID = -1
        
            If objdb.SetConnection(mCnn) Then
               'mSql = "select faCancelledVouchers.intVoucherID from faCancelledVouchers where intVoucherID=" & mVoucherID & ""
                mSql = " select faReceiptCancellationRequest.intVoucherID from faReceiptCancellationRequest where intVoucherID=" & mVoucherID & ""
                Rec.Open mSql, mCnn
                mCnn.Execute mSql
                If Not (Rec.EOF And Rec.BOF) Then
                    If Rec!intVoucherID <> "" Then
                      MsgBox "Receipt already selected for cancellation", vbInformation, "Saankhya"
                      txtReceiptsNo.Text = ""
                      txtReceiptsDate.Text = ""
    
                      Exit Sub
                    End If
                End If
                Rec.Close
            End If
      End If
    End Sub
    Private Sub cmdSecondtAuthorize_Click()
        Dim mCnn As New ADODB.Connection
        Dim Rec   As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        If objdb.SetConnection(mCnn) Then
            mSql = "SELECT  tnyStatusAO From faReceiptCancellationRequest Where intVoucherID= " & mVoucherID & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!tnyStatusAO = 1 Then
                    crtStatus = 1
                Else
                    crtStatus = 0
                End If
            End If
            Rec.Close
        End If
        If crtStatus = 0 Then
            MsgBox "Accounts Officer not Authorized yet.....", vbInformation, "Saankhya"
            Exit Sub
      
        Else
            If objdb.SetConnection(mCnn) Then
                mSql = "Update faReceiptCancellationRequest Set tnyStatusSec = 1, numAuthorisedBySec = " & gbUserID & " ,dtAuthorisationDateSec= '" & CheckDateInMMM(txtSecondtAuthorizedDate.Text) & "' Where intID = " & val(txtProceedingsNo.Tag)
                Rec.Open mSql, mCnn
                cmdSecondtAuthorize.Enabled = False
            Else
                MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
            End If
           ' Rec.Close
        End If
        
        If objdb.SetConnection(mCnn) Then
            mSql = "Update faVouchers Set tnysync=Null,tnyStatus = 4, tnyCancelFlag = 1,tnyChangeFlag=Null Where intVoucherID = " & mVoucherID & " "
            mCnn.Execute mSql
            mSql = "Update faTransactions Set tnysync=Null,tnyStatus = 4 Where intVoucherID = " & mVoucherID & ""
            mCnn.Execute mSql
            mSql = "Update faCancelledVouchers Set tnyApproveStatus = 1, tinType=9 Where intVoucherID = " & mVoucherID & " "
            mCnn.Execute mSql
           MsgBox "Authorized Successfully", vbInformation, "Saankhya"
        End If
        Call FormInitialize
        Me.Hide
        frmListOfReceiptCancellationRequest.Show
    End Sub
    Private Sub cmdSecondtAuthorize_GotFocus()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        If objdb.SetConnection(mCnn) Then
            mSql = " SELECT  faReceiptCancellationRequest.intID FROM faReceiptCancellationRequest INNER JOIN"
            mSql = mSql + "  faVouchers ON faReceiptCancellationRequest.intVoucherID = faVouchers.intVoucherID"
            mSql = mSql + " WHERE faVouchers.intVoucherNo =  " & txtReceiptsNo.Text & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!intID <> "" Then
                     txtProceedingsNo.Tag = (Rec!intID)
                End If
            End If
            Rec.Close
        Else
            MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
        End If
    End Sub
    Private Sub dtpProceedingsDate_CloseUp()
        'txtProceedingsDate.Text = DdMmmYy(dtpProceedingsDate.value)
        'If txtProceedingsDate > gbTransactionDate Then
        'MsgBox "Proceedings Date cannot be greater than current date", vbInformation
        ''txtProceedingsDate.SetFocus
        'txtProceedingsDate.Text = DdMmmYy(gbTransactionDate)
        'End If
    End Sub
    Private Sub Form_Activate()
       ' Call FormInitialize
       Me.Top = 500
        Me.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FormInitialize
        Call CheckLastPostingDate
    End Sub
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        'dtpProceedingsDate.value = gbTransactionDate
        
        Call PopulateList(cmbReason, "SELECT vchCancelReason,intCancelID From faCancelReason Where (intCancelID <> 5) ORDER BY vchCancelReason", , True, True, True, enuSourceString.Saankhya)
       ' mSearch = 0
        crtStatus = 0
        txtReceiptsNo.Enabled = False
        txtRequestedBy.Enabled = False
        txtRequestedByDate.Enabled = False
        txtReceiptsDate.Enabled = False
        txtStationaryNo.Enabled = False
        If gbSeatGroupID = gbSeatGroupAccountsSuperintended Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then                  'Accounts Officer
            cmdFirstAuthorize.Enabled = True
            cmdSecondtAuthorize.Enabled = False
            txtFirstAuthorizedBy.Text = gbSeatName
            txtFirstAuthorizedDate.Text = DdMmmYy(gbDate)
            cmdRequest.Enabled = False
            txtProceedingsNo.Enabled = False
            txtProceedingsDate.Enabled = False
            cmdProceedings.Enabled = False
            cmdSearchReceipt.Enabled = False
            'dtpProceedingsDate.Enabled = False
            cmbReason.Enabled = False
            txtStationaryCount.Text = 1
            txtStationaryCount.Enabled = False
            txtStationaryNo.Enabled = False
            txtRequestedBy.Enabled = False
            txtRequestedByDate.Enabled = False
            txtRemarks.Enabled = False
        ElseIf gbSeatGroupID = gbSeatGroupSecretary Then              'Secretary
            cmdSecondtAuthorize.Enabled = True
            cmdFirstAuthorize.Enabled = False
            txtSecondtAuthorizedBy.Text = gbSeatName
            txtSecondtAuthorizedDate.Text = DdMmmYy(gbDate)
            cmdRequest.Enabled = False
            txtProceedingsNo.Enabled = False
            txtProceedingsDate.Enabled = False
            cmdProceedings.Enabled = False
            cmbReason.Enabled = False
            txtRequestedBy.Enabled = False
            txtRequestedByDate.Enabled = False
            cmdSearchReceipt.Enabled = False
            cmbReason.Enabled = False
            'dtpProceedingsDate.Enabled = False
            txtRemarks.Enabled = False
            txtStationaryNo.Enabled = False
        ElseIf gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier Then  'Cashier or ChiefCashier
            txtRequestedBy.Text = gbSeatName
            txtRequestedByDate.Text = DdMmmYy(gbDate)
            cmdRequest.Enabled = True
            txtProceedingsNo.Enabled = False
            txtProceedingsDate.Enabled = False
           
            txtFirstAuthorizedBy.Enabled = False
            txtFirstAuthorizedDate.Enabled = False
            txtSecondtAuthorizedBy.Enabled = False
            txtSecondtAuthorizedDate.Enabled = False
            cmdFirstAuthorize.Enabled = False
            cmdSecondtAuthorize.Enabled = False
            txtStationaryCount.Enabled = False
            txtStationaryNo.Enabled = False
        End If
    End Sub

    Private Sub txtFirstAuthorizedDate_LostFocus()
        If Trim(txtFirstAuthorizedDate.Text) <> "" Then
           txtFirstAuthorizedDate.Text = CheckDateInMMM(txtFirstAuthorizedDate.Text)
        End If
    End Sub
    Private Sub txtProceedingsDate_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
          KeyAscii = 0
        End If
        
    End Sub
    Private Sub txtProceedingsDate_LostFocus()
         txtProceedingsDate.Text = CheckDateInMMM(txtProceedingsDate.Text)
         If txtProceedingsDate > gbTransactionDate Then
            MsgBox "Proceedings Date cannot be greater than current date", vbInformation
            txtProceedingsDate.SetFocus
        End If
    End Sub
    Private Function SaveValidation() As Boolean
        If Trim(txtProceedingsNo.Text) = "" Then
            MsgBox "Please Enter the Proceedings number", vbInformation, "Saankhya"
            'txtProceedingsNo.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtProceedingsDate.Text) = "" Then
            MsgBox "Please Enter the Proceedings Date", vbInformation, "Saankhya"
            txtProceedingsDate.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtReceiptsNo.Text) = "" Then
            MsgBox "Please select the Receipt number ", vbInformation, "Saankhya"
            SaveValidation = False
            Exit Function
        End If
        If cmbReason.ListIndex <= 0 Then
             MsgBox "Please Select the Reason for Cancellation", vbInformation
             cmbReason.SetFocus
             SaveValidation = False
             Exit Function
        End If
        If txtStationaryCount.Text = "" Then
             MsgBox "Please Give the Stationary Count", vbInformation
             txtStationaryCount.SetFocus
             SaveValidation = False
             Exit Function
        End If
        If txtStationaryNo.Text = "" Then
             MsgBox "Please Give the Stationary Number", vbInformation
             txtStationaryNo.SetFocus
             SaveValidation = False
             Exit Function
        End If
        If Trim(txtRequestedBy.Text) = "" Then
            MsgBox "Please Enter the name ", vbInformation, "Saankhya"
            txtRequestedBy.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtRequestedByDate.Text) = "" Then
            MsgBox "Please Enter the date ", vbInformation, "Saankhya"
            txtRequestedByDate.SetFocus
            SaveValidation = False
            Exit Function
        End If
        SaveValidation = True
    End Function
    Private Sub txtProceedingsNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtReceiptsDate_LostFocus()
        txtReceiptsDate.Text = CheckDateInMMM(txtReceiptsDate.Text)
       
    End Sub

    Private Sub txtReceiptsNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    
    Private Sub txtStationaryNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
          KeyAscii = 0
        End If
    End Sub

    Private Sub txtRequestedByDate_LostFocus()
        txtRequestedByDate.Text = CheckDateInMMM(txtRequestedByDate.Text)
    End Sub
    Private Sub txtSecondtAuthorizedDate_LostFocus()
        If Trim(txtSecondtAuthorizedDate.Text) <> "" Then
            txtSecondtAuthorizedDate.Text = CheckDateInMMM(txtSecondtAuthorizedDate.Text)
        End If
    End Sub

    Private Function CancelReceipt() As Boolean
    
            Dim aryIn As Variant
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim mStatCount As Integer
            Dim mStatNo As Long
            Dim Rec As New ADODB.Recordset
            Dim mFormat As String
            Dim mCancelID As Integer
            
            If objdb.SetConnection(mCnn) Then
           
                Rec.Open "Select intCancelID From faCancelReason where vchCancelReason = '" & Trim(cmbReason.Text) & "' ", mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    mCancelID = Rec!intCancelID
                End If
                Rec.Close
              
           
            End If
            
            mStatCount = val(txtStationaryCount.Text)
            mStatNo = val(txtStationaryNo.Text)
            'If objDb.SetConnection(mCnn) Then
            '   Rec.Open "Select Count(*) as Cnt from faCancelledVouchers", mCnn
             '   If Not (Rec.EOF Or Rec.BOF) Then
             '       mFormat = Rec!Cnt
              '  End If
              '  If Rec.State = 1 Then Rec.Close
                While (mStatCount)
                    aryIn = Array(mVoucherID, _
                                10, _
                                Null, _
                                Trim(txtReceiptsNo.Text), _
                                gbUserID, _
                                gbCounterID, _
                                gbSeatID, _
                                mCancelID, _
                                gbTransactionDate, _
                                mStatNo, _
                                Null, _
                                9 _
                                ) 'mFormat
                    objdb.ExecuteSP "spSaveCancelledVouchers", aryIn, , , mCnn, adCmdStoredProc
                    mStatCount = mStatCount - 1
                    mStatNo = mStatNo + 1
                Wend
           ' End If
            CancelReceipt = True
        Exit Function
    End Function

    Private Function CancelPropertyTax(txtRecieptNo As Double, mVoucherID As Long)
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            'Dim mCnnSanchaya As New ADODB.Connection
            Dim Rec As New Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            Dim arrIn As Variant
            Dim mQry As String
            
            
            Dim blnConfig As Boolean
            Dim blnOtherZoneOfficeFlag As Boolean
            
            If objdb.SetConnection(mCnn) Then
                mQry = "Select tnyLinkWithPropertyTax from faConfig"
                Rec.Open mQry, mCnn
                If IsNull(Rec!tnyLinkWithPropertyTax) Then
                    blnConfig = False
                ElseIf val(Rec!tnyLinkWithPropertyTax) = 1 Then
                    blnConfig = True
                Else
                    blnConfig = False
                End If
                If Rec.State = 1 Then Rec.Close
                
                mSql = "Select numZoneID as ZoneID from faVouchers Where intVoucherNo = " & Trim(txtRecieptNo)
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!ZoneID <> gbLocationID Then
                        blnOtherZoneOfficeFlag = True
                    Else
                        blnOtherZoneOfficeFlag = False
                    End If
                End If
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
            
            If blnConfig = True Then
                Set mCnn = Nothing
                If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
                    If blnOtherZoneOfficeFlag = False Then
                        arrIn = Array(Trim(txtRecieptNo))
                        objdb.ExecuteSP "spReverseDemandFromSaankhya", arrIn, , , mCnn
                    Else
                        '---------------------------------------------------------------'
                        ' Other Zone Office Collection Modified on 13-aug-2009 By cijith'
                        '---------------------------------------------------------------'
                        arrIn = Array(gbLocationID, mVoucherID)
                        objdb.ExecuteSP "HOSaanOtherCollectionCancel", arrIn, , , mCnn
                        '----------------------------------------------------------'
                    End If
                Else
                    MsgBox "Connection To Sanchaya Does not Exist, Please Contact your System Administrtor", vbInformation
                End If
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Sub CheckLastPostingDate()   '-----------------LAST POSTING VALIDATION------------------
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim dtCurrentDate As Date
        
        Call SetgbLastPostingDate
        objdb.SetConnection mCnn
        mSql = "Select GETDATE()CurrentDate From faFinancialYear "
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            dtCurrentDate = Format(Rec!currentdate, "dd-mmm-yyyy")
            If CDate(dtCurrentDate) <= CDate(gbLastPostingDate) Then
                MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
                cmdRequest.Enabled = False
                cmdFirstAuthorize.Enabled = False
                cmdSecondtAuthorize.Enabled = False
                Exit Sub
            End If
            
        End If
        
    End Sub
    Private Function PTaxWebDemand(mVoucherID As Long)
        
        Dim mCollPost       As String
        Dim mColZoneID      As String
        Dim mBuildingIdWeb  As String
        Dim mColAmt            As String
        Dim mColDate        As String
        Dim mColReceiptNo   As String
        Dim mColBookNo      As String
        Dim mColPeriodId     As String
        Dim mColYearID       As String
        Dim mHash           As String
        Dim mCollOut        As String
'                    Dim node            As IXMLDOMNode
'                    Dim DataNodes       As IXMLDOMNodeList
        Dim mUrl            As String
        Dim objSOAP         As Variant
        Dim mLen            As Integer
        Dim mColAccID       As String
        Dim mColKeyID       As String
        Dim Rec             As New ADODB.Recordset
        
    mUrl = gbDefaultUrlSanchayaPost
    Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
    objSOAP.MSSoapInit mUrl + "?WSDL"
        Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                
                mColAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                mColDate = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                mColReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                mColBookNo = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
                mColPeriodId = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                mColYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                mBuildingIdWeb = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                mColZoneID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                mColAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                mColKeyID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                If mColAccID <> gbAcHeadIDPenalInterest Then
                    mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(mColZoneID) + "#" + CStr(mBuildingIdWeb) + "#"
                    mCollPost = mCollPost + CStr(mColYearID) + "#" + CStr(mColPeriodId) + "#" + CStr(mVoucherID) + "#"
                    mCollPost = mCollPost + CStr(mColBookNo) + "#" + CStr(mColReceiptNo) + "#" + CStr(mColDate) + "#"
                    mCollPost = mCollPost + CStr(gbFinancialYearID) + "#" + CStr(mColAmt) + "#" + CStr(gbLBName) + "#"
                    mCollPost = mCollPost + CStr(mColAccID) + "#" + CStr(mColKeyID)
                End If
                Rec.MoveNext
                mCollPost = mCollPost + "~"
            Wend
            mLen = Len(mCollPost) - 1
            mCollPost = Left$(mCollPost, mLen - 1)
            mHash = CStr(mVoucherID) + CStr(mBuildingIdWeb) + "ikm#9567" + CStr(mColDate) + "*ikm#9567"
            mCollOut = objSOAP.Saankhyaa_CollectionPostingCancel(mCollPost, mHash)
        End If
    End Function
  
