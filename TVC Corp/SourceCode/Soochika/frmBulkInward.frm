VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSoochikaBulkInward 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Bulk Inward"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBulkInward.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11775
   Begin VB.Frame FrameBulkInward 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bulk Inward"
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11775
      Begin VB.ListBox lstSubject 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   2325
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   2325
         TabIndex        =   33
         Top             =   690
         Width           =   8775
      End
      Begin VB.ComboBox cboseatid 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   6480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cboSeat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   6000
         Width           =   2535
      End
      Begin VB.ComboBox cmbDept 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox cmbForwardTo 
         Height          =   375
         Left            =   3960
         TabIndex        =   27
         Top             =   6675
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtTotalItems 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6675
         TabIndex        =   5
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtMainRequestId 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   1
         Top             =   690
         Width           =   465
      End
      Begin VB.ComboBox cmbRequestType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmBulkInward.frx":0442
         Left            =   11115
         List            =   "frmBulkInward.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1830
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   5295
         Width           =   7050
      End
      Begin VB.ComboBox cmbHospital 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4815
         Width           =   7050
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   360
         Left            =   10050
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4830
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   360
         Left            =   10050
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5490
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   360
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3450
         Width           =   1335
      End
      Begin VB.CommandButton cmdApproval 
         Caption         =   "&Approve"
         Enabled         =   0   'False
         Height          =   360
         Left            =   10050
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4140
         Width           =   1335
      End
      Begin VB.ComboBox cmbForwardtoold 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmBulkInward.frx":0446
         Left            =   1830
         List            =   "frmBulkInward.frx":0459
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   6675
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.TextBox txtStartInwardNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1830
         TabIndex        =   15
         Top             =   1830
         Width           =   1515
      End
      Begin VB.ComboBox cmbRequestSubType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         ItemData        =   "frmBulkInward.frx":0495
         Left            =   2325
         List            =   "frmBulkInward.frx":0497
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1245
         Width           =   9285
      End
      Begin VB.TextBox txtSubRequestId 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   3
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   360
         Left            =   10065
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2805
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpInwDate 
         Height          =   330
         Left            =   1830
         TabIndex        =   16
         Top             =   240
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65208321
         CurrentDate     =   38344
      End
      Begin VSFlex8LCtl.VSFlexGrid vsfInward 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   9615
         _cx             =   16960
         _cy             =   4048
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   14142647
         ForeColorSel    =   -2147483634
         BackColorBkg    =   11587566
         BackColorAlternate=   8438015
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBulkInward.frx":0499
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
         Editable        =   1
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
         BackColorFrozen =   14868689
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComCtl2.DTPicker dtpApplDate 
         Height          =   330
         Left            =   6660
         TabIndex        =   24
         Top             =   240
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65208321
         CurrentDate     =   38344
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   30
         Top             =   6030
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   6030
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No of Items"
         Height          =   195
         Left            =   4950
         TabIndex        =   26
         Top             =   1890
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Date"
         Height          =   195
         Left            =   4950
         TabIndex        =   25
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Type "
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lbltype 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Request"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   765
         Width           =   1395
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   5370
         Width           =   765
      End
      Begin VB.Label lblHospital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   4875
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded to"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   6750
         Width           =   1125
      End
      Begin VB.Label lblinwarddate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inward Date"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Inward No"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1890
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmSoochikaBulkInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flgInward As Integer
Dim SevanaMainSubid As Integer

Private Sub cboSeat_Change()
    cboseatid.ListIndex = cboSeat.ListIndex
End Sub
Private Sub cboSeat_Click()
    cboseatid.ListIndex = cboSeat.ListIndex
End Sub

Private Sub cmbDept_Click()
    If (cmbDept.ListIndex > -1) Then
        If gbLinkWithSoochika <> 1 Then 'For unicode version added by vipin on 24/09/2012
            Call PopulateList(cboseatid, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDept.ItemData(cmbDept.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
            Call PopulateList(cboSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDept.ItemData(cmbDept.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
        Else
            Call PopulateList(cboseatid, "SELECT  intid,chvsection From tblSection  inner join TblUser on TblUser.intSection=tblSection.intid WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & cmbDept.ItemData(cmbDept.ListIndex) & " order by intID", , True, , True, enuSourceString.Soochika)
            Call PopulateList(cboSeat, "SELECT  chvsection,chvsection From tblSection  inner join TblUser on TblUser.intSection=tblSection.intid WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & cmbDept.ItemData(cmbDept.ListIndex) & " order by intID", , True, , True, enuSourceString.Soochika)
        End If
    End If
End Sub
Private Sub cmbRequestType_Click()
    Dim con As New ADODB.Connection
    Dim objdb As New clsDB
    If cmbRequestType.ListIndex <> -1 Then
        If gbLinkWithSoochika <> 1 Then
            PopulateList cmbRequestSubType, "SELECT TypeofSubRequest, intID From mSubjectSevanaSubtype where intSubTypeID='" & cmbRequestType.ItemData(cmbRequestType.ListIndex) & "'", , , , True, enuSourceString.SoochikaUnicode
        Else
            PopulateList cmbRequestSubType, "SELECT TypeofSubRequest, intID From TblSubjectSubType where intSubTypeID='" & cmbRequestType.ItemData(cmbRequestType.ListIndex) & "'", , , , True, enuSourceString.Soochika
        End If
    End If
    If cmbRequestType.ListIndex = 4 Then
        lblHospital.Visible = False
        cmbHospital.Visible = False
    Else
        lblHospital.Visible = True
        cmbHospital.Visible = True
    End If
    If cmbRequestType.ListIndex <> -1 Then
       txtMainRequestId.Text = cmbRequestType.ItemData(cmbRequestType.ListIndex)
    End If
End Sub
Private Sub cmbRequestSubType_Click()
    Dim con As New ADODB.Connection
    Dim objdb As New clsDB
    If cmbRequestSubType.ListIndex <> -1 Then
       txtSubRequestId.Text = cmbRequestSubType.ItemData(cmbRequestSubType.ListIndex)
'       Select Case cmbRequestSubType.ItemData(cmbRequestSubType.ListIndex)
'       Case 2, 3
'            Label4.Caption = "Arrival Date"
'            cmbHospital.ListIndex = -1
'            cmbHospital.Enabled = False
'       Case 7, 9, 11, 13, 15, 17, 19, 26, 28, 30, 32, 34, 41, 43, 45, 51, 53, 55, 58, 60, 61
'            PopulateList cmbForwardTo, "SELECT TypeofSubRequest, intid  FROM  TblSubjectSubType WHERE intid=3", , , , True, enuSourceString.SOOCHIKA
'            cmbForwardTo.ListIndex = 0
'            cmbForwardTo.Enabled = False
'            Label4.Caption = "Application Date"
'            cmbHospital.Enabled = True
'       Case Else
'            PopulateList cmbForwardTo, "SELECT TypeofSubRequest, intid  FROM  TblSubjectSubType WHERE intid IN (0,1,2)", , , , True, enuSourceString.SOOCHIKA
'            cmbForwardTo.ListIndex = 0
'            cmbForwardTo.Enabled = True
'            Label4.Caption = "Application Date"
'            cmbHospital.Enabled = True
'       End Select
       GetKiosk (cmbRequestSubType.ItemData(cmbRequestSubType.ListIndex))
    End If
End Sub
Public Sub GetKiosk(ByVal SubTypeID As Variant)
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    
    If gbLinkWithSoochika <> 1 Then
    objdb.CreateNewConnection mCnn, enuSourceString.SoochikaUnicode
    mSql = "select tnyToSeat from mSubjectSevanaSubtype where intID=" & SubTypeID
    Else
    objdb.CreateNewConnection mCnn, enuSourceString.Soochika
    mSql = "select tnyToSeat from tblSubjectSubType where intid=" & SubTypeID
    End If
    
    If SubTypeID = 93 Then
        MsgBox "This subtype is blocked", vbInformation
        txtSubRequestId.Text = ""
        cmbRequestSubType.ListIndex = 0
        Exit Sub
    Else
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
    '        For i = 0 To cmbForwardTo.ListCount - 1
    '            If cmbForwardTo.ItemData(i) = Rec!tnyToSeat Then
    '                cmbForwardTo.ListIndex = i
    '            End If
    '        Next
            cmbForwardTo.Text = Rec!tnyToSeat
        End If
    End If
End Sub
Private Sub cmdApproval_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.value = Now
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.ListIndex >= 0 Then ctl.ListIndex = -1
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Rows = 1
        End If
    Next ctl
    EnCtls
    cmdNew.Enabled = False
    cmdPrint.Enabled = True
    
    'Added By Akheel on 18/08/2009 Purpose:Getting the next Inward no for the application
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    If gbLinkWithSoochika <> 1 Then
    
    If objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False Then
        MsgBox "Soochika Connection Failed", vbDefaultButton1
        Exit Sub
        Else
       Rec.Open "select isnull(max(numCurrentNo),1) as MaxFileID from tInwardDetails where year(dtDateofReceipt)=year(getdate())", mCnn
        
    End If
    Else
    If objdb.CreateNewConnection(mCnn, enuSourceString.Soochika) = False Then
        MsgBox "Soochika Connection Failed", vbDefaultButton1
        Exit Sub
        Else
      Rec.Open "select isnull(max(fldCurrentNo),1) as MaxFileID from TblTappalDetails where year(Flddateofreceipt)=year(getdate())", mCnn
    End If
    End If
    
    If Not (Rec.EOF Or Rec.BOF) Then
        txtStartInwardNo.Text = CDbl(Rec!MaxFileID) 'CDbl(Right(Rec!MaxFileID, 6) + 1)
    End If
    Rec.Close
    mCnn.Close
    flgInward = 0
    SevanaMainSubid = 0
End Sub

Private Sub cmdPrint_Click()
    Dim i As Integer
    If ValidateValues Then
'        Open "prn" For Output As #1
'        Print #1, " "
'        Print #1, "----------------------------------------------------------------------------"
'
'        Print #1, Space(1) & "Total Inwards: " & Val(txtTotalItems) & Space(5) & "Inward Date. :" & Trim(dtpInwDate.Value)
'        Print #1, Space(1) & "Request Type: " & Trim(cmbRequestType.Text)
'        Print #1, Space(1) & "Request Sub Type: " & Trim(cmbRequestSubType.Text)
'        Print #1, Space(1) & "Hospital: " & cmbHospital.Text
'        Print #1, Space(1) & "Remarks: " & Trim(txtRemarks);
'        Print #1, Space(1) & "Forwarded To: " & cmbForwardto.Text
'        Print #1, "-----------------------------------------------------------------------------"
'        Print #1, vbCrLf
'        Close #1
        cmdApproval.Enabled = True
    End If
End Sub

Private Sub cmdSave_Click()
    Dim inwdate As String
    Dim AppnDate As String
    Dim vInward(25)
    Dim SooInward As Variant
    Dim dbCnn As New ADODB.Connection
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim vOut, vReceipt As Variant
    Dim mSql As String
    Dim SubID As Variant
    Dim Subject As Variant
    Dim Rec As New ADODB.Recordset
    ReDim vReceipt(4)
    ReDim SooInwardU(10)
    ReDim SooInwardUTr(8)
    ReDim SooInwardUFromSoochika(25)
    ReDim ArrSooSave(12)
    ReDim SooInward(11)
    ReDim vOut(1)
    Dim vOut1 As Variant
    Dim MainSubTypeID As Variant
    Dim FileID As Variant
    
    AppnDate = Format(dtpApplDate.value, "DD-MMM-YYYY")
    inwdate = Format(dtpInwDate.value, "DD-MMM-YYYY")
    If ValidateValues Then
    
    If gbLinkWithSoochika <> 1 Then
            objdb.CreateNewConnection dbCnn, enuSourceString.SevanaRegn
            objdb.CreateNewConnection mCnn, enuSourceString.SoochikaUnicode
                    SubID = txtMainRequestId.Text
                    Subject = txtSubject.Text
         For intRow = 1 To vsfInward.Rows - 1
            SooInwardU(0) = Trim(vsfInward.TextMatrix(intRow, 2))    'Sendername
            SooInwardU(1) = Trim(vsfInward.TextMatrix(intRow, 3))    'Locality
            SooInwardU(2) = Subject                                  'Subject
            SooInwardU(3) = SubID                                    'Subject ID
            Rec.Open "select numSubjectSuiteID from mSubjectSuite where numSubjectID=" & txtMainRequestId.Text & " and intSuiteID=112", mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                If Rec!numSubjectSuiteID <> "" Then
                MainSubTypeID = Rec!numSubjectSuiteID
                SooInwardU(4) = txtSubRequestId.Text
                End If
            End If
            Rec.Close
            SooInwardU(5) = cboseatid.Text 'Forward Seat
            SooInwardU(6) = cboseatid.Text 'Current Seat
            SooInwardU(7) = gbnumZonalID
            SooInwardU(8) = gbDistID
            SooInwardU(9) = 112
            SooInwardU(10) = gbLBID
            objdb.ExecuteSP "SpSaveBulkInward", SooInwardU, vOut, , mCnn, CommandTypeEnum.adCmdStoredProc
            
            'MsgBox vOut(0, 0)
            
            SooInwardUTr(0) = vOut(0, 0)
'''            SooInwardUTr(1) = gbSeatID  'Forward Seat
'''            SooInwardUTr(2) = cboseatid.Text 'Current Seat
'''
             'interchnage the seatd 07.10.16 on added on 14 oct 2016
            SooInwardUTr(1) = cboseatid.Text  'Forward Seat
            SooInwardUTr(2) = gbSeatID 'Current Seat
            
            Rec.Open "select isnull(numcurrentUserID,0) as CurrentUserID from tSeatDetails where numSeatID=" & cboseatid.Text, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
            SooInwardUTr(3) = Rec!CurrentUserID
            End If
            SooInwardUTr(4) = gbUserID
            SooInwardUTr(5) = "Processing"
            SooInwardUTr(6) = ""
            SooInwardUTr(7) = ""
            SooInwardUTr(8) = 0
            objdb.ExecuteSP "SpSaveInwardTrackDetails", SooInwardUTr, , , mCnn, CommandTypeEnum.adCmdStoredProc
            Rec.Close
          
                If MainSubTypeID <> 0 Then
                        
                            SooInwardUFromSoochika(0) = CDbl(Right(vOut(0, 0), 6))
                            SooInwardUFromSoochika(1) = Null
                            SooInwardUFromSoochika(2) = MainSubTypeID
                            If cmbHospital.ItemData(cmbHospital.ListIndex) > 0 Then
                            SooInwardUFromSoochika(3) = cmbHospital.ItemData(cmbHospital.ListIndex)
                            Else
                            SooInwardUFromSoochika(3) = Null
                            End If
                            
                            Rec.Open "Select tnyType,tnyToSeat from mSubjectSevanaSubType where intid=" & txtSubRequestId.Text, mCnn
                            If Not (Rec.EOF Or Rec.BOF) Then
                            SooInwardUFromSoochika(4) = Rec!tnyToSeat
                            End If
                            Rec.Close
                            SooInwardUFromSoochika(5) = dtpApplDate.value
                            SooInwardUFromSoochika(6) = Null
                            SooInwardUFromSoochika(7) = Trim(vsfInward.TextMatrix(intRow, 3))
                            SooInwardUFromSoochika(8) = Null
                            SooInwardUFromSoochika(9) = Null
                            SooInwardUFromSoochika(10) = Null
                            SooInwardUFromSoochika(11) = Null
                            SooInwardUFromSoochika(12) = Null
                            SooInwardUFromSoochika(13) = Null
                            SooInwardUFromSoochika(14) = Trim(vsfInward.TextMatrix(intRow, 2))
                            SooInwardUFromSoochika(15) = Null
                            SooInwardUFromSoochika(16) = gbDistID
                            SooInwardUFromSoochika(17) = 32
                            SooInwardUFromSoochika(18) = Null
                            SooInwardUFromSoochika(19) = txtSubRequestId.Text
                            SooInwardUFromSoochika(20) = Null
                            SooInwardUFromSoochika(21) = Null
                            SooInwardUFromSoochika(22) = "Date Entered By" & gbUserName & " -" & txtRemarks.Text
                            SooInwardUFromSoochika(23) = Null
                            SooInwardUFromSoochika(24) = Null
                            SooInwardUFromSoochika(25) = Null
                       '     SooInwardUFromSoochika(26) = Null
                            'SooInwardUFromSoochika(24) = cmbHospital.Text
                            objdb.ExecuteSP "spSaveInwardFromSoochika", SooInwardUFromSoochika, vOut1, , dbCnn, CommandTypeEnum.adCmdStoredProc
                            
                            
                            ArrSooSave(0) = SooInwardUTr(0)
                            ArrSooSave(1) = vOut1(0, 0)
                            ArrSooSave(2) = SooInwardUFromSoochika(2)
                            ArrSooSave(3) = SooInwardUFromSoochika(19)
                            ArrSooSave(4) = cmbHospital.ItemData(cmbHospital.ListIndex)
                            ArrSooSave(5) = cmbHospital.Text
                            
                            ArrSooSave(6) = ""
                            ArrSooSave(7) = ""
                            ArrSooSave(8) = ""
                            ArrSooSave(9) = ""
                            ArrSooSave(10) = ""
                            ArrSooSave(11) = ""
                            ArrSooSave(12) = ""
                            objdb.ExecuteSP "SpSaveInwardSevanaDetails", ArrSooSave, , , mCnn, CommandTypeEnum.adCmdStoredProc
                   End If
            
                    If SevanaMainSubid <> 0 Then
                         vInward(0) = CDbl(Right(vOut(0, 0), 6))    '   @InWNo_1    [varchar](10),
                         vInward(1) = inwdate       '   @InwDate_2  [varchar](15),
                         vInward(2) = SevanaMainSubid ' cmbRequestType.ItemData(cmbRequestType.ListIndex)       '   @InwRequest_3   [int],
                         If cmbHospital.ListIndex >= 0 Then
                             vInward(3) = cmbHospital.ItemData(cmbHospital.ListIndex)  '@InwHospital_5  [int],
                         Else
                             vInward(3) = Null
                         End If
                         vInward(4) = cmbForwardTo.Text ' cmbForwardTo.ItemData(cmbForwardTo.ListIndex)       '   @InwForward_6   [varchar](100),
                         vInward(5) = AppnDate       '   @AppnDate_7 [varchar](15),
                         vInward(14) = Trim(vsfInward.TextMatrix(intRow, 2))      '   @Name   [varchar](250),
                         vInward(19) = cmbRequestSubType.ItemData(cmbRequestSubType.ListIndex)       '   @InRequestSub   [int],
                         vInward(22) = Trim(txtRemarks)
                         vInward(23) = 0
                         vInward(25) = 0
                         
                        objdb.ExecuteSP "insert_tInward_2", vInward, vOut, , dbCnn, CommandTypeEnum.adCmdStoredProc
                    End If
        Next intRow
        MsgBox "Data Saved Successfully  ", vbInformation
        Call DisCtls
        cmbDept.ListIndex = 0
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        cmdApproval.Enabled = False
        cmdPrint.Enabled = False
        Exit Sub
    
    Else
                        objdb.CreateNewConnection dbCnn, enuSourceString.SevanaRegn
                        objdb.CreateNewConnection mCnn, enuSourceString.Soochika
                        Rec.Open "select intsubID,chvSubject from TblSubjectCoding where intMainSubID=" & txtMainRequestId.Text, mCnn
                                If Not (Rec.EOF Or Rec.BOF) Then
                                    SubID = Rec!intSubID
                                    Subject = Rec!chvSubject
                                End If
                        Rec.Close
                        mCnn.BeginTrans
                        dbCnn.BeginTrans
                        On Error GoTo rollback
                '         rec.Open "select Soochika
                         For intRow = 1 To vsfInward.Rows - 1
                            '//////////////////////////////////////////////////////////////////////
                            'Added By Akheel on 18/08/2009 For saving the data to soochika database
                            SooInward(0) = val(vsfInward.TextMatrix(intRow, 1))     'Inward No
                            SooInward(1) = Trim(vsfInward.TextMatrix(intRow, 2))    'Sendername
                            SooInward(2) = Trim(vsfInward.TextMatrix(intRow, 3))    'Locality
                            SooInward(3) = Subject                                  'Subject
                            SooInward(4) = SubID                                    'Subject ID
                            SooInward(5) = cboseatid.Text                           'Current Seat
                            SooInward(6) = gbnumSeatID                              'Forward Seat
                            SooInward(7) = gbnumZonalID
                                        If SevanaMainSubid <> 0 Then
                                            SooInward(8) = cmbRequestSubType.ItemData(cmbRequestSubType.ListIndex)  'subtypeID
                                            SooInward(11) = 102
                                        Else
                                            SooInward(8) = Null
                                            SooInward(11) = 105
                                        End If
                            SooInward(9) = gbDistID
                                        If SevanaMainSubid <> 0 Then
                                            SooInward(10) = cmbHospital.Text
                                        Else
                                            SooInward(10) = Null
                                        End If
                            
                            objdb.ExecuteSP "SpSaveBulkInward", SooInward, vOut, , mCnn, CommandTypeEnum.adCmdStoredProc
                            
                                    If SevanaMainSubid <> 0 Then
                                         vInward(0) = CDbl(Right(vOut(0, 0), 6))    '   @InWNo_1    [varchar](10),
                                         vInward(1) = inwdate       '   @InwDate_2  [varchar](15),
                                         vInward(2) = SevanaMainSubid ' cmbRequestType.ItemData(cmbRequestType.ListIndex)       '   @InwRequest_3   [int],
                                         If cmbHospital.ListIndex >= 0 Then
                                             vInward(3) = cmbHospital.ItemData(cmbHospital.ListIndex)  '@InwHospital_5  [int],
                                         Else
                                             vInward(3) = Null
                                         End If
                                         vInward(4) = cmbForwardTo.Text ' cmbForwardTo.ItemData(cmbForwardTo.ListIndex)       '   @InwForward_6   [varchar](100),
                                         vInward(5) = AppnDate       '   @AppnDate_7 [varchar](15),
                                         vInward(14) = Trim(vsfInward.TextMatrix(intRow, 2))      '   @Name   [varchar](250),
                                         vInward(19) = cmbRequestSubType.ItemData(cmbRequestSubType.ListIndex)       '   @InRequestSub   [int],
                                         vInward(22) = Trim(txtRemarks)
                                         vInward(23) = 0
                                         vInward(25) = 0
                                         
                                        objdb.ExecuteSP "insert_tInward_2", vInward, vOut, , dbCnn, CommandTypeEnum.adCmdStoredProc
                                    End If
                                Next intRow
                                MsgBox "Data Saved Successfully  ", vbInformation
                                Call DisCtls
                                cmbDept.ListIndex = 0
                                cmdNew.Enabled = True
                                cmdSave.Enabled = False
                                cmdApproval.Enabled = False
                                cmdPrint.Enabled = False
    End If
   
    Exit Sub
    End If
rollback: MsgBox "Error Saving data", vbInformation
    mCnn.RollbackTrans
    dbCnn.RollbackTrans
End Sub
Private Sub Form_Activate()
    Me.Height = 6960
    Me.Width = 11895
End Sub
Private Sub Form_Load()
    Dim objdb As New clsDB
    Dim con As New ADODB.Connection
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    gSubCenterForm Me
    If objdb.CreateNewConnection(con, enuSourceString.SevanaRegn) = False Then
        MsgBox "Sevana Connection Failed", vbDefaultButton1
        Exit Sub
    End If
    
    If gbLinkWithSoochika <> 1 Then
    If objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False Then
        MsgBox "Soochika Connection Failed", vbDefaultButton1
        Exit Sub
    End If
    End If
    If gbLinkWithSoochika = 1 Then
    If objdb.CreateNewConnection(mCnn, enuSourceString.Soochika) = False Then
        MsgBox "Soochika Urban Connection Failed", vbDefaultButton1
        Exit Sub
    End If
    End If
    
    PopulateList cmbRequestType, "SELECT TypeofRequest, intid From mInwardType", , , , True, enuSourceString.SevanaRegn
    PopulateList cmbHospital, "SELECT chvEngHospital, intID From mHospital", , , , True, enuSourceString.SevanaRegn
    If gbLinkWithSoochika <> 1 Then
     PopulateList cmbDept, "SP_SelectDepartment 1", , False, True, True, enuSourceString.SoochikaUnicode
     Else
    PopulateList cmbDept, "spselectdepartment", , False, True, True, enuSourceString.Soochika
    End If
'    FillCombo cmbForwardto, "SELECT chvForwardTo, intid  FROM  mFORWARDTO WHERE intid IN (0,1,2)", con
    dtpApplDate.value = Date
    dtpInwDate.value = Date
    dtpInwDate.MaxDate = Date
    dtpApplDate.MaxDate = Date
    con.Close
    Set con = Nothing
    cmdApproval.Enabled = False
    cmdSave.Enabled = False
    flgInward = 0
    DisCtls
    cmbForwardTo.Visible = False
    Label1.Visible = False
    Me.Width = 11880
End Sub
Private Sub txtInwardNo_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
           KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (MsgBox("Do you want to close this window??", vbYesNo) = vbYes) Then
        Cancel = 0
    Else
        Cancel = 1
    End If
End Sub

Private Sub lstSubject_DblClick()
    txtSubject.Text = lstSubject.Text
    lstSubject.Visible = False
    GetSevanaSubType
End Sub
Private Sub txtMainRequestId_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
           KeyAscii = 0
    End If
End Sub
Private Sub txtMainRequestId_LostFocus()
'    Dim intLstIdx As Integer
'    If Trim(txtMainRequestId.Text) <> "" Then
'        If Trim(txtMainRequestId.Text) <= 5 Then
'            cmbRequestType.ListIndex = txtMainRequestId.Text - 1
'        Else
'            MsgBox "Item not Found", vbInformation
'            txtMainRequestId.Text = ""
'        End If
'    End If
     If txtMainRequestId.Text <> "" Then
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        If gbLinkWithSoochika <> 1 Then
            objdb.CreateNewConnection mCnn, enuSourceString.SoochikaUnicode
            mSql = "Select chvSubject from mSubject where numSubjectID= " & txtMainRequestId.Text
        Else
            objdb.CreateNewConnection mCnn, enuSourceString.Soochika
            mSql = "Select chvSubject from tblSubjectCoding where intSubID= " & txtMainRequestId.Text
        End If
        
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF Or Rec.BOF) Then
            txtSubject.Text = Rec!chvSubject
            GetSevanaSubType
        Else
            MsgBox "Invalid subject id", vbInformation
            txtMainRequestId.Text = ""
        End If
    End If
End Sub
Public Sub GetSevanaSubType()
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    If gbLinkWithSoochika <> 1 Then
    objdb.CreateNewConnection mCnn, enuSourceString.SoochikaUnicode
    mSql = "select  mSubject.numSubjectID as intsubid,mSubjectSuite.numSubjectSuiteID as intMainsubID from mSubject Inner Join mSubjectSuite on mSubjectSuite.numSubjectID=mSubject.numSubjectID where mSubject.chvsubject='" & txtSubject.Text & "'"
    Else
    objdb.CreateNewConnection mCnn, enuSourceString.Soochika
    mSql = "select intsubid,intMainsubID from tblsubjectcoding where chvsubject='" & txtSubject.Text & "'"
    End If
    Rec.Open mSql, mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        txtMainRequestId.Text = Rec!intSubID
        SevanaMainSubid = IIf(IsNull(Rec!intMainSubID), 0, Rec!intMainSubID)
        If IsNull(Rec!intMainSubID) = False Then
        
        If gbLinkWithSoochika <> 1 Then
            PopulateList cmbRequestSubType, "select TypeofSubRequest,intID from mSubjectSevanaSubtype where intSubTypeID=" & Rec!intMainSubID & "", , False, True, True, enuSourceString.SoochikaUnicode
        Else
            PopulateList cmbRequestSubType, "select TypeofSubRequest,intID from tblSubjectSubType where intSubTypeID=" & Rec!intMainSubID & "", , False, True, True, enuSourceString.Soochika
        End If
            cmbRequestSubType.Enabled = True
            txtSubRequestId.Enabled = True
            cmbHospital.Enabled = True
            txtRemarks.Enabled = True
        Else
            txtSubRequestId.Enabled = False
            cmbRequestSubType.Enabled = False
            cmbHospital.Enabled = False
            txtRemarks.Enabled = False
        End If
    End If
End Sub
Private Sub txtStartInwardNo_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
           KeyAscii = 0
    End If
End Sub

Private Sub txtSubject_KeyPress(KeyAscii As Integer)
    If txtSubject.Text <> "" Then
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        lstSubject.Clear
        objdb.CreateNewConnection mCnn, enuSourceString.Soochika
        mSql = "select chvSubject from tblsubjectcoding where chvsubject like '%" & txtSubject.Text & "%'"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
            While Not Rec.EOF
                lstSubject.AddItem (Rec!chvSubject)
                Rec.MoveNext
            Wend
            lstSubject.Visible = True
        Else
            lstSubject.Visible = False
        End If
    End If
End Sub

Private Sub txtSubject_LostFocus()
    'lstSubject.Visible = False
End Sub

Private Sub txtSubRequestId_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
           KeyAscii = 0
    End If
End Sub

Private Sub txtSubRequestId_LostFocus()
'    Dim intLstIdx As Integer
'    If Trim(txtSubRequestId.Text) <> "" Then
'        If Val(Trim(txtSubRequestId.Text)) <> -1 Then
'            intLstIdx = GetListIndex(Val(Trim(txtSubRequestId)), cmbRequestSubType)
'            If intLstIdx <> -100 Then
'                cmbRequestSubType.ListIndex = intLstIdx
'            Else
'                MsgBox "Invalid Sub Request ID", vbInformation
'                cmbRequestSubType.ListIndex = 0
'                txtSubRequestId.SetFocus
'                Exit Sub
'            End If
'        End If
'    End If

' Modified By Akheel on 18/08/2009
    Dim flag
    flag = 0
    If txtSubRequestId.Text <> "" Then
        For i = 0 To cmbRequestSubType.ListCount - 1
        If cmbRequestSubType.ItemData(i) = val(txtSubRequestId.Text) Then
            cmbRequestSubType.ListIndex = i
                flag = 1
                GetKiosk (txtSubRequestId.Text)
            End If
        Next
        If flag <> 1 Then
            MsgBox "Item not found", vbDefaultButton1
        End If
    End If
'    If txtSubTypeID.Text = "2" Or txtSubTypeID.Text = "3" Then
'        Label2.Caption = "Arrival Date"
'    Else
'        Label2.Caption = "Application Date"
'    End If
    
End Sub

Private Sub txtTotalItems_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
           KeyAscii = 0
    End If
End Sub

Private Sub txtTotalItems_LostFocus()
    Dim i As Integer, intTotal As Integer, InwardNo As Long
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    
    If gbLinkWithSoochika <> 1 Then
    
    If objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False Then
        MsgBox "Soochika Connection Failed", vbDefaultButton1
        Exit Sub
        Else
       Rec.Open "select isnull(max(numCurrentNo),1) as MaxFileID from tInwardDetails where year(dtDateofReceipt)=year(getdate())", mCnn
        
    End If
    Else
    objdb.CreateNewConnection mCnn, enuSourceString.Soochika
    Rec.Open "select isnull(max(fldCurrentNo),0)+1 as MaxFileID from TblTappalDetails where year(Flddateofreceipt)=year(getdate())", mCnn
    End If
    
    If val(txtTotalItems) > 0 Then
        If flgInward = 0 Then
            intTotal = val(txtTotalItems)
            InwardNo = Rec!MaxFileID ' Val(txtStartInwardNo)
            txtStartInwardNo.Text = InwardNo
            vsfInward.Rows = intTotal + 1
            Me.MousePointer = vbHourglass
            For i = 1 To intTotal
                vsfInward.TextMatrix(i, 0) = i
    '            Do Until InwardNoExist(InwardNo, dtpInwDate.Value) = False
    '                InwardNo = InwardNo + 1
    '            Loop
                vsfInward.TextMatrix(i, 1) = i 'InwardNo
                'mCnn.Execute "spstartupProcess " & InwardNo & ",1"
                'InwardNo = InwardNo + 1
                gSubSetFont vsfInward, 1, 2, vsfInward.Rows - 1, vsfInward.Cols - 1, "Verdana"
            Next i
            Me.MousePointer = vbNormal
            'flgInward = 1
        End If
    Else
        MsgBox "Please enter how many number of items... ", vbInformation, "Bulk Inward Registration"
        txtTotalItems.SetFocus
        flgInward = 0
    End If
    
End Sub

Private Sub vsfInward_KeyPress(KeyAscii As Integer)
    If vsfInward.Col = 2 Then
        If KeyAscii <> 8 Then
            vsfInward.TextMatrix(vsfInward.Row, 2) = vsfInward.TextMatrix(vsfInward.Row, 2) + UCase(Chr(KeyAscii))
        End If
        If KeyAscii = 8 Then
            If Len(vsfInward.TextMatrix(vsfInward.Row, vsfInward.Col)) > 0 Then
                vsfInward.TextMatrix(vsfInward.Row, vsfInward.Col) = mID(vsfInward.TextMatrix(vsfInward.Row, vsfInward.Col), 1, Len(vsfInward.TextMatrix(vsfInward.Row, vsfInward.Col)) - 1)
            End If
        End If
    End If
End Sub

Public Function InwardNoExist(ByVal InwardNo As Long, ByVal inwdate As Date) As Boolean
    Dim Y As Variant
    Y = Year(inwdate)
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim objdb As New clsDB
    If objdb.CreateNewConnection(con, enuSourceString.SevanaRegn) = False Then
        MsgBox "Sevana Connection Failed", vbDefaultButton1
        Exit Function
    End If
    'rs.Open "SP_CheckInwardNoExist '" & InwardNo & "','" & Y & "'", con, adOpenDynamic, adLockOptimistic
    rs.Open "SP_CheckInwardNoExist '" & InwardNo & "','" & Y & "'", con
    If Not rs.EOF Then
        If rs(0) > 0 Then
            InwardNoExist = True
        Else
            InwardNoExist = False
        End If
    End If
    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
End Function

Public Function ValidateValues() As Boolean
Dim inwarddate As Variant
    ValidateValues = False
'    If cmbRequestType.ListIndex < 0 Then
'        MsgBox "Select the inward type", vbInformation
'        cmbRequestType.SetFocus
'        Exit Function
'    End If
    If Trim(txtSubject.Text) = "" Then
        MsgBox "Please enter subject ", vbInformation
        txtSubject.SetFocus
        Exit Function
    End If
    If SevanaMainSubid <> 0 Then
        If cmbRequestSubType.ListIndex < 0 Then
            MsgBox "Select the inward sub type", vbInformation
            cmbRequestSubType.SetFocus
            Exit Function
        End If
        If cmbHospital.ListIndex < 0 Then
           MsgBox "Please select the hospital.. ", vbInformation
           If cmbHospital.Enabled = True Then cmbHospital.SetFocus
           Exit Function
        End If
    End If
    Dim intRow As Integer
    For intRow = 1 To vsfInward.Rows - 1
        If Trim(vsfInward.TextMatrix(intRow, 2)) = "" Then
            MsgBox "Please enter the name ", vbInformation
            Exit Function
        End If
    Next intRow
    For intRow = 1 To vsfInward.Rows - 1
        If Trim(vsfInward.TextMatrix(intRow, 3)) = "" Then
            MsgBox "Please enter the Locality ", vbInformation
            Exit Function
        End If
    Next intRow
'
'    If cmbForwardTo.ListIndex <= 0 Then
'        MsgBox "Please select the forwarded point.. ", vbInformation
'        If cmbForwardTo.Enabled = True Then cmbForwardTo.SetFocus
'        Exit Function
'    End If
   
    If Not val(txtStartInwardNo) > 0 Then
        MsgBox "Please enter the starting Inward No", vbInformation
        Exit Function
    End If
    If Not val(txtTotalItems) > 0 Then
        MsgBox "Please enter the total Inwards", vbInformation
        Exit Function
    End If
    If cmbDept.ListIndex < 0 Then
        MsgBox "Please select department", vbInformation
        Exit Function
    End If
    If cboSeat.ListIndex < 0 Then
        MsgBox "Please select Seat ", vbInformation
        Exit Function
    End If
    inwarddate = Format(dtpInwDate.value, "DD-MMM-YYYY")
'    For intRow = 1 To vsfInward.Rows - 1
'     If lFunCheckInwardNoExisistance(Val(vsfInward.TextMatrix(intRow, 1)), inwarddate) = True Then
'      MsgBox "Inward No. already exists..Can't continue ..", vbInformation
'      txtStartInwardNo.SetFocus
'      Exit Function
'      End If
'     Next intRow
    ValidateValues = True
End Function

Public Sub DisCtls()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Enabled = False
        End If
    Next ctl
End Sub
Public Sub EnCtls()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Enabled = True
        End If
    Next ctl
    dtpInwDate.Enabled = False
End Sub




