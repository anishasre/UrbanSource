VERSION 5.00
Begin VB.Form frmGST 
   Caption         =   "GST IdentificationNumber"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "frmGST"
   ScaleHeight     =   1935
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   630
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtGstIn 
      Height          =   375
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN:"
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Top             =   660
      Width           =   735
   End
End
Attribute VB_Name = "frmGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    Dim mSQL    As String
    Dim mCnn    As New ADODB.Connection
    Dim objDB   As New clsDB
    Dim arrIn   As Variant
    Dim mYear   As Integer
    Dim mActive As Integer
    Dim Last2 As String
    Dim str As String
        If chkActive.Value Then
            mActive = 1
        Else
            mActive = 0
        End If
'        If mActive = 1 Then
'
'        End If
        If txtGstIn.Text = "" Then
            MsgBox "Please enter GSTIN", vbCritical
            Exit Sub
        End If
        
        If Len(txtGstIn.Text) > 15 Or Len(txtGstIn.Text) < 15 Or Left$(LCase(Right$(txtGstIn.Text, 2)), 1) <> "z" Or Left$(txtGstIn.Text, 2) <> 32 Then
            MsgBox "Please Enter correct GSTIN", vbCritical
            Exit Sub
        End If
        If gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupAccountsClerk Then
            arrIn = Array(txtGstIn.Text, mActive, Format(DateTime.Now, "dd/mmm/yyyy"), Null, 0)
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            arrIn = Array(txtGstIn.Text, mActive, Format(DateTime.Now, "dd/mmm/yyyy"), Null, 1)
        Else
            MsgBox "Please Login as Accountant", vbApplicationModal
            Exit Sub
        End If
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        objDB.ExecuteSP "spSaveGSTIN", arrIn, , , mCnn, adCmdStoredProc
        Call FillGSTIN
        MsgBox "Saved Successfully", vbCritical
        'If Right$(txtGstIn.Text, 2) <> "z5" Or Right$(txtGstIn.Text, 2) <> "Z5" Then
        
End Sub
Private Sub FillGSTIN()
        Dim mSQL    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec    As New ADODB.Recordset
        Dim objDB   As New clsDB
        Dim arrIn   As Variant
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = "Select * from faGSTNo where tnyActive=1"
        Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        If Rec.RecordCount > 0 Then
            txtGstIn.Text = IIf(IsNull(Rec!vchGSTNo), "", Rec!vchGSTNo)
            chkActive.Value = Rec!tnyActive
            If Rec!tnyStatus = 1 Then
                cmdSave.Caption = "Approved"
                cmdSave.Enabled = False
            Else
                cmdSave.Enabled = True
            End If
            'cmdSave.Enabled = False
        End If
End Sub

Private Sub Form_Load()
    If gbSeatGroupID = gbSeatGroupChiefCashier Then
    
        cmdSave.Caption = "SAVE"
    ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSave.Caption = "Approve"
    End If
    Call FillGSTIN
End Sub
