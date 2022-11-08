VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmMonthlyUpdation 
   Caption         =   "Monthly Updation"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   4365
   Begin MSComctlLib.ProgressBar pbAccHead 
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   1470
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3660
      Top             =   1890
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   1
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   3180
      TabIndex        =   8
      Top             =   1830
      Width           =   1095
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmMonthlyUpdation.frx":0000
      Left            =   1170
      List            =   "frmMonthlyUpdation.frx":002E
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1020
      Width           =   3105
   End
   Begin VB.ComboBox cmbFinancialYear 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   570
      Width           =   3075
   End
   Begin VB.ListBox lstLocalBody 
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2070
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton cmdLocalBody 
      Caption         =   "..."
      Height          =   285
      Left            =   3900
      TabIndex        =   2
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox txtLocalBody 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   180
      Width           =   2625
   End
   Begin VB.Label lblUpdation 
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   1875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   195
      Left            =   615
      TabIndex        =   6
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Financial Year"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Local Body"
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   795
   End
End
Attribute VB_Name = "frmMonthlyUpdation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private Sub cmdLocalBody_Click()
        Dim objDB As New clsDB
        Dim mCon As New ADODB.Connection
        If objDB.SetConnection(mCon) Then
            PopulateList lstLocalBody, "SELECT vchLocalBody,intLocalBodyID FROM faLocalBody", , True, True, True, enuSourceString.Saankhya
            lstLocalBody.Visible = True
            lstLocalBody.SetFocus
        End If
    End Sub

    Private Sub cmdSave_Click()
        lSubSaveUpdations
        Call FormInitialize
    End Sub

    Private Sub Form_Activate()
        Me.Height = 2820
        Me.Width = 4710
        FormInitialize
    End Sub

    Private Sub Form_Load()
        lblUpdation.Caption = ""
        WindowsXPC.InitIDESubClassing
        PopulateList cmbFinancialYear, "SELECT intFinancialYear,intFinancialYearID FROM faFinancialYear", , True, True, True, enuSourceString.Saankhya
    End Sub

    Private Sub lstLocalBody_DblClick()
        If lstLocalBody.ListIndex > 0 Then
            txtLocalBody.Text = lstLocalBody.List(lstLocalBody.ListIndex)
            txtLocalBody.Tag = lstLocalBody.ItemData(lstLocalBody.ListIndex)
            lstLocalBody.Visible = False
        End If
    End Sub

    Private Sub lstLocalBody_GotFocus()
        lstLocalBody.Top = txtLocalBody.Top
        lstLocalBody.Left = 1500
        lstLocalBody.Height = 2000
        lstLocalBody.Width = 3000
        lstLocalBody.ZOrder (0)
    End Sub

    Private Sub lstLocalBody_LostFocus()
        lstLocalBody.Visible = False
        Me.Refresh
    End Sub
    
    Private Sub FormInitialize()
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            End If
            If TypeOf ctrl Is ComboBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
            End If
            If TypeOf ctrl Is ListBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
            End If
        Next
        pbAccHead.Visible = False
    End Sub
    
    Private Sub lSubSaveUpdations()
        On Error GoTo Err
        Dim mCon As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecUpdate As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mVarrIn As Variant
        Dim mVarrAccHead As Variant
        Dim mLoop As Integer
        Dim mMonthlyBalance As Double
        If Trim(txtLocalBody.Text) = "" Then
            MsgBox "Select LocalBody", vbInformation
            cmdLocalBody.SetFocus
            Exit Sub
        End If
        If cmbFinancialYear.ListIndex <= 0 Then
            MsgBox "Select Financial Year", vbInformation
            cmbFinancialYear.SetFocus
            Exit Sub
        End If
        If cmbMonth.ListIndex <= 0 Then
            MsgBox "Select Month", vbInformation
            cmbMonth.SetFocus
            Exit Sub
        End If
        If (objDB.SetConnection(mCon)) Then
            Set Rec = objDB.ExecuteSP("SELECT intAccountHeadID,vchAccountHead,vchAccountHeadCode FROM faAccountHeads", , , False, mCon, adCmdText)
            mVarrAccHead = Rec.GetRows
            pbAccHead.Value = 0
            pbAccHead.Max = UBound(mVarrAccHead, 2)
            If Rec.State = adStateOpen Then
                Rec.Close
            End If
            If IsArray(mVarrAccHead) Then
                mCon.BeginTrans
                pbAccHead.Visible = True
                For mLoop = 0 To UBound(mVarrAccHead, 2)
                    DoEvents
                
                    lblUpdation.Caption = CStr(CInt(mLoop / pbAccHead.Max * 100)) & " % Completed"
                    mVarrIn = Array(mVarrAccHead(0, mLoop), cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex), cmbMonth.ItemData(cmbMonth.ListIndex))
                    If RecUpdate.State = adStateOpen Then
                        RecUpdate.Close
                    End If
                    Set RecUpdate = objDB.ExecuteSP("spGetMonthlyAccountHeadBalance", mVarrIn, , , mCon, adCmdStoredProc)
                    If Not RecUpdate.EOF Then
                        mMonthlyBalance = RecUpdate.Fields(0)
                        RecUpdate.Close
                        mVarrIn = Array(Val(txtLocalBody.Tag), _
                                        cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex), _
                                        cmbMonth.ItemData(cmbMonth.ListIndex), _
                                        mVarrAccHead(0, mLoop), _
                                        mVarrAccHead(2, mLoop), _
                                        IIf((mMonthlyBalance >= 0), mMonthlyBalance, 0#), _
                                        IIf((mMonthlyBalance < 0), mMonthlyBalance * -1, 0#) _
                                        )
                        objDB.ExecuteSP "spSaveMonthlyHeadTotals", mVarrIn, , , mCon, adCmdStoredProc
                    End If
                    If pbAccHead.Value < pbAccHead.Max Then
                        pbAccHead.Value = pbAccHead.Value + 1
                    End If
                Next mLoop
                mCon.CommitTrans
                If mCon.State = 1 Then
                    mCon.Close
                End If
                MsgBox "Account head balance updated successfully", vbInformation, "Saankhya"
                lblUpdation.Caption = ""
                pbAccHead.Visible = False
            End If
        End If
        Exit Sub
Err:
    mCon.RollbackTrans
    mCon.Close
    MsgBox Error, vbCritical
    End Sub
