VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   45
      TabIndex        =   7
      Top             =   -60
      Width           =   4395
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2430
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1110
         Width           =   1845
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   2430
         TabIndex        =   1
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtDB 
         Height          =   300
         Left            =   2295
         TabIndex        =   9
         Top             =   2175
         Width           =   1560
      End
      Begin VB.TextBox txtMac 
         Height          =   300
         Left            =   75
         TabIndex        =   8
         Top             =   2175
         Width           =   1845
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
         Height          =   390
         Left            =   2460
         TabIndex        =   6
         Top             =   1485
         Width           =   1830
      End
      Begin VB.TextBox txtIP 
         Height          =   300
         Left            =   2430
         TabIndex        =   3
         Top             =   780
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SQL Server Admin Pass Word"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1155
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Server Name"
         Height          =   195
         Left            =   1425
         TabIndex        =   0
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server's IP Address"
         Height          =   195
         Left            =   990
         TabIndex        =   2
         Top             =   825
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegister_Click()
   
        '1. Check IP Input and Pass Word
            txtIP.Text = Trim(txtIP)
            If txtIP.Text = "" Then
                MsgBox "Please Enter your DB_Finance Server IP", vbInformation
                txtIP.SetFocus
            End If
            If txtPassword.Text = "" Then
                MsgBox "Please Enter your Database Administration PassWord", vbInfoBackground
                txtPassword.SetFocus
            End If
        '2. Delete dsnF
            Call DSNDelete("dsnF", "SQL Server", True)
     
        '3. Create dsnF
            Dim mDSNDetails As tDSNAttrib
            Dim mErrString As String
            
            On Error GoTo SkipErr1:
            With mDSNDetails
                .Database = "DB_Finance"
                .Driver = "SQL Server"
                .Server = Trim(txtIP.Text)
                .TrustedConnection = False              'True = Use NT authentication
                .PassWord = txtPassword.Text
                .UserID = "sa"
                .Dsn = "dsnF"
                .Description = "Saankhay KMAS"
                .Type = ServerBased
                .SystemDSN = True                       'Create a System DSN
            End With
            mErrString = DSNCreate(mDSNDetails)
            If ErrString <> "" Then
                MsgBox "Failed to Connect the Server", vbInformation
                Exit Sub
            End If
        '4. GetMac
            Dim Rec As New ADODB.Recordset
            Dim objDb As New clsDB
            Dim mADOCnn As New ADODB.Connection
            Dim mArrOut As Variant
            
            mADOCnn.ConnectionString = "PROVIDER=MSDASQL; dsn=dsnF;uid=sa;pwd=" & txtPassword.Text & ";database=DB_Finance;"
            mADOCnn.Open
            On Error GoTo SkipErr2:
            Set Rec = objDb.ExecuteSP("spGetAddress", , mArrOut, , mADOCnn, adCmdStoredProc)
            If Not IsArray(mArrOut) Then
                MsgBox "Didn't able to find the Server: Missing [APR]", vbInformation
                Exit Sub
            End If
            Rec.Close
            txtMac.Text = mArrOut(0, 0)
        '5. GetDB version
            Dim mSQL As String
            mSQL = "SELECT ISNULL(Max(Restore_history_id),0) DBVersionID FROM msdb..restorehistory WHERE destination_database_name = 'DB_Finance'"
            Set Rec = GetRecordSet(mSQL, adOpenStatic, adLockReadOnly, mADOCnn)
            If (Rec.BOF And Rec.EOF) Then
                MsgBox "Didn't able to get the DB Version", vbInformation
                Exit Sub
            End If
            txtDB.Text = Rec!DBVersionID
            Rec.Close
            
        '6. Save Registration Details
            Dim vcbSVR_1        As Variant   '[varbinary](100)
            Dim vcbSvrIP_2      As Variant   '[varchar](50)
            Dim vcbDB_3         As Variant   '[varbinary](50)
            Dim dtRegDate_4     As Variant   '[datetime]
            Dim vchRegKey_5     As Variant   '[varchar](50)
            Dim vchLicenceKey_6 As String   '[varchar](250)
            
            Dim objMD           As New clsMD5
            Dim arrInput        As Variant  ' INPUT ARRAY
            vcbSVR_1 = Trim(txtServer.Text)
            vcbSvrIP_2 = Trim(txtIP.Text)
            vcbDB_3 = val(txtDB.Text)
            dtRegDate_4 = Date
            vchRegKey_5 = Null
            vchLicenceKey_6 = gbLocalBodyID & "-" & gbLocationID
            vchLicenceKey_6 = objMD.DigestFileToHexStr(vchLicenceKey_6)
            
            arrInput = Array(vcbSVR_1, vcbSvrIP_2, vcbDB_3, dtRegDate_4, vchRegKey_5, vchLicenceKey_6)
            objDb.ExecuteSP "spUpdateLicence", arrInput, , , mADOCnn, adCmdStoredProc
            cmdRegister.Enabled = False
        Exit Sub
SkipErr1:
        MsgBox " Didn't able to Establish ODBC Connection", vbInformation
        Exit Sub
SkipErr2:
        MsgBox "Error in Reading Address", vbInformation
        Exit Sub
End Sub

