VERSION 5.00
Begin VB.Form frmSoochikaStartup 
   Caption         =   "Startup Process"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   4635
      Begin VB.CommandButton btnOK 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "O K"
         Height          =   345
         Left            =   1980
         TabIndex        =   3
         Top             =   1980
         Width           =   1275
      End
      Begin VB.TextBox txtCurrentNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1020
         Width           =   2325
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enter the previous inward No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   300
         TabIndex        =   1
         Top             =   780
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmSoochikaStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    Dim mVarrIn As Variant
    Dim mVarrOut As Variant
    Dim objDb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim Maxinwardno As Variant
    If txtCurrentNo.Text = "" Then
        MsgBox "Please enter last inward number", vbInformation, "Soochika Initialisation"
        Exit Sub
    Else
        If gbSoochikaVer <> 5 Then
                ReDim mVarrIn(1)
                If (objDb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
                    MsgBox "Soochika Connection Failure", vbInformation, "Soochika"
                    Exit Sub
                End If
                mVarrIn(0) = val(txtCurrentNo.Text)
                mVarrIn(1) = 0
                Set Rec = objDb.ExecuteSP("spStartupProcess", mVarrIn, , , mCnn, adCmdStoredProc)
                objDb.ExecuteSP "spUpdateReason", , , , mCnn, adCmdStoredProc
                MsgBox "Successfully initialised the Soochika Application"
                Unload Me
                frmSoochikaInward.Show
        Else
            
                ReDim mVarrIn(5)
                If (objDb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
                    MsgBox "Soochika Connection Failure", vbInformation, "Soochika"
                    Exit Sub
                End If
                Set Rec = mCnn.Execute("select right('000000' +cast(isnull(max(numcurrentno),0) as varchar),6) as InwardNo,isnull(convert(varchar,max(dtdateofReceipt),103),getdate()) as DateofReceipt from tInwardDetails where year(dtdateofreceipt)=year(getdate())")
                If Not (Rec.EOF Or Rec.BOF) Then
                    Maxinwardno = Rec!InwardNo
                End If
                If (Maxinwardno < val(txtCurrentNo.Text)) Then
                    mVarrIn(0) = val(gbLBType)
                    mVarrIn(1) = val(gbLocalBodyID)
                    mVarrIn(2) = gbLBID
                    mVarrIn(3) = Year(Date)
                    mVarrIn(4) = val(txtCurrentNo.Text)
                    mVarrIn(5) = gbUserID
                    
                    Set Rec = objDb.ExecuteSP("SpSetFileID", mVarrIn, , , mCnn, adCmdStoredProc)
                    mCnn.Execute "update tInterruption set flgReason=0,dtSolveReason=getdate() where flgReason=1"
                    MsgBox "Successfully initialised the Soochika Application"
                    Unload Me
                    frmUSoochikaInward.cmdNew.Enabled = True
                    frmUSoochikaInward.Show
                Else
                    MsgBox "Inward number duplication !!!"
                End If
        End If
    End If
End Sub
