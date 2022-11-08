VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmChangeDate 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                   ~~ C h a n g e  D a t e ~~"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   4335
      Top             =   1980
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00DAF2F2&
      Caption         =   "&Change"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   780
   End
   Begin VB.TextBox txtNewDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1545
      TabIndex        =   3
      Top             =   1005
      Width           =   1620
   End
   Begin MSComCtl2.DTPicker dtpNewDate 
      Height          =   330
      Left            =   3150
      TabIndex        =   4
      Top             =   1005
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   582
      _Version        =   393216
      Format          =   62390273
      CurrentDate     =   39548
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   690
      TabIndex        =   2
      Top             =   1050
      Width           =   795
   End
   Begin VB.Label lblCurrentDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31-Mar-2008"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1545
      TabIndex        =   1
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   405
      TabIndex        =   0
      Top             =   660
      Width           =   1080
   End
End
Attribute VB_Name = "frmChangeDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Option Explicit
    
    Private Sub cmdChange_Click()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New Recordset
        Dim arrOutPut As Variant
        Dim mSQL As String
        Dim mFromDate As String
        Dim mToDate As String
        Dim mServerDate As String
        
        mFromDate = gbStartingDate  'Format(DateAdd("m", -1, Now), "dd-mmm-yyyy")
        mToDate = Format(gbTransactionDate, "dd-mmm-yyyy")
        'mFromDate = Format(DateAdd("d", -10, Now), "dd-mmm-yyyy")
        
        '----------------------------------------------------------------------------------'
        ' Added by Vinod on 26-Mar-2011                                                  '
        '----------------------------------------------------------------------------------'
            objDB.SetConnection mCnn
            Set Rec = mCnn.Execute("Select GetDate()")
            If IsDate(Rec.Fields(0)) Then
                mServerDate = DdMmmYy(Rec.Fields(0))
            Else
                MsgBox "Didn't able to Access Server Date", vbInformation
                Exit Sub
            End If
            Rec.Close
            Set mCnn = Nothing
        
        '----------------------------------------------------------------------------------'
        ' Added by Anisha on 19-Jan-2010                                                  '
        '----------------------------------------------------------------------------------'
                mSQL = "Select intFinancialYearID From faFinancialYear Where  dtStartingDate<='" & Format(mFromDate, "dd/mmm/yy") & "' and dtEndingDate>='" & Format(mFromDate, "dd/mmm/yy") & "'"
                objDB.SetConnection mCnn
                Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
        '----------------------------------------------------------------------------------'
        If IsDate(txtNewDate) Then
            If Not (Rec.EOF Or Rec.BOF) Then
                'If (CDate(mFromDate) <= CDate(txtNewDate) And CDate(txtNewDate) <= CDate(mToDate)) Then   ' And Rec!intFinancialYearID = gbFinancialYearID Then
                If (CDate(mFromDate) <= CDate(txtNewDate) And CDate(txtNewDate) <= CDate(mServerDate)) Then   ' And Rec!intFinancialYearID = gbFinancialYearID Then
                    gbTransactionDate = txtNewDate.Text
                    lblCurrentDate.Caption = DdMmmYy(gbTransactionDate)
                    If gbTransactionDate <> mServerDate Then
                        frmMenu.lblTransactionDate.Caption = DdMmmYy(gbTransactionDate)
                        frmMenu.Timer1.Enabled = True
                    Else
                        frmMenu.lblTransactionDate.Caption = DdMmmYy(gbTransactionDate)
                        frmMenu.Timer1.Enabled = False
                        frmMenu.imgWarning.Visible = False
                        frmMenu.lblSplash.Visible = False
                    End If
                Else
                    'MsgBox ("Please Enter valid date:: One Month less than Actual date ")
                    MsgBox ("Please Enter valid date:: Date must be with in this Financial Year & Current Date! ")
                End If
            Else
                    'MsgBox ("Please Enter valid date:: One Month less than Actual date ")
                    MsgBox ("Please Enter valid date:: Date must be with in this Financial Year & Current Date!  ")
            End If
        End If
        If Month(gbTransactionDate) < 4 Then
            gbFinancialYearID = Year(gbTransactionDate) - 1
        Else
            gbFinancialYearID = Year(gbTransactionDate)
        End If
    End Sub

    Private Sub dtpNewDate_CloseUp()
        Dim dtNewDate As Date
        dtNewDate = dtpNewDate.Value
        If dtNewDate >= gbStartingDate And dtNewDate <= gbEndingDate Then
            txtNewDate.Text = DdMmmYy(dtNewDate)
        Else
            MsgBox "Enter a valid Date", vbInformation
            txtNewDate.Text = ""
        End If
    End Sub
    
    Private Sub dtpNewDate_DropDown()
        If IsDate(lblCurrentDate.Caption) Then
            dtpNewDate.Value = lblCurrentDate.Caption
        End If
    End Sub
    
    Private Sub Form_Load()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim arrOutPut As Variant
        Dim mSQL As String
        
        XPC.InitSubClassing
        '----------------------------------------------------------------------------------'
        ' Blocked by Aiby on 14-May-2008                                                   '
        '----------------------------------------------------------------------------------'
        '        mSQL = "Select dtLastTransactionDate From faFinancialYear "
        '        objDB.SetConnection mCnn
        '        objDB.ExecuteSP mSQL, , arrOutput, , mCnn, adCmdText
        '        If IsArray(arrOutput) Then
        '            If IsDate(arrOutput(0, 0)) Then
        '                lblCurrentDate.Caption = CheckDateInMMM(CStr(arrOutput(0, 0)))
        '                gbTransactionDate = lblCurrentDate.Caption
        '            End If
        '        End If
        '----------------------------------------------------------------------------------'

        Call CheckLastPostingDate
        lblCurrentDate.Caption = DdMmmYy(gbTransactionDate)
        
    End Sub
    Private Sub CheckLastPostingDate()   '-----------------LAST POSTING VALIDATION------------------
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSQL As String
        Dim Rec As New Recordset
        Dim dtCurrentDate As Date
        
        Call SetgbLastPostingDate
        
        objDB.SetConnection mCnn
        mSQL = "Select GETDATE()CurrentDate From faFinancialYear "
        Set Rec = GetRecordSet(mSQL)
        If Not (Rec.BOF And Rec.EOF) Then
            dtCurrentDate = Format(Rec!currentdate, "dd-mmm-yyyy")
            If CDate(dtCurrentDate) <= CDate(gbLastPostingDate) Then
                MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
                cmdChange.Enabled = False
                Exit Sub
            End If
            
        End If
        
    End Sub
    Private Sub txtNewDate_LostFocus()
        If Trim(txtNewDate) <> "" Then
            txtNewDate.Text = CheckDateInMMM(txtNewDate.Text)
        End If
    End Sub
    
