VERSION 5.00
Begin VB.Form frmAFSFinalization 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AFS - Finanlization"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   Icon            =   "frmAFSFinalization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGO 
      Caption         =   "PROCEED"
      Height          =   315
      Left            =   10140
      TabIndex        =   6
      Top             =   1080
      Width           =   945
   End
   Begin VB.ComboBox cmbFinancialYear 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7140
      TabIndex        =   4
      Top             =   1020
      Width           =   2415
   End
   Begin VB.ComboBox cmbCategory 
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
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   990
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2250
      ScaleHeight     =   735
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ANNUAL FINANCIAL STATEMENTS-PREPARATION"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1710
         TabIndex        =   1
         Top             =   180
         Width           =   5670
      End
   End
   Begin VB.Label lblFinancialyear 
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year"
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
      Left            =   6030
      TabIndex        =   5
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
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
      Left            =   2310
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   0
      Picture         =   "frmAFSFinalization.frx":1CCA
      Stretch         =   -1  'True
      Top             =   45
      Width           =   2280
   End
End
Attribute VB_Name = "frmAFSFinalization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Private Sub cmdGo_Click()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    
    Dim mCate As Integer
    Dim mYear As Integer
    Dim mMonth As Integer
    Dim arrIn As Variant
    Dim arrInput As Variant
    Dim arrInputc As Variant
    Dim mSheduleGroup As String
    Dim mMajorCode As String
    Dim mMajorHead As String
    Dim mAccHeadType As String
    Dim mSheduleTitle As String
    Dim mAmt As Double
    Dim arrOutPut As Variant
    Dim mAFSID As Integer
    If cmbCategory.ListIndex > 0 Then
        mCate = cmbCategory.ItemData(cmbCategory.ListIndex)
    Else
        MsgBox "Please Select AFS Report"
        Exit Sub
    End If
    If cmbFinancialYear.ListIndex > 0 Then
        mYear = cmbFinancialYear.ItemData(cmbCategory.ListIndex)
    Else
        MsgBox "Please Select AFS Report"
        Exit Sub
    End If
  
    If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            If mCate = 1 Then
               mMonth = 4
                
                While mMonth < 12
                    arrIn = Array(mMonth, mYear)
                    Set Rec = objdb.ExecuteSP("SpGETAFSBL", arrIn, , , mCnn, adCmdStoredProc)
                    While Not (Rec.EOF)
   
                        If Not (Rec.EOF And Rec.BOF) Then
                           
                            mSheduleTitle = IIf(IsNull(Rec!vchScheduleTitle), "", Rec!vchScheduleTitle)
                            mSheduleGroup = IIf(IsNull(Rec!vchScheduleGroup), "", Rec!vchScheduleGroup)
                            mMajorCode = IIf(IsNull(Rec!vchMajorAccountHeadCode), "", Rec!vchMajorAccountHeadCode)
                            mMajorHead = IIf(IsNull(Rec!vchMajorAccountHead), "", Rec!vchMajorAccountHead)
                            mAccHeadType = IIf(IsNull(Rec!AccountHeadCode), "", Rec!AccountHeadCode)
                            mAmt = IIf(IsNull(Rec!transactionamount), "", Rec!transactionamount)
                            
                            'vsGrid.TextMatrix(mCnt, 0) = mRec!intAccountHeadID
                            'vsGrid.TextMatrix(mCnt, 0) = mRec!intAccountHeadID
                            
                        End If
                        Rec.MoveNext
                       Wend
                    
                        arrInput = Array(-1, mYear, mMonth)
                        Call objdb.ExecuteSP("spSaveAFSMonthly", arrInput, arrOutPut, , mCnn)
                        If IsNumeric(arrOutPut(0, 0)) Then
                            mAFSID = arrOutPut(0, 0)
                        Else
                           ' GoTo ErrRollBack:
                        End If
                        
                        arrInputc = Array(mAFSID, mMajorCode, mSheduleGroup, mSheduleTitle, mAccHeadType, mAmt, 0)
                        Call objdb.ExecuteSP("spSaveAFSMonthly", arrInput, arrOutPut, , mCnn)
                       
                    mMonth = mMonth + 1
                    Wend
             End If
    End If
End Sub
    Private Sub Form_Load()
        FillCat
        FillYear
    End Sub
    Private Sub FillCat()
      Dim mSql As String
       
      mSql = " Select 'Balance Sheett' cat ,1 catType Union All Select 'Income And Expenditure' cat , 2 catType "
      mSql = mSql + "Union All Select 'Receipt And Payment' cat,3 catType"
      Call PopulateList(cmbCategory, mSql, True, True, True, True)
      
    End Sub
    Private Sub FillYear()
        Dim mSql As String
        mSql = "Select Max(intYearID) From faAFS Where "
        mSql = " Select * From faFinancialYear Where intFinancialYearID>(Select intFinancialYearID From faVouchers Where intTransactionTypeiD=3000 group by intFinancialYearID)"
        Call PopulateList(cmbFinancialYear, mSql, , True, , True)
    End Sub
