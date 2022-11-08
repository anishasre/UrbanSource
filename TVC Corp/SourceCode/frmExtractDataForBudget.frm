VERSION 5.00
Begin VB.Form frmExtractDataForBudget 
   Caption         =   "Extract data for Budget"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
   Icon            =   "frmExtractDataForBudget.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4470
   Begin VB.CommandButton cmdExtractProjectData 
      Caption         =   "Extract Project Data"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   788
      TabIndex        =   1
      Top             =   1215
      Width           =   2895
   End
   Begin VB.CommandButton cmdExtractDataFromSaankhyaDB 
      Caption         =   "Extract data from Saankhya DB"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   788
      TabIndex        =   0
      Top             =   675
      Width           =   2895
   End
End
Attribute VB_Name = "frmExtractDataForBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '*********************************************************************************************'
    '        Form to extract data from Saankhya & Sulekha DB for Budget Preparation               '
    '*********************************************************************************************'
    
    Private Sub cmdExtractDataFromSaankhyaDB_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL    As String
        
        '*********************************************************************************************'
        '               Form to extract data from Saankhya DB for Budget Preparation                  '
        '*********************************************************************************************'
        On Error GoTo err
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "Insert Into faFunctionwiseTransactions(intFunctionID,intAccountHeadID,intTransactionTypeID,fltAmount,tnyFlag,intYearID)"
            mSQL = mSQL + " Select faTransactionChildForBudget.intFunctionID,intAccountHeadID,intTransactionTypeID,Sum(fltAmount)As fltAmount,0,intYearID"
            mSQL = mSQL + " From faTransactionChildForBudget"
            mSQL = mSQL + " LEFT JOIN faFunctionaryFunctions on faTransactionChildForBudget.intFunctionID =faFunctionaryFunctions.intFunctionID"
            mSQL = mSQL + " Where tnyVoucherGroupID = 10"
            mSQL = mSQL + " AND NOT faTransactionChildForBudget.intFunctionID is Null"
            mSQL = mSQL + " AND not faFunctionaryFunctions.intFunctionID is Null"
            mSQL = mSQL + " Group By faTransactionChildForBudget.intFunctionID,intAccountHeadID,intTransactionTypeID,tnyFlag,intYearID"
            mSQL = mSQL + " Order By faTransactionChildForBudget.intFunctionID,intAccountHeadID"
            mCnn.Execute mSQL
            cmdExtractDataFromSaankhyaDB.Enabled = False
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdExtractProjectData_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL    As String
        '*********************************************************************************************'
        '                   Form to extract data from Sulekha DB for Budget Preparation               '
        '*********************************************************************************************'
        On Error GoTo err
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "Insert Into faFunctionwiseTransactions(intFunctionID,intAccountHeadID,intTransactionTypeID,fltAmount,tnyFlag,intYearID)"
            mSQL = mSQL + " Select intFunctionID,intAccountHeadID,Null,Sum(fltEstAmt),1,suProjectDetails.intYearID From suProjectDetails"
            mSQL = mSQL + " Inner Join suEstimation On suProjectDetails.decProjectID = suEstimation.decProjectID"
            mSQL = mSQL + " Inner Join faMicroFunctionHeads On suProjectDetails.intMicroSectorID = faMicroFunctionHeads.intMicroSectorID"
            mSQL = mSQL + " Inner Join faFunctions On faMicroFunctionHeads.vchMicroFunctionCode = faFunctions.vchFunctionCode"
            mSQL = mSQL + " Inner Join faAccountHeads On faMicroFunctionHeads.vchMicroHeadCode = faAccountHeads.vchAccountHeadCode"
            mSQL = mSQL + " Where suProjectDetails.intMicroSectorID In(Select intMicroSectorID From faMicroFunctionHeads)"
            mSQL = mSQL + " Group By intFunctionID,intAccountHeadID,suProjectDetails.intYearID"
            mSQL = mSQL + " Order By intFunctionID,intAccountHeadID"
            mCnn.Execute mSQL
            cmdExtractProjectData.Enabled = False
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
        Me.Width = 4590
        Me.Height = 3045
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL    As String
        Dim Rec     As New ADODB.Recordset
        
        On Error GoTo err
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
        '*********************************************************************************************'
        '        Inserting data to faTransactionChildForBudger to avoid changes                       '
        '        in Original transaction table(faTransactionChild) for Budget Preparation             '
        '*********************************************************************************************'
            mSQL = "Select * From faTransactionChildForBudget"
            Rec.Open mSQL, mCnn
            If (Rec.EOF And Rec.BOF) Then
                mSQL = "Insert Into faTransactionChildForBudget"
                mSQL = mSQL + " Select faTransactions.intTransactionID,dtTransactionDate,intSerialNo,intFunctionID,intAccountHeadID,intTransactionTypeID,tnyVoucherGroupID,intFinancialYearID,0,fltAmount"
                mSQL = mSQL + " From faTransactions"
                mSQL = mSQL + " Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID"
                mSQL = mSQL + " Where faTransactions.tnyStatus <> 4 Or faTransactions.tnyStatus Is Null"
                mSQL = mSQL + " And dtTransactionDate Between '01-Apr-2010' And '31-Dec-2010'"
                mCnn.Execute mSQL
            End If
            Rec.Close
            
            mSQL = "Select * From faFunctionwiseTransactions Where  tnyFlag = 0"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                cmdExtractDataFromSaankhyaDB.Enabled = False
            Else
                cmdExtractDataFromSaankhyaDB.Enabled = True
            End If
            Rec.Close
            
            mSQL = "Select * From faFunctionwiseTransactions Where  tnyFlag = 1"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                cmdExtractProjectData.Enabled = False
            Else
                cmdExtractProjectData.Enabled = True
            End If
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
