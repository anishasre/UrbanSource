VERSION 5.00
Begin VB.Form frmSectionWiseTransactionTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmSectionWiseTransactionTypes"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   LinkTopic       =   "frmSectionWiseTransactionTypes"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSesarchBank 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9780
      TabIndex        =   10
      Top             =   75
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtBankName 
      Height          =   315
      Left            =   5490
      TabIndex        =   9
      Top             =   75
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-->"
      Height          =   495
      Left            =   4965
      TabIndex        =   6
      Top             =   2565
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "<--"
      Height          =   495
      Left            =   4935
      TabIndex        =   5
      Top             =   3585
      Width           =   495
   End
   Begin VB.ListBox lstTransactions 
      Height          =   4545
      Left            =   5475
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      Top             =   870
      Width           =   4665
   End
   Begin VB.ListBox lstSelectedTransactions 
      Height          =   4545
      Left            =   165
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   870
      Width           =   4740
   End
   Begin VB.CommandButton cmdRestoreDefault 
      Caption         =   "Restore Default"
      Height          =   525
      Left            =   5130
      TabIndex        =   2
      Top             =   5595
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   525
      Left            =   3510
      TabIndex        =   1
      Top             =   5595
      Width           =   1500
   End
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   375
      TabIndex        =   0
      Text            =   "cmbSection"
      Top             =   75
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Bank:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4905
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label2 
      Caption         =   "List of Transaction Types"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6390
      TabIndex        =   8
      Top             =   540
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Transactions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1005
      TabIndex        =   7
      Top             =   495
      Width           =   1995
   End
End
Attribute VB_Name = "frmSectionWiseTransactionTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private Sub FillTransactionTypesInList()
        '---------------------------------------------------------'
        ' This will fill the lstSelected TransactionTypes
        ' according to the Selected Section
        '---------------------------------------------------------'
         Dim mSql As String
         Dim mLoop As Integer
         Dim objDb As New clsDb
         Dim Rec As New Recordset
         Dim mCnn As New ADODB.Connection
         Dim objAcc As New clsAccounts
         
         objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
         
         mSql = " SELECT  vchtransactionType,faSectionWiseTransactionTypes.intTransactionTypeID,isnull(faTransactionType.vchBankHeadCode,0) vchBankHeadCode "
         mSql = mSql + "  FROM faSectionWiseTransactionTypes INNER JOIN "
         mSql = mSql + "  faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
         mSql = mSql + "  Where intGroupID = 10 And faSectionWiseTransactionTypes.intSectionID =  " & cmbSection.ItemData(cmbSection.ListIndex)
         mSql = mSql + "  AND tnyList =1 "
         mSql = mSql + " ORDER BY vchTransactionType "
         PopulateList lstSelectedTransactions, mSql, , , , True
         Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
         
'         If Not (Rec.BOF Or Rec.EOF) Then
'            If (Rec!vchBankHeadCode) <> 0 Then
'                txtBankName.Tag = val(Rec!vchBankHeadCode)
'                objAcc.SetAccounts CStr(val(Rec!vchBankHeadCode))
'                txtBankName.Text = CStr(Rec!vchBankHeadCode) + "  " + objAcc.AccountHead
'            End If
'         End If
         mSql = " SELECT Distinct vchtransactionType,faSectionWiseTransactionTypes.intTransactionTypeID "
         mSql = mSql + "  FROM faSectionWiseTransactionTypes INNER JOIN "
         mSql = mSql + "  faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
         mSql = mSql + "  Where intGroupID = 10 And faSectionWiseTransactionTypes.intSectionID <>  " & cmbSection.ItemData(cmbSection.ListIndex)
         mSql = mSql + "  AND tnyList =1 And faSectionWiseTransactionTypes.intTransactionTypeID not in(0"
         For mLoop = 0 To lstSelectedTransactions.ListCount - 1
            mSql = mSql + "," + CStr(lstSelectedTransactions.ItemData(mLoop))
         Next
         mSql = mSql + ")"
         mSql = mSql + " ORDER BY vchTransactionType "
         PopulateList lstTransactions, mSql, , , , True

         Rec.Close
    End Sub
   
    Private Sub FillSection()
        Dim mSql As String
        mSql = "SELECT vchSectionName,intSectionID from faSection "
        PopulateList cmbSection, mSql, , True, True, True
    End Sub
    
    Private Sub cmbSection_Click()
'        txtBankName.Tag = ""
'         txtBankName.Text = ""
        Call FillTransactionTypesInList
    End Sub
    Private Sub cmdAdd_Click()
        Dim mLoop As Integer
            For mLoop = 0 To lstTransactions.ListCount - 1
                If lstTransactions.Selected(mLoop) Then
                    lstSelectedTransactions.AddItem lstTransactions.List(mLoop)
                    lstSelectedTransactions.ItemData(lstSelectedTransactions.NewIndex) = lstTransactions.ItemData(mLoop)
                End If
            Next mLoop

        For mLoop = 0 To lstTransactions.ListCount - 1
            If mLoop > lstTransactions.ListCount - 1 Then
                Exit For
            End If
            If lstTransactions.Selected(mLoop) Then
                lstTransactions.RemoveItem (mLoop)
                mLoop = mLoop - 1
            End If
        Next mLoop
    End Sub
    Private Sub cmdRemove_Click()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDb
        Dim Rec As New Recordset
        Dim mSql As Variant
        Dim mArrIn As Variant
        Dim mLoop As Long
        Dim mDefault As Integer
        Dim mCount As Integer
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCount = lstSelectedTransactions.ListCount
        For mLoop = 0 To mCount
            If mLoop = lstSelectedTransactions.ListCount Then Exit For
            If lstSelectedTransactions.Selected(mLoop) Then
                lstSelectedTransactions.RemoveItem (mLoop)
                mCount = mCount - 1
                mLoop = mLoop - 1
            End If
        Next mLoop
    End Sub

    Private Sub cmdRestoreDefault_Click()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDb
        Dim Rec As New Recordset
        Dim mSql As Variant
        
        If SaveValidation = False Then Exit Sub
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " DELETE FROM faSectionWiseTransactionTypes  where tnyDefaultFlag= 0 ;"
        mSql = mSql + " Update faSectionWiseTransactionTypes set tnyList= 1  WHERE faSectionWiseTransactionTypes.intSectionID= " & cmbSection.ItemData(cmbSection.ListIndex)
        objDb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        MsgBox " Restore Settings Done"
    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDb
        Dim Rec As New Recordset
        Dim mArrIn As Variant
        Dim mLoop As Integer
        Dim mCount As Integer
        Dim mVarrIn As Variant
        Dim mVarOut As Variant
        Dim mSql As String
        
        If SaveValidation = False Then Exit Sub
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCnn.Execute " DELETE FROM  faSectionWiseTransactionTypes  Where faSectionWiseTransactionTypes.tnyDefaultFlag =0 AND faSectionWiseTransactionTypes.intSectionID = " & cmbSection.ItemData(cmbSection.ListIndex) '' & " AND faSectionWiseTransactionTypes.intTransactionTypeID = " & lstSelectedTransactions.ItemData(mLoop)
        mCnn.Execute " UPDATE faSectionWiseTransactionTypes SET tnyList =0 WHERE faSectionWiseTransactionTypes.intSectionID= " & cmbSection.ItemData(cmbSection.ListIndex) ''& " AND faSectionWiseTransactionTypes.intTransactionTypeID = " & lstSelectedTransactions.ItemData(mLoop)
'        If val(txtBankName.Tag) = 0 Then
'            MsgBox " Bank Not Found',vbInformation"
'            Exit Sub
'        End If
        For mLoop = 0 To lstSelectedTransactions.ListCount - 1
            mArrIn = Array((cmbSection.ItemData(cmbSection.ListIndex)), _
                        lstSelectedTransactions.ItemData(mLoop), _
                        0, 1)
            objDb.ExecuteSP "spSaveSectionWiseTransactionTypes", mArrIn, , , mCnn, adCmdStoredProc
'             '-------------------------------------------------------------------------------'
'             '                Updating  TransactionType                                      '
'             '-------------------------------------------------------------------------------'
'
'            mSql = " Update faTransactionType Set vchBankHeadCode = " & txtBankName.Tag & " "
'            mSql = mSql + " Where faTransactionType.intTransactionTypeID = " & lstSelectedTransactions.ItemData(mLoop)
'            mCnn.Execute mSql
        Next mLoop
   
        MsgBox "saved successfully", vbInformation
        Call FillTransactionTypesInList
    End Sub
'    Private Sub cmdSesarchBank_Click()
'        'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = 2"
'        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND  faAccountHeads.vchAccountHeadCode Like '450%'"
'        frmSearchAccountHeads.Show 1
'        If gbSearchID <> -1 Then
'            txtBankName.Text = CStr(gbSearchStr)
'            txtBankName.Tag = Left(gbSearchStr, 9)
'        End If
'        gbSearchStr = ""
'        gbSearchID = -1
'    End Sub
    Private Sub Form_Activate()
        Me.Left = 3500
        Me.Top = 1500
    End Sub
    
    Private Sub Form_Initialize()
        cmbSection.ListIndex = -1
        lstSelectedTransactions.ListIndex = -1
        lstTransactions.ListIndex = -1
    End Sub

    Private Sub Form_Load()
        Call FillSection
    End Sub
    Private Function SaveValidation() As Boolean
        If cmbSection.ListIndex = -1 Then
            SaveValidation = False
            MsgBox "Please select a Section"
        Else
            SaveValidation = True
        End If
'        If txtBankName.Text = "" Then
'            SaveValidation = False
'            MsgBox " Please select the Bank Name"
'        Else
'            SaveValidation = True
'        End If
    End Function







