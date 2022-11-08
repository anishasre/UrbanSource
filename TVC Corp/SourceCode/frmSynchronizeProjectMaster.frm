VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSynchronizeProjectMaster 
   Caption         =   "Synchronize Project Master"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   1230
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdPort 
      Caption         =   "Port data from Sulekha"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   510
      TabIndex        =   0
      Top             =   420
      Width           =   2445
   End
End
Attribute VB_Name = "frmSynchronizeProjectMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    '*********************************************************************************************'
    '              Form to Synchronize the Project Master From DB_SulekhaFormulation                 '
    '*********************************************************************************************'
    Private Sub cmdPort_Click()
        Dim mCnnSulekha     As New ADODB.Connection
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim RecSulekha      As New ADODB.Recordset
        Dim RecFund         As New ADODB.Recordset
        Dim RecApproved     As New ADODB.Recordset
        Dim RecSource       As New ADODB.Recordset
        Dim RecSpillOver    As New ADODB.Recordset
        Dim mSql            As String
        Dim mSqlFund        As String
        Dim mSqlSource      As String
        Dim mArrIn          As Variant
        Dim mArrInput       As Variant
        Dim mArrInFund      As Variant
        Dim mProjectNameEng As String
        Dim mCount          As Variant
        Dim mCnt            As Variant
        Dim mFieldCount     As Variant
        Dim mYearID         As Variant
        Dim mOrderNo        As Variant
        Dim mOrderDate      As Variant
        Dim mStaus          As Variant
        
        '*********************************************************************************************'
        '                       Procedure to port data From DB_SulekhaFormulation                     '
        '*********************************************************************************************'
        mCount = ""
        If objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha) = True Then
            pbProgress.value = 0
            mSql = "Select Count(*) As Count From ProjectDetails Where tnyPhase=4 "
'            mSql = mSql + " Inner Join ProjectApproval On ProjectApproval.decProjectID = ProjectDetails.decProjectID"
'            mSql = mSql + " Inner Join ProjectSettings ON ProjectApproval.intSubmitId = ProjectSettings.intSubmitID AND ProjectSettings.intLBID = ProjectDetails.intLBID"
            RecSulekha.Open mSql, mCnnSulekha
            If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                mCount = IIf(IsNull(RecSulekha!count), "", RecSulekha!count)
            End If
            RecSulekha.Close
            mSql = ""
            mSql = " Select Count(*) as Count From SpecialPrograms"
            RecSulekha.Open mSql, mCnnSulekha
            If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                mCnt = IIf(IsNull(RecSulekha!count), "", RecSulekha!count)
            End If
            RecSulekha.Close
            If mCount <> "" Then
                pbProgress.Max = mCount + 1
            End If
            
'            mSql = "Select ProjectDetails.decProjectID,ProjectDetails.intLBID,ProjectDetails.intYearID,ProjectDetails.intProjectSlNo,ProjectDetails.chvProjectSlNo,ProjectDetails.chvProjectName,ProjectDetails.chvProjectNameEng,ProjectDetails.intProjCatID,chvOrderNo,chvOrderDate,intSecTypeID,ProjectDetails.intImplOfficerID,ProjectDetails.intMicroSecID,chvYear From ProjectDetails"
'            mSql = mSql + " Inner Join ProjectApproval On ProjectApproval.decProjectID = ProjectDetails.decProjectID"
'            mSql = mSql + " Inner Join ProjectSettings ON ProjectApproval.intSubmitId = ProjectSettings.intSubmitID AND ProjectSettings.intLBID = ProjectDetails.intLBID"
'            mSql = mSql + " Inner Join M_Year On ProjectDetails.intYearID = M_Year.intYearID"
            
            mSql = "Select isNull(tnyPhase,0) SpillOver,ProjectDetails.decProjectID,ProjectDetails.intLBID,ProjectDetails.intYearID,ProjectDetails.intProjectSlNo,ProjectDetails.chvProjectSlNo,ProjectDetails.chvProjectName," & vbNewLine
            mSql = mSql + " ProjectDetails.chvProjectNameEng,ProjectDetails.intProjCatID,intSecTypeID,ProjectDetails.intImplOfficerID,ProjectDetails.intMicroSecID,chvYear From ProjectDetails" & vbNewLine
            mSql = mSql + " Inner Join M_Year On ProjectDetails.intYearID = M_Year.intYearID " & vbNewLine
            ''------Added On 10/05/12 By Anisha
            mSql = mSql + " Where tnyPhase=4"
            ''----------------------------------
            RecSulekha.Open mSql, mCnnSulekha
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            While Not RecSulekha.EOF
            
                If RecSulekha!SpillOver <> 4 Then
''''                    mSqlFund = "Select * From FundDetails Where decProjectID = " & IIf(IsNull(RecSulekha!decProjectID), 0, RecSulekha!decProjectID)
''''                    mSqlFund = mSqlFund + " And intYearID = " & IIf(IsNull(RecSulekha!intYearID), 0, RecSulekha!intYearID)
''''                    RecFund.Open mSqlFund, mCnnSulekha
''''
''''                    If Not (RecFund.EOF And RecFund.BOF) Then
''''                        mCnn.Execute "Delete From suEstimation Where decProjectID =" & IIf(IsNull(RecSulekha!decProjectID), 0, RecSulekha!decProjectID)
''''                        For mFieldCount = 0 To RecFund.Fields.count - 1
''''                            mSqlSource = "Select intSourceFundID From suSourceOfFund Where vchSourceFundShortName = '" & mID(RecFund.Fields(mFieldCount).Name, 4) & "'"
''''                            'mSqlSource = mSqlSource + " And intGroupID =1"
''''                            If Not (IsNull(RecSulekha!chvYear)) Then
''''                                mYearID = mID(RecSulekha!chvYear, 1, 4)
''''                            End If
''''
''''                            RecSource.Open mSqlSource, mCnn
''''
''''                            If Not (RecSource.EOF And RecSource.BOF) Then
''''                                If RecFund.Fields(mFieldCount).value <> 0 Then
''''                                    mArrInFund = Array(-1, _
''''                                                    IIf(IsNull(RecSulekha!decProjectID), "", RecSulekha!decProjectID), _
''''                                                    mYearID, _
''''                                                    IIf(IsNull(RecSource!intSourceFundID), "", RecSource!intSourceFundID), _
''''                                                    RecFund.Fields(mFieldCount).value _
''''                                                    )
''''                                    objdb.ExecuteSP "spUpdateFundDetails", mArrInFund, , , mCnn, adCmdStoredProc
''''                                End If
''''                            End If
''''                        RecSource.Close
''''                        Next
''''                    End If
                Else
                '''Spill Over Project Synchronisation
                
                     objdb.CreateNewConnection mCnnSulekha, enuSourceString.Sulekha
                     mArrInFund = Array(IIf(IsNull(RecSulekha!decProjectID), 0, RecSulekha!decProjectID))
                     Set RecSpillOver = objdb.ExecuteSP("FundDetails_SpillOver", mArrInFund, , , mCnnSulekha, adCmdStoredProc)
                     
                     If Not (RecSpillOver.EOF And RecSpillOver.BOF) Then
                        mCnn.Execute "Delete From suEstimation Where decProjectID =" & IIf(IsNull(RecSulekha!decProjectID), 0, RecSulekha!decProjectID) & " AND intYearID=" & gbFinancialYearID
                        For mFieldCount = 0 To RecSpillOver.Fields.count - 1
                            mSqlSource = "Select intSourceFundID From suSourceOfFund Where vchSourceFundShortName = '" & mID(RecSpillOver.Fields(mFieldCount).Name, 4) & "'"
                            mYearID = gbFinancialYearID 'mID(RecSulekha!chvYear, 1, 4)
                            
                            RecSource.Open mSqlSource, mCnn
    
                            If Not (RecSource.EOF And RecSource.BOF) Then
                                If RecSpillOver.Fields(mFieldCount).value <> 0 Then
                                    mArrInFund = Array(-1, _
                                                    IIf(IsNull(RecSulekha!decProjectID), "", RecSulekha!decProjectID), _
                                                    mYearID, _
                                                    IIf(IsNull(RecSource!intSourceFundID), "", RecSource!intSourceFundID), _
                                                    RecSpillOver.Fields(mFieldCount).value _
                                                    )
                                    objdb.ExecuteSP "spUpdateFundDetails", mArrInFund, , , mCnn, adCmdStoredProc
                                End If
                            End If
                        RecSource.Close
                        Next
                    End If
                End If
                
                'RecFund.Close
                
                mSql = "Select ProjectDetails.decProjectID,ProjectDetails.intLBID,ProjectDetails.intYearID,ProjectDetails.intProjectSlNo,ProjectDetails.chvProjectSlNo,ProjectDetails.chvProjectName,ProjectDetails.chvProjectNameEng,ProjectDetails.intProjCatID,chvOrderNo,chvOrderDate,intSecTypeID,ProjectDetails.intImplOfficerID,ProjectDetails.intMicroSecID,chvYear From ProjectDetails"
                mSql = mSql + " Inner Join ProjectApproval On ProjectApproval.decProjectID = ProjectDetails.decProjectID"
                mSql = mSql + " Inner Join ProjectSettings ON ProjectApproval.intSubmitId = ProjectSettings.intSubmitID AND ProjectSettings.intLBID = ProjectDetails.intLBID"
                mSql = mSql + " Inner Join M_Year On ProjectDetails.intYearID = M_Year.intYearID"
                mSql = mSql + " Where ProjectDetails.decProjectID = " & IIf(IsNull(RecSulekha!decProjectID), 0, RecSulekha!decProjectID)
                 ''------Added On 10/05/12 By Anisha
                mSql = mSql + " And tnyPhase=4"
                RecApproved.Open mSql, mCnnSulekha
                If Not (RecApproved.EOF And RecApproved.BOF) Then
                    mOrderNo = IIf(IsNull(RecApproved!chvOrderNo), "", RecApproved!chvOrderNo)
                    mOrderDate = IIf(IsNull(RecApproved!chvOrderDate), "", RecApproved!chvOrderDate)
                    mStaus = 1
                Else
                    mOrderNo = ""
                    mOrderDate = ""
                    mStaus = 0
                End If
                RecApproved.Close
                
                mProjectNameEng = "ProjectNo " + RecSulekha!chvProjectSlNo
'                If Not (IsNull(RecSulekha!chvYear)) Then
'                    mYearID = mID(RecSulekha!chvYear, 1, 4)
'                End If
                mArrIn = Array(IIf(IsNull(RecSulekha!decProjectID), "", RecSulekha!decProjectID), _
                              IIf(IsNull(RecSulekha!intLBID), "", RecSulekha!intLBID), _
                              mYearID, _
                              IIf(IsNull(RecSulekha!intProjectSlNo), "", RecSulekha!intProjectSlNo), _
                              IIf(IsNull(RecSulekha!chvProjectSlNo), "", RecSulekha!chvProjectSlNo), _
                              IIf(IsNull(RecSulekha!chvProjectName), "", RecSulekha!chvProjectName), _
                              IIf(IsNull(RecSulekha!chvProjectNameEng), mProjectNameEng, RecSulekha!chvProjectNameEng), _
                              IIf(IsNull(RecSulekha!intProjCatID), "", RecSulekha!intProjCatID), _
                              mOrderNo, _
                              mOrderDate, _
                              IIf(IsNull(RecSulekha!intSecTypeID), "", RecSulekha!intSecTypeID), _
                              IIf(IsNull(RecSulekha!intImplOfficerID), "", RecSulekha!intImplOfficerID), _
                              IIf(IsNull(RecSulekha!intMicroSecID), "", RecSulekha!intMicroSecID), _
                              mStaus _
                            )
                objdb.ExecuteSP "spUpdateProjectDetails", mArrIn, , , mCnn, adCmdStoredProc
                If pbProgress.value < pbProgress.Max + 1 Then
                    pbProgress.value = pbProgress.value + 1
                End If
                RecSulekha.MoveNext
                
                
            Wend
            RecSulekha.Close
            '**********************************
'            Porting From SpecialPrograms of Sulekha To SuSpecialPrograms of Saankhya
            '**********************************
            objdb.CreateNewConnection mCnnSulekha, enuSourceString.Sulekha
            mSql = ""
            mSql = "Select * From SpecialPrograms"
            RecSulekha.Open mSql, mCnnSulekha
            While Not RecSulekha.EOF
                mArrInput = Array(IIf(IsNull(RecSulekha!decProjectID), "", RecSulekha!decProjectID), _
                                  IIf(IsNull(RecSulekha!intPlanID), "", RecSulekha!intPlanID))
                objdb.ExecuteSP "spUpdateSpecialPrograms", mArrInput, , , mCnn, adCmdStoredProc
                RecSulekha.MoveNext
            Wend
            RecSulekha.Close
            '**********************************
'            Porting From Implementing Officer of Sulekha To suImplementingOfficer of Saankhya
            '**********************************
            mSql = ""
            mSql = "Select intImplOfficerID,chvImplOfficerDesgEng,chrImplOfficerCode,intLBTypeID From M_ImplOfficer"
            RecSulekha.Open mSql, mCnnSulekha
            While Not RecSulekha.EOF
                mArrInput = Array(IIf(IsNull(RecSulekha!intImplOfficerID), "", RecSulekha!intImplOfficerID), _
                                IIf(IsNull(RecSulekha!chvImplOfficerDesgEng), "", RecSulekha!chvImplOfficerDesgEng), _
                                IIf(IsNull(RecSulekha!chrImplOfficerCode), "", RecSulekha!chrImplOfficerCode), _
                                IIf(IsNull(RecSulekha!intLBTypeID), "", RecSulekha!intLBTypeID) _
                                )
                objdb.ExecuteSP "spUpdateImplementingOfficer", mArrInput, , , mCnn, adCmdStoredProc
                RecSulekha.MoveNext
            Wend
            RecSulekha.Close
            mCnn.Close
            mCnnSulekha.Close
            MsgBox mCount & " Projects Updated", vbInformation
        Else
            MsgBox " Connection failed", vbInformation
        End If
        'MsgBox mCount & " rows Updated", vbInformation
        Unload Me
    End Sub
