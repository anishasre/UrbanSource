VERSION 5.00
Begin VB.Form frmPortSulekha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utility"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   Icon            =   "frmPortSulekha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPortFundDetailsFromSulekha 
      Caption         =   "SYNCHRONIZE DATA"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   765
      TabIndex        =   3
      Top             =   2745
      Width           =   2265
   End
   Begin VB.CommandButton cmdSyncProjectsToFinance 
      Caption         =   "Sync with Requisition"
      Height          =   570
      Left            =   765
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdExpense 
      Caption         =   "Port  Expense  Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   720
      TabIndex        =   1
      Top             =   1170
      Width           =   2265
   End
   Begin VB.CommandButton cmdPort 
      Caption         =   "Port  Requisition Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2265
   End
End
Attribute VB_Name = "frmPortSulekha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExpense_Click()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim objDB As New clsDB
    Dim mCnnSulekha As New ADODB.Connection
    Dim RecSulekha As New ADODB.Recordset

    If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
            
                mSQL = " Update ExpenseDetails SET intYearID = 2012, tnyTransfer = 0 WHERE intYearID <> 2012"
                mCnnSulekha.Execute mSQL
            
                mSQL = "SELECT   faVouchers.intlocalBodyID, faVouchers.intFinancialYearID,faAllotments.numProjectID,faAllotments.intSourceID,faVouchers.fltAmount,"
                mSQL = mSQL + " faPayOrder.intVoucherID,  ISNULL(faPayOrder.tnyCancelled, 0) tnyCancelled FROM  faPayOrder "
                mSQL = mSQL + " INNER JOIN  faVouchers   ON faPayOrder.intVoucherID=faVouchers.intVoucherID"
                mSQL = mSQL + " INNER  JOIN faAllotments  ON  faPayOrder.intAllotmentID=faAllotments.intID"
                mSQL = mSQL + " Where numProjectID <> 0 "
                mSQL = mSQL + " Order By  numProjectID"
                Rec.Open mSQL, mCnn
            
                If Not (Rec.EOF And Rec.BOF) Then
                While Not Rec.EOF
                    mSQL = "Select  intVoucherID,tnyTransfer ,ISNULL(tnyCancelation,0) tnyCancellation from ExpenseDetails Where intVoucherID=" & Rec!intVoucherID
                    RecSulekha.Open mSQL, mCnnSulekha
                    If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                        If RecSulekha!tnyTransfer = 1 Then                                                     ':: BOCKED by Aiby::: And IIf(IsNull(RecSulekha!tnyCancelation), 0, RecSulekha!tnyCancelation) = 0) Then
                            If Rec!tnyCancelled <> RecSulekha!tnyCancellation Then
                                mSQL = " UPDATE ExpenseDetails SET tnyCancelation = " & Rec!tnyCancelled & " , "
                                mSQL = mSQL + " tnyTransfer = 0 "
                                mSQL = mSQL + " WHERE intVoucherID=" & Rec!intVoucherID
                                objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                            End If
                        End If
                    Else
                        mSQL = "INSERT INTO ExpenseDetails"
                        mSQL = mSQL + " VALUES(" & Rec!intLocalBodyID & "," & Rec!intFinancialYearID & " ," & Rec!numProjectID & " , -1,  " & Rec!intSourceID & ", " & Rec!fltAmount & ",  " & Rec!intVoucherID & ", " & Rec!tnyCancelled & ", 0)"
                        objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                    End If
                    RecSulekha.Close
                    Rec.MoveNext
                Wend
                Rec.Close
                
                
                mSQL = "Select faPayOrder.intVoucherID From faPendingTaskRequest "
                mSQL = mSQL + " INNER JOIN faPayOrder On faPayOrder.intPayOrderID = faPendingTaskRequest.numDemandID"
                mSQL = mSQL + " WHERE intTaskID IN (7,11) "
                Rec.Open mSQL, mCnn
                
                If Not (Rec.EOF And Rec.BOF) Then
                While Not Rec.EOF
                    If IsNumeric(Rec!intVoucherID) Then
                        mSQL = " Update ExpenseDetails SET tnyTransfer = 0 WHERE intVoucherID = " & Rec!intVoucherID
                        mCnnSulekha.Execute mSQL
                    End If
                    Rec.MoveNext
                Wend
                End If
                Rec.Close

                
                MsgBox " Expense Details Ported Successfully", vbInformation, "Saankhya"
            Else
                MsgBox "No records exists", vbInformation, "Saankhya"
                Exit Sub
            End If
            'Rec.Close
            
            
            mSQL = " Update ExpenseDetails SET tnyCancelation = 1, tnyTransfer = 0 WHERE ISNULL(intVoucherID,0) = 0"
            mCnnSulekha.Execute mSQL

            
            mCnnSulekha.Close
        Else
            MsgBox "Connection to Sulekha doesnot exists", vbInformation, "Saankhya"
            Exit Sub
        End If
        mCnn.Close
    Else
        MsgBox "Connection to Finance doesnot exists", vbInformation, "Saankhya"
        Exit Sub
    End If
End Sub

Private Sub cmdPort_Click()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim mCnnSulekha    As New ADODB.Connection
        Dim RecSulekha     As New ADODB.Recordset
        
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                
                mSQL = "If NOT Exists(SELECT * FROM information_schema.columns WHERE table_name = 'RequisitionDetails' And column_name = 'tnyVerified') "
                mSQL = mSQL + " ALTER TABLE RequisitionDetails Add tnyVerified  tinyint Null; "
                mCnnSulekha.Execute mSQL
                
                mSQL = " Delete From RequisitionDetails WHERE intReqID IN (Select intReqID From RequisitionDetails Group by intReqID Having Count(*) > 1  )"
                mCnnSulekha.Execute mSQL
                 
                mSQL = " Update RequisitionDetails  Set tnyVerified = 0 WHERE tnyVerified <> 1 "
                mCnnSulekha.Execute mSQL
                
                mSQL = " Update RequisitionDetails SET intYearID = 2012 WHERE dtAllotmentDate Between '01-Apr-2012' AND '31-Mar-2013' AND intYearID = 2013 "
                mCnnSulekha.Execute mSQL
                
                mSQL = "SELECT  intID,vchAllotmentNo,ISNULL(dtAllotmentDate,dtRequisitionDate) dtAllotmentDate ,fltAuthorizedAmt,intSourceID,numProjectID,intLBID,intFinancialYearID,tnyStatus FROM faAllotments"
                mSQL = mSQL + " WHERE numProjectID <> 0 And intFinancialYearID=2012 And tnyStage = 2 And tnyStatus in(1,2) " 'And isnull(fltTotalAltReceived,0) <> 0 "
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    While Not Rec.EOF
                        'mSql = "SELECT vchAllotmentNo,decProjectID,tnyCancel,tnyTransfer FROM RequisitionDetails WHERE vchAllotmentNo = " & Rec!vchAllotmentNo
                        mSQL = "SELECT * FROM RequisitionDetails WHERE  intReqID = " & Rec!intID
                        'If Rec!intID = 100 Then Stop
                        RecSulekha.Open mSQL, mCnnSulekha
                        If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                            If RecSulekha!tnyVerified = 1 Then
                                GoTo Skip:
                            End If
                            If (RecSulekha!tnyTransfer = 1) Then
                                If Rec!tnyStatus <> RecSulekha!tnyCancel _
                                    Or Rec!vchAllotmentNo <> RecSulekha!vchAllotmentNo _
                                    Or Rec!numProjectID <> RecSulekha!decProjectID _
                                    Or Rec!fltAuthorizedAmt <> RecSulekha!fltAmt _
                                    Or Rec!dtAllotmentDate <> RecSulekha!dtAllotmentDate _
                                    Or Rec!intFinancialYearID <> RecSulekha!intYearID _
                                Then
                                    mSQL = " Update RequisitionDetails  Set " 'intReqID= " & Rec!intID & " ,"
                                    mSQL = mSQL + " vchAllotmentNo = '" & Rec!vchAllotmentNo & "',"
                                    mSQL = mSQL + "dtAllotmentDate=' " & IIf(IsNull(Rec!dtAllotmentDate), "", DdMmmYy(Rec!dtAllotmentDate)) & " '"
                                    mSQL = mSQL + ",fltAmt=" & Rec!fltAuthorizedAmt & " ,"
                                    mSQL = mSQL + "intFundSrcID =" & Rec!intSourceID & ","
                                    mSQL = mSQL + "decProjectID =" & Rec!numProjectID & ","
                                    'mSql = mSql + "intLBID=" & Rec!intLBID & ","
                                    mSQL = mSQL + " intYearID = " & Rec!intFinancialYearID & ","
                                    mSQL = mSQL + " tnyCancel = " & Rec!tnyStatus & ","
                                    mSQL = mSQL + " tnyTransfer = 0 ,"
                                    mSQL = mSQL + " tnyVerified = 1 "
                                    mSQL = mSQL + " Where intReqID = " & Rec!intID
                                    objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                                    
                                Else
                                    mSQL = " Update RequisitionDetails  Set "
                                    mSQL = mSQL + " tnyVerified = 1 "
                                    mSQL = mSQL + " Where intReqID = " & Rec!intID
                                    objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                                End If
Skip:
                            ElseIf (RecSulekha!tnyTransfer = 0) Then
                                  mSQL = " Update RequisitionDetails  SET "
                                  mSQL = mSQL + " vchAllotmentNo = '" & Rec!vchAllotmentNo & "',"
                                  mSQL = mSQL + "dtAllotmentDate=' " & IIf(IsNull(Rec!dtAllotmentDate), "", DdMmmYy(Rec!dtAllotmentDate)) & " '"
                                  mSQL = mSQL + ",fltAmt=" & Rec!fltAuthorizedAmt & " ,"
                                  mSQL = mSQL + "intFundSrcID=" & Rec!intSourceID & ","
                                  mSQL = mSQL + "decProjectID=" & Rec!numProjectID & ","
                                  mSQL = mSQL + "intLBID=" & Rec!intLBID & ","
                                  mSQL = mSQL + " intYearID = " & Rec!intFinancialYearID & ","
                                  mSQL = mSQL + "tnyCancel=" & Rec!tnyStatus & ", "
                                  mSQL = mSQL + "tnyVerified= 1 "
                                  'mSql = mSql + " Where vchAllotmentNo =" & Rec!vchAllotmentNo
                                  mSQL = mSQL + " Where intReqID = " & Rec!intID
                                  objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                            End If
                        Else
ADDNEW:
                                mSQL = "INSERT INTO RequisitionDetails "
                                mSQL = mSQL + " VALUES(" & Rec!intID & ",'" & IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo) & " ',' " & IIf(IsNull(Rec!dtAllotmentDate), Null, DdMmmYy(Rec!dtAllotmentDate)) & " ',"
                                mSQL = mSQL + " " & Rec!fltAuthorizedAmt & "," & Rec!intSourceID & "," & Rec!numProjectID & "," & Rec!intLBID & "," & Rec!intFinancialYearID & "," & Rec!tnyStatus & ",0, 1 )"
                                objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
                        End If
                        RecSulekha.Close
                        Rec.MoveNext
                    Wend
                    
                    mSQL = "Update RequisitionDetails  Set tnyCancel = 2, tnyTransfer = 0 WHERE ISNULL(tnyVerified,0) = 0"
                    mCnnSulekha.Execute mSQL
                    
                    MsgBox " Requisition Details Ported Successfully", vbInformation, "Saankhya"
                Else
                    MsgBox "No Record Exists", vbInformation, "Saankhya"
                    Exit Sub
                End If
                Rec.Close
                mCnnSulekha.Close
            Else
                MsgBox "Connection to Sulekha doesnot exists", vbInformation, "Saankhya"
                Exit Sub
            End If
            mCnn.Close
        Else
            MsgBox "Connection to Finance doesnot exists", vbInformation, "Saankhya"
            Exit Sub
     End If
End Sub
Private Sub cmdPortFundDetailsFromSulekha_Click()
    Call FnUpdateProjectMaster
    Call FnUpdateEstimationDetails
    MsgBox "UPDATED SUCCESSFULLY", vbInformation
    cmdPortFundDetailsFromSulekha.Enabled = False
End Sub
Private Sub FnUpdateProjectMaster()
    Dim mCn As New ADODB.Connection
    Dim mCnnSulekha As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecSulekha As New ADODB.Recordset
    Dim mArrIn As Variant
    Dim objDB As New clsDB
    Dim mSQL As String
    Dim mID As Long

    If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
        If objDB.SetConnection(mCn) Then
            mSQL = "Select ProjectDetails.*,SubjectCheckList.* from ProjectDetails"
            mSQL = mSQL + " Inner Join SubjectCheckList On SubjectCheckList.decProjectID=ProjectDetails.decProjectID"
            mSQL = mSQL + " Where intApproval = 7"
            RecSulekha.Open mSQL, mCnnSulekha
            If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                While Not RecSulekha.EOF
                    mSQL = ""
                    mSQL = " Select * From suProjectDetails Where decProjectID=" & RecSulekha!decProjectID & " And intYearID=" & RecSulekha!intYearID & " "
                    Rec.Open mSQL, mCn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mSQL = ""
                        mSQL = " Update suProjectDetails Set "
                        mSQL = mSQL + " intLBID =" & RecSulekha!intLBID & ", intProjectSlNo =" & RecSulekha!intProjectSlNo & ", "
                        mSQL = mSQL + " chvProjectSlNo = '" & RecSulekha!chvProjectSlNo & "',"   ', chvProjectName ='" & RecSulekha!chvProjectName & "'
                        mSQL = mSQL + " chvProjectnameEnglish ='" & RecSulekha!chvProjectNameEng & "', intProjCatID =" & RecSulekha!intProjCatID & ","
                        mSQL = mSQL + " chvDPCOrderNo = '" & IIf(IsNull(RecSulekha!nchApprovalNo), "", RecSulekha!nchApprovalNo) & "', dtDPCOrderDate ='" & IIf(IsNull(RecSulekha!dtApprovalDate), "", DdMmmYy(RecSulekha!dtApprovalDate)) & "', "
                        mSQL = mSQL + " intSectorTypeID =" & RecSulekha!intSubSecID & ",  intImplementingOfficerID = " & RecSulekha!intImplOfficerID & ", intMicroSectorID = " & RecSulekha!intMajorSecID & ","
                        mSQL = mSQL + " vchApproverFullName ='" & RecSulekha!chvFullName & "', vchApproverDesignation ='" & RecSulekha!chvDesignation & "'"
                        mSQL = mSQL + " Where decProjectID=" & RecSulekha!decProjectID & " And intYearID=" & RecSulekha!intYearID & " "
                        objDB.ExecuteSP mSQL, , , , mCn, adCmdText
                    Else
                        mSQL = ""
                        mSQL = "INSERT INTO suProjectDetails"
                        mSQL = mSQL + " (decProjectID, intLBID, intYearID, intProjectSlNo, chvProjectSlNo, chvProjectName, chvProjectnameEnglish, intProjCatID, chvDPCOrderNo,"
                        mSQL = mSQL + " dtDPCOrderDate, intSectorTypeID, intPlanID, intImplementingOfficerID, intMicroSectorID, tnyStatus, vchApproverFullName, vchApproverDesignation)"
                        mSQL = mSQL + " VALUES (" & RecSulekha!decProjectID & "," & RecSulekha!intLBID & "," & RecSulekha!intYearID & " "
                        mSQL = mSQL + " ," & RecSulekha!intProjectSlNo & ",'" & RecSulekha!chvProjectSlNo & "','" & RecSulekha!chvProjectName & "'"
                        mSQL = mSQL + " ,'" & RecSulekha!chvProjectNameEng & "'," & RecSulekha!intProjCatID & ",'" & RecSulekha!nchApprovalNo & "','" & DdMmmYy(RecSulekha!dtApprovalDate) & "'"
                        mSQL = mSQL + " ," & RecSulekha!intSubSecID & ",NULL," & RecSulekha!intImplOfficerID & "," & RecSulekha!intMajorSecID & ",7,'" & RecSulekha!chvFullName & "','" & RecSulekha!chvDesignation & "')"
                        objDB.ExecuteSP mSQL, , , , mCn, adCmdText
                    End If
                    Rec.Close
                    RecSulekha.MoveNext
                Wend
            Else
                MsgBox "No Record Exists", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "Connection to Finance Database Doesn't exists", vbInformation
        End If
        'MsgBox "UPDATED SUCCESSFULLY", vbInformation
        RecSulekha.Close
        mCnnSulekha.Close
        mCn.Close
    Else
        MsgBox "Connection to Sulekha Database Doesn't exists", vbInformation
        Exit Sub
    End If
End Sub
Private Sub FnUpdateEstimationDetails()
    Dim mCn As New ADODB.Connection
    Dim mCnnSulekha As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim RecSulekha As New ADODB.Recordset
    Dim mArrIn As Variant
    Dim objDB As New clsDB
    Dim mSQL As String
    Dim mID As Long
    

    If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
        If objDB.SetConnection(mCn) Then
            mID = 1
            mSQL = "Delete From suEstimation"
            objDB.ExecuteSP mSQL, , , , mCn, adCmdText
            mSQL = "Select * from FundDetails"
            RecSulekha.Open mSQL, mCnnSulekha
            If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                While Not RecSulekha.EOF
                        mSQL = ""
                        mSQL = "INSERT INTO suEstimation ( intID,decProjectID, intYearID, intFundID, fltEstAmt)"
                        mSQL = mSQL + " VALUES( " & mID & "," & RecSulekha!decProjectID & "," & RecSulekha!intYearID & "," & RecSulekha!intFundSrcID & "," & RecSulekha!fltAmt & ")"
                        mID = mID + 1
                        objDB.ExecuteSP mSQL, , , , mCn, adCmdText
                        RecSulekha.MoveNext
                Wend
            Else
                MsgBox "No Record Exists", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "Connection to Finance Database Doesn't exists", vbInformation
        End If
        'MsgBox "UPDATED SUCCESSFULLY", vbInformation
        RecSulekha.Close
        mCnnSulekha.Close
        mCn.Close
    Else
        MsgBox "Connection to Sulekha Database Doesn't exists", vbInformation
        Exit Sub
    End If
End Sub
Private Sub cmdSyncProjectsToFinance_Click()
            
            Dim mCn As New ADODB.Connection
            Dim mCnnSulekha As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim RecSulekha As New ADODB.Recordset
            Dim mArrIn As Variant
            Dim objDB As New clsDB
            Dim mSQL As String
            Dim mProjectID As Double
            
            
            
            If objDB.SetConnection(mCn) Then
                mSQL = " Select numProjectID From faAllotments "
                mSQL = mSQL + " LEFT JOIN suProjectDetails ON suProjectDetails.decProjectID = numProjectID "
                mSQL = mSQL + " Where IsNull(numProjectID, 0) <> 0 And decProjectID Is Null "
                Rec.Open mSQL, mCn, adOpenStatic, adLockReadOnly
                If Not (Rec.BOF And Rec.EOF) Then
                    If (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                    While Not Rec.EOF
                        mProjectID = Rec!numProjectID
                        mSQL = "Select * from ProjectDetails "
                        mSQL = mSQL + " left join SubjectCheckList On SubjectCheckList.decProjectID=ProjectDetails.decProjectID "
                        mSQL = mSQL + " Where ProjectDetails.decProjectID= " & mProjectID
                        RecSulekha.Open mSQL, mCnnSulekha
                        If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                            mArrIn = Array(mProjectID, _
                                      gbLBID, _
                                      gbFinancialYearID, _
                                      IIf(IsNull(RecSulekha!intProjectSlNo), "", RecSulekha!intProjectSlNo), _
                                      IIf(IsNull(RecSulekha!chvProjectSlNo), "", RecSulekha!chvProjectSlNo), _
                                      IIf(IsNull(RecSulekha!chvProjectName), "", RecSulekha!chvProjectName), _
                                      IIf(IsNull(RecSulekha!chvProjectNameEng), "", RecSulekha!chvProjectNameEng), _
                                      IIf(IsNull(RecSulekha!intProjCatID), "", RecSulekha!intProjCatID), _
                                      IIf(IsNull(RecSulekha!nchApprovalNo), "", RecSulekha!nchApprovalNo), _
                                      IIf(IsNull(RecSulekha!dtApprovalDate), "", RecSulekha!dtApprovalDate), _
                                      IIf(IsNull(RecSulekha!intSecID), "", RecSulekha!intSecID), _
                                      IIf(IsNull(RecSulekha!intImplOfficerID), "", RecSulekha!intImplOfficerID), _
                                      IIf(IsNull(RecSulekha!intSubSecID), "", RecSulekha!intSubSecID), _
                                      9, _
                                      IIf(IsNull(RecSulekha!chvFullName), "", RecSulekha!chvFullName), _
                                      IIf(IsNull(RecSulekha!chvDesignation), "", RecSulekha!chvDesignation), _
                                      Null _
                                    )
                            objDB.ExecuteSP "spUpdateProjectDetails", mArrIn, , , mCn, adCmdStoredProc
                        End If
                        RecSulekha.Close
                        Rec.MoveNext
                    Wend
                    End If
                End If
                
            End If 'Endif of mCn OPEN CONNECTION
            
End Sub
