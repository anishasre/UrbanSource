VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProceedings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
   ' Variable Declarations '
Private mintProceedingsID                         As Integer
Private mvchProceedingsNo                         As String
Private mdtProceedingsDate                        As String
Private mintApplicationID                         As Integer
Private mintModuleID                              As Integer
Private mintVoucherID                             As Variant
Private mintVoucherNo                             As Variant
Private mvchRemarks                               As String
Private mtnyUsed                                  As Integer
Private mtnyRemoved                               As Integer
   ' Property Definitions '
Public Property Let ProceedingsID(ByVal vData As Integer)
    mintProceedingsID = vData
End Property
Public Property Get ProceedingsID() As Integer
    ProceedingsID = mintProceedingsID
End Property
Public Property Let ProceedingsNo(ByVal vData As String)
    mvchProceedingsNo = vData
End Property
Public Property Get ProceedingsNo() As String
    ProceedingsNo = mvchProceedingsNo
End Property
Public Property Let ProceedingsDate(ByVal vData As String)
    mdtProceedingsDate = vData
End Property
Public Property Get ProceedingsDate() As String
    ProceedingsDate = mdtProceedingsDate
End Property
Public Property Let ApplicationID(ByVal vData As Integer)
    mintApplicationID = vData
End Property
Public Property Get ApplicationID() As Integer
    ApplicationID = mintApplicationID
End Property
Public Property Let ModuleID(ByVal vData As Integer)
    mintModuleID = vData
End Property
Public Property Get ModuleID() As Integer
    ModuleID = mintModuleID
End Property
Public Property Let VoucherID(ByVal vData As Variant)
    mintVoucherID = vData
End Property
Public Property Get VoucherID() As Variant
    VoucherID = mintVoucherID
End Property
Public Property Let VoucherNo(ByVal vData As Variant)
    mintVoucherNo = vData
End Property
Public Property Get VoucherNo() As Variant
    VoucherNo = mintVoucherNo
End Property
Public Property Let Remarks(ByVal vData As String)
    mvchRemarks = vData
End Property
Public Property Get Remarks() As String
    Remarks = mvchRemarks
End Property
Public Property Let Used(ByVal vData As Integer)
    mtnyUsed = vData
End Property
Public Property Get Used() As Integer
    Used = mtnyUsed
End Property
Public Property Let Removed(ByVal vData As Integer)
    mtnyRemoved = vData
End Property
Public Property Get Removed() As Integer
    Removed = mtnyRemoved
End Property
   ' Sub or Function Definitions   '
Public Sub getProceedingsByID()
       Dim objDB As New clsDB
       Dim mCnn As New ADODB.Connection
       Dim Rec As New ADODB.Recordset
       Dim mSQL As String
       '       Craeting Connction              '
       If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
           MsgBox "Connection to Saankhya Not Present", vbCritical
           Exit Sub
       End If
   mSQL = "Select * From faProceedings Where intProceedingsID = '" & mintProceedingsID & "'"
   Rec.Open mSQL, mCnn
   Call clearVariables
   If Not (Rec.EOF And Rec.BOF) Then
    mintProceedingsID = IIf(IsNull(Rec!intProceedingsID), -1, Rec!intProceedingsID)
    mvchProceedingsNo = IIf(IsNull(Rec!vchProceedingsNo), "", Rec!vchProceedingsNo)
    mdtProceedingsDate = IIf(IsNull(Rec!dtProceedingsDate), "", Rec!dtProceedingsDate)
    mintApplicationID = IIf(IsNull(Rec!intApplicationID), -1, Rec!intApplicationID)
    mintModuleID = IIf(IsNull(Rec!intModuleID), -1, Rec!intModuleID)
    mintVoucherID = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
    mintVoucherNo = IIf(IsNull(Rec!intVoucherNo), -1, Rec!intVoucherNo)
    mvchRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
    mtnyUsed = IIf(IsNull(Rec!tnyUsed), -1, Rec!tnyUsed)
    mtnyRemoved = IIf(IsNull(Rec!tnyRemoved), -1, Rec!tnyRemoved)
    End If
       Rec.Close
       mCnn.Close
End Sub
Public Sub getProceedingsByNo()
       Dim objDB As New clsDB
       Dim mCnn As New ADODB.Connection
       Dim Rec As New ADODB.Recordset
       Dim mSQL As String
       '       Craeting Connction              '
       If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
           MsgBox "Connection to Saankhya Not Present", vbCritical
           Exit Sub
       End If
   mSQL = "Select * From faProceedings Where vchProceedingsNo = '" & mvchProceedingsNo & "'"
   Rec.Open mSQL, mCnn
   Call clearVariables
   If Not (Rec.EOF And Rec.BOF) Then
    mintProceedingsID = IIf(IsNull(Rec!intProceedingsID), -1, Rec!intProceedingsID)
    mvchProceedingsNo = IIf(IsNull(Rec!vchProceedingsNo), "", Rec!vchProceedingsNo)
    mdtProceedingsDate = IIf(IsNull(Rec!dtProceedingsDate), "", Rec!dtProceedingsDate)
    mintApplicationID = IIf(IsNull(Rec!intApplicationID), -1, Rec!intApplicationID)
    mintModuleID = IIf(IsNull(Rec!intModuleID), -1, Rec!intModuleID)
    mintVoucherID = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
    mintVoucherNo = IIf(IsNull(Rec!intVoucherNo), -1, Rec!intVoucherNo)
    mvchRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
    mtnyUsed = IIf(IsNull(Rec!tnyUsed), -1, Rec!tnyUsed)
    mtnyRemoved = IIf(IsNull(Rec!tnyRemoved), -1, Rec!tnyRemoved)
    End If
       Rec.Close
       mCnn.Close
End Sub
   ' Sub for Save   '
Public Sub SaveProceedings()
   Dim objDB As New clsDB
   Dim mCnn As New ADODB.Connection
   Dim Rec As New ADODB.Recordset
   Dim mCount As Integer
   Dim mSQL As String
   Dim mArrayInput As Variant
   '       Craeting Connction              '
   If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
       MsgBox "Connection to Saankhya Not Present", vbCritical
       Exit Sub
   End If
   mArrayInput = Array( _
                       mintProceedingsID, _
                       mvchProceedingsNo, _
                       mdtProceedingsDate, _
                       mvchRemarks, _
                       mtnyUsed, _
                       mtnyRemoved, _
                       mintModuleID)
   objDB.ExecuteSP "spSaveProceedings", mArrayInput, , , mCnn
End Sub

   ' Clear Variables   '
Private Sub clearVariables()
   mintProceedingsID = -1
   mvchProceedingsNo = ""
   mdtProceedingsDate = ""
   mintApplicationID = Empty
   mintModuleID = Empty
   mintVoucherID = ""
   mintVoucherNo = ""
   mvchRemarks = ""
   mtnyUsed = -1
   mtnyRemoved = -1
End Sub

