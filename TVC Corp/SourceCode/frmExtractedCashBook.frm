VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmExtractedCashBook 
   BorderStyle     =   0  'None
   Caption         =   "Extracted Cash Book"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   Icon            =   "frmExtractedCashBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5852.161
   ScaleMode       =   0  'User
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2295
      TabIndex        =   8
      Top             =   6795
      Width           =   2085
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "UNDO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9045
      TabIndex        =   7
      Top             =   6750
      Width           =   2085
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   6795
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   13065
      Begin VB.TextBox txtClosingBalance 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8415
         TabIndex        =   5
         Top             =   405
         Width           =   2445
      End
      Begin VB.TextBox txtFinancialYear 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1575
         TabIndex        =   4
         Top             =   405
         Width           =   2445
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Balance As On"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5985
         TabIndex        =   3
         Top             =   405
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   405
         Width           =   1320
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5205
      Left            =   45
      TabIndex        =   0
      Top             =   1440
      Width           =   13050
      _cx             =   23019
      _cy             =   9181
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmExtractedCashBook.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmExtractedCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mYearID As Integer
Dim mMonthID As Integer
Dim mLoadMode As Integer ' 1-Yearly 2-Monthly
Dim dtClosingDate As Variant

Private Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    If mLoadMode = 1 Then
        dtClosingDate = "31/Mar/" + CStr(mYearID + 1)
    Else
       Call MonthLastDay(mMonthID)
    End If
 
    mSQL = " SELECT A.vchAccountHeadCode,A.vchAccountHead,A.intAccountHeadID,SUM(A.fltAmount) fltAmount FROM"
    mSQL = mSQL + "     ("
    mSQL = mSQL + "     SELECT   faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,faDailyExtracts.intAccountHeadID intAccountHeadID, fltAmount  FROM faDailyExtracts"
    mSQL = mSQL + "     INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faDailyExtracts.intAccountHeadID"
    mSQL = mSQL + "     WHERE   "
    mSQL = mSQL + "     faDailyExtracts.intFinancialYearID=" & mYearID & ""
    
    If mLoadMode = 2 Then
        mSQL = mSQL + "         AND faDailyExtracts.intmonthID < = " & mMonthID & ""
    End If
    
    mSQL = mSQL + "         AND faAccountHeads.vchAccountHeadCode LIKE '450%'"
    mSQL = mSQL + "     ) A"
    mSQL = mSQL + "     Group By"
    mSQL = mSQL + "        A.vchAccountHeadCode,"
    mSQL = mSQL + "        A.vchAccountHead,"
    mSQL = mSQL + "        A.intAccountHeadID"
    mSQL = mSQL + "     Order By"
    mSQL = mSQL + "        A.vchAccountHeadCode"
        
    Rec.CursorLocation = adUseClient
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
        vsGrid.TextMatrix(mRowCnt, 2) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    Rec.Close

End Sub

Private Sub cmdClose_Click()
    If mLoadMode = 1 Then
        frmExtractYearWiseList.FillGrid
        frmExtractYearWiseList.frmProgressBar.Visible = False
    Else
        frmExtactedMonthWiseList.FillGrid
    End If
    Unload Me
End Sub

Private Sub cmdUndo_Click()
    Dim mCnn    As New ADODB.Connection
    Dim objDB   As New clsDB
    Dim mArrIn  As Variant
    Dim mSQL    As String
    

    
    If objDB.SetConnection(mCnn) Then
        mSQL = "DELETE FROM faPostingIndex WHERE intFinYearID=" & mYearID & " AND intMonthID=" & mMonthID & " AND tnyVerifyCash=1  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        
        mSQL = ""
        mSQL = " DELETE FROM faDailyExtracts WHERE  intFinancialYearID=" & mYearID & ""
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        
        mSQL = ""
        mSQL = "UPDATE faDailyIndex SET tnyExtractFlag=NULL, tnySyncFlag=NULL WHERE intFinYearID=" & mYearID & " "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        
        cmdVerify.Enabled = True
        cmdUndo.Enabled = False
    End If
End Sub

Private Sub cmdVerify_Click()
    Dim mCnn            As New ADODB.Connection
    Dim objDB           As New clsDB
    Dim mArrIn          As Variant
    Dim mFlag           As Boolean
    Dim mCurrentDate    As Date
    Dim mSQL            As String
    Dim Rec             As New ADODB.Recordset
    
    If objDB.SetConnection(mCnn) Then
          
        mSQL = "SELECT GETDATE() CurrentDate"
        Set Rec = GetRecordSet(mSQL)
        If Not (Rec.BOF And Rec.EOF) Then
            mCurrentDate = Format(Rec!currentdate, "dd-mmm-yyyy")
            If CDate(mCurrentDate) >= CDate(dtClosingDate) Then
                mFlag = True
            Else
                mFlag = False
            End If
        End If
        Rec.Close
        
    
       If mFlag = True Then
            mArrIn = Array(-1, mYearID, IIf(mMonthID = 0, 3, mMonthID), dtClosingDate, _
                            dtClosingDate, _
                            Null, _
                            Null, _
                            0, _
                            0, _
                            1, _
                            0 _
                            )
            objDB.ExecuteSP "spSavePostingIndexCashBank", mArrIn, , , mCnn, adCmdStoredProc
            MsgBox "Verified Successfully!", vbInformation, "Saankhya"
            cmdVerify.Enabled = False
       Else
        MsgBox "Please Check the Closing Date!!!!!!", vbInformation
        Exit Sub
       End If
    End If
End Sub

Private Sub Form_Activate()
    Me.Left = 1850
    Me.Top = 2200
End Sub

Private Sub Form_Load()
    dtClosingDate = "31/Mar/" + CStr(mYearID + 1)
    txtClosingBalance.Text = dtClosingDate
    txtClosingBalance.Enabled = False
    txtFinancialYear.Text = mYearID
    txtFinancialYear.Enabled = False
    
    cmdUndo.Visible = False
    Call mCheckStatus
    Call FillGrid
End Sub

Public Property Let YearID(mData As Integer)
    mYearID = mData
End Property

Public Property Get YearID() As Integer
    YearID = mYearID
End Property

Private Sub Form_Unload(Cancel As Integer)
    If mLoadMode = 1 Then
        frmExtractYearWiseList.FillGrid
        frmExtractYearWiseList.frmProgressBar.Visible = False
    Else
        frmExtactedMonthWiseList.FillGrid
    End If
End Sub

Public Property Let MonthID(mData As Integer)
    mMonthID = mData
End Property

Public Property Get MonthID() As Integer
    MonthID = mMonthID
End Property
Public Property Let LoadMode(mData As Integer)
    mLoadMode = mData
End Property

Public Property Get LoadMode() As Integer
    LoadMode = mLoadMode
End Property
Private Function mCheckStatus() As Integer
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mStatus As Integer
    
    mStatus = -1
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSQL = " SELECT * FROM faPostingIndex WHERE   intMonthID=" & mMonthID & " AND tnyVerifyCash=1 AND intFinYearID=" & mYearID
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        mStatus = Rec!tnyStage
    End If
    If mStatus <> 2 And mStatus > -1 Then
        cmdUndo.Visible = True
    End If
    Rec.Close
    mCnn.Close
End Function
Private Function MonthLastDay(intMonthID As Integer) As Variant

  Dim dtMonthLastDate  As Variant
  
  Select Case intMonthID
        
            Case 4
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 5
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 6
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 7
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 8
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 9
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 10
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 11
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 12
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 1
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID + 1)
            Case 2
                    dtMonthLastDate = "28/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID + 1)
            Case 3
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID + 1)
            End Select
            
            dtClosingDate = dtMonthLastDate
            txtClosingBalance.Text = dtClosingDate
            
End Function
