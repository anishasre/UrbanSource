VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmGoDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GO Details"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   Icon            =   "frmGoDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   330
      Width           =   12150
      _cx             =   21431
      _cy             =   7805
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
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
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGoDetails.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label Label5 
      Caption         =   "Canelled"
      Height          =   195
      Left            =   8010
      TabIndex        =   5
      Top             =   4815
      Width           =   690
   End
   Begin VB.Label Label4 
      Caption         =   "Used"
      Height          =   195
      Left            =   9090
      TabIndex        =   4
      Top             =   4815
      Width           =   465
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   195
      Left            =   7650
      TabIndex        =   3
      Top             =   4815
      Width           =   330
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Height          =   195
      Left            =   8775
      TabIndex        =   2
      Top             =   4815
      Width           =   285
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Go For Funds"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   9825
   End
End
Attribute VB_Name = "frmGoDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim intSourceFundID As Integer
    Dim intSuFundTrID   As Integer
    Private Sub Form_Load()
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mCnt    As Integer
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = " Select dtGODate,suGoForFunds.vchRefNo GoNo,suGoForFunds.vchDescription, vchSourceFundName,suGoForFunds.fltAmount Amount,fltAmountUpto," & vbNewLine
            mSql = mSql + " vchpayOrderNo,faVouchers.intVoucherNo VrNo,intRefId,suGoForFunds.intSourceOfFundID FundID," & vbNewLine
            mSql = mSql + " suGoForFunds.intPayOrderID PoId,suGoForFunds.intVoucherID VrID," & vbNewLine
            mSql = mSql + " Case IsNull(faPayOrder.tnyStatus, 0) "
            mSql = mSql + " when 0 then 0 "
            mSql = mSql + " when 1 then case isNull(faPayOrder.tnyCancelled,0) when 0 then 1 When 1 then 4 End"
            mSql = mSql + " end  status "
            mSql = mSql + " From suGoForFunds" & vbNewLine
            mSql = mSql + " Inner Join suSourceOfFund On suGoForFunds.intSourceOfFundID=suSourceOfFund.intSourceFundID" & vbNewLine
            mSql = mSql + " Left Join faPayOrder On faPayOrder.intPayOrderID=suGoForFunds.intPayOrderID" & vbNewLine
            mSql = mSql + " Left Join faVouchers On faVouchers.intVoucherID=suGoForFunds.intVoucherID" & vbNewLine

            Rec.CursorLocation = adUseClient
            Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
            vsGrid.Rows = 1
            If Not (Rec.BOF And Rec.EOF) Then
                vsGrid.Rows = Rec.RecordCount + 1
                vsGrid.Col = 0
                vsGrid.Row = 1
                vsGrid.ColSel = 11
                vsGrid.RowSel = vsGrid.Rows - 1
                mSql = Rec.GetString(, , vbTab, Chr(13))
                vsGrid.Clip = mSql
            End If
            Rec.Close
        End If
        
        For mCnt = 1 To vsGrid.Rows - 1
            If val(vsGrid.TextMatrix(mCnt, 11)) = 1 Then
                vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 11) = vbGreen
            ElseIf val(vsGrid.TextMatrix(mCnt, 11)) = 1 Then
                vsGrid.Cell(flexcpBackColor, mCnt, 0, vsGrid.Row, 11) = vbRed
            End If
        Next
        
    End Sub
    Private Sub FillGrid()
        'Dim f
    End Sub
    Private Sub vsGrid_DblClick()
        Call vsGrid_KeyDown(13, 0)
    End Sub

    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo Err:
            If KeyCode = vbKeyEscape Then
                Unload Me
            ElseIf KeyCode = 13 Then
                If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
                    If intSuFundTrID = gbTransactionTypeUnUtilizedAmount Then
                        If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = intSourceFundID Then
                                If val(vsGrid.TextMatrix(vsGrid.Row, 12)) <> 1 Then
                                    If val(vsGrid.TextMatrix(vsGrid.Row, 10)) < 1 Then
                                        gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 1)
                                        gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 2)
                                        gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                                        Unload Me
                                    Else
    '                                    If MsgBox("Do you want to Replace Pay Order for the GO?", vbYesNo, "Saankhya") = vbYes Then
    '
    '                                    End If
                                    End If
                                Else
                                    MsgBox "Selected Go Already Used ..."
                                    Exit Sub
                                End If
                            
                        Else
                            MsgBox "Selected Fund not Matching..."
                            Exit Sub
                        End If
                    Else
                        If val(vsGrid.TextMatrix(vsGrid.Row, 12)) <> 1 Then
                            If val(vsGrid.TextMatrix(vsGrid.Row, 10)) < 1 Then
                                gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 1)
                                gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 2)
                                gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 8)
                                Unload Me
                            Else
    '                                    If MsgBox("Do you want to Replace Pay Order for the GO?", vbYesNo, "Saankhya") = vbYes Then
    '
    '                                    End If
                            End If
                        Else
                            MsgBox "Selected Go Already Used ..."
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Public Property Get SuFund() As Integer
        SuFund = intSourceFundID
    End Property

    Public Property Let SuFund(Data As Integer)
        intSourceFundID = Data
    End Property
    Public Property Get SuFundTr() As Integer
        SuFundTr = intSourceFundID
    End Property

    Public Property Let SuFundTr(Data As Integer)
        intSuFundTrID = Data
    End Property


