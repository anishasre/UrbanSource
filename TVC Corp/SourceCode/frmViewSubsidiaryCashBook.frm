VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmViewSubsidiaryCashBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subsidiary Cash Book Transactions"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7260
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   6180
      Width           =   2490
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4380
      Left            =   165
      TabIndex        =   2
      Top             =   1770
      Width           =   9585
      _cx             =   16907
      _cy             =   7726
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmViewSubsidiaryCashBook.frx":0000
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9855
      TabIndex        =   1
      Top             =   6540
      Width           =   9915
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4290
         TabIndex        =   4
         Top             =   60
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   720
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   0
      Width           =   9915
   End
   Begin VB.Image imgChild 
      Height          =   240
      Left            =   1275
      Picture         =   "frmViewSubsidiaryCashBook.frx":008F
      Top             =   1740
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgReport 
      Height          =   240
      Left            =   405
      Picture         =   "frmViewSubsidiaryCashBook.frx":01D9
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFolder 
      Height          =   480
      Left            =   645
      Picture         =   "frmViewSubsidiaryCashBook.frx":061B
      Top             =   1650
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6525
      TabIndex        =   3
      Top             =   6210
      Width           =   630
   End
End
Attribute VB_Name = "frmViewSubsidiaryCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    '*********************************************************************************************'
    '              Form to view the Transaction Details of a Subsidiary Cash Book                 '
    '*********************************************************************************************'
    Private Sub FillGrid()
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim RecParent       As New ADODB.Recordset
        Dim RecChild        As New ADODB.Recordset
        Dim mSQL            As String
        Dim mSQLParent      As String
        Dim mSQLChild       As String
        Dim objDB           As New clsDB
        Dim mAmount         As Variant
        Dim mCashBalance    As Variant
        Dim mPaymentID      As Variant
        Dim mRemitanceID    As Variant
        
        On Error GoTo err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mAmount = ""
        txtTotal.Text = "0.00"
        mCashBalance = ""
    '    mSQL = "Select Sum(fltAmount) As Amount From faSubsidiaryCashBook"
    '    mSQL = mSQL + " Left Join faSubsidiaryAccountHeads On faSubsidiaryCashBook.intCashBookID = faSubsidiaryAccountHeads.intSubsidiaryAccountHeadID"
    '    mSQL = mSQL + " Where intTypeID = 50"
    '    mSQL = mSQL + " And tnyStatus = 2"
    '    Rec.Open mSQL, mCnn
    '    If Not (Rec.EOF And Rec.BOF) Then
    '        mAmount = IIf(IsNull(Rec!Amount), "", Rec!Amount)
    '    End If
    '    Rec.Close
    '
        vsGrid.Rows = 1
        vsGrid.OutlineBar = flexOutlineBarCompleteLeaf
    '    vsGrid.AddItem "Subsidiary Cash Book"
    '    vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
    '    vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 0
    '    vsGrid.Cell(flexcpPicture, vsGrid.Rows - 1, 0) = imgFolder
    '    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HFFFFFF
    '    vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = mAmount
        
        mSQL = "Select * From faSubsidiaryCashBook"
        mSQL = mSQL + " Left Join faSubsidiaryAccountHeads On faSubsidiaryCashBook.intCashBookID = faSubsidiaryAccountHeads.intSubsidiaryAccountHeadID"
        mSQL = mSQL + " Where intTypeID = 50"
        'mSQL = mSQL + " And tnyStatus = 2"
        Rec.Open mSQL, mCnn
        While Not Rec.EOF
            vsGrid.AddItem IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
            'vsGrid.IsCollapsed(vsGrid.Rows - 1) = flexOutlineCollapsed
            vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
            vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 0
            vsGrid.Cell(flexcpPicture, vsGrid.Rows - 1, 0) = imgFolder
            vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HFFFFFF
            vsGrid.Cell(flexcpFontBold, vsGrid.Rows - 1, 1) = True
            vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!fltAmount), "", Format(Rec!fltAmount, "0.00"))
            vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!intID), "", Rec!intID)
            vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!intCashBookID), "", Rec!intCashBookID)
            
            vsGrid.AddItem "Transferred Amount"
            vsGrid.IsCollapsed(vsGrid.Rows - 1) = flexOutlineCollapsed
            vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
            vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 1
            vsGrid.Cell(flexcpPicture, vsGrid.Rows - 1, 0) = imgReport
            vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HE0E0E0
            vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!fltAmount), "", Format(Rec!fltAmount, "0.00"))
            vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!intID), "", Rec!intID)
            vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!intCashBookID), "", Rec!intCashBookID)
            
            txtTotal.Text = val(txtTotal.Text) + IIf(IsNull(Rec!fltAmount), 0, val(Rec!fltAmount))
            txtTotal.Text = Format(txtTotal.Text, "0.00")
            mSQLParent = "Select * From faSubsidiaryCashBook"
            mSQLParent = mSQLParent + " Where intCashBookID =" & CInt(vsGrid.TextMatrix(vsGrid.Rows - 1, 3))
            mSQLParent = mSQLParent + " And numUserID = " & IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
            'mSQLParent = mSQLChild + " And intTypeID = 20"
            mSQLParent = mSQLParent + " And intID > " & IIf(IsNull(Rec!intID), "", Rec!intID)
            mSQLParent = mSQLParent + " And numSeatID = " & IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
            RecParent.Open mSQLParent, mCnn
            While Not RecParent.EOF
                'mPaymentID = ""
                If mPaymentID <> "" Then
                    If Not (IsNull(RecParent!intTypeID)) Then
                        If RecParent!intTypeID = 50 Then
                            GoTo LB
                        End If
                    End If
                End If
                If Not (IsNull(RecParent!intTypeID)) Then
                    If RecParent!intTypeID = 20 Then
                        mPaymentID = IIf(IsNull(RecParent!intID), "", RecParent!intID)
                        vsGrid.AddItem "Paid Amount"
                        vsGrid.IsCollapsed(vsGrid.Rows - 1) = flexOutlineCollapsed
                        vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
                        vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 1
                        vsGrid.Cell(flexcpPicture, vsGrid.Rows - 1, 0) = imgReport
                        vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HE0E0E0
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(RecParent!fltAmount), "", Format(RecParent!fltAmount, "0.00"))
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(RecParent!intID), "", RecParent!intID)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(RecParent!intCashBookID), "", RecParent!intCashBookID)
                        
                        If mPaymentID <> "" Then
                            mSQLChild = "Select * From faSubsidiaryCashBookChild"
                            mSQLChild = mSQLChild + " Where intID = " & mPaymentID
                            RecChild.Open mSQLChild, mCnn
                            While Not RecChild.EOF
                                vsGrid.AddItem IIf(IsNull(RecChild!vchPayee), "", RecChild!vchPayee)
                                vsGrid.IsCollapsed(vsGrid.Rows - 1) = flexOutlineCollapsed
                                vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
                                vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 2
                                vsGrid.Cell(flexcpPicture, vsGrid.Rows - 1, 0) = imgChild
                                vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HC0C0C0
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(RecChild!fltAmount), "", Format(RecChild!fltAmount, "0.00"))
                                RecChild.MoveNext
                            Wend
                            RecChild.Close
                        End If
                    End If
                    If RecParent!intTypeID = 10 Then
                        vsGrid.AddItem "Balance Amount"
                        vsGrid.IsCollapsed(vsGrid.Rows - 1) = flexOutlineCollapsed
                        vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
                        vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 1
                        vsGrid.Cell(flexcpPicture, vsGrid.Rows - 1, 0) = imgReport
                        vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 1) = &HE0E0E0
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(RecParent!fltAmount), "", Format(RecParent!fltAmount, "0.00"))
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(RecParent!intID), "", RecParent!intID)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(RecParent!intCashBookID), "", RecParent!intCashBookID)
                    End If
                End If
                RecParent.MoveNext
            Wend
LB:         RecParent.Close
            Rec.MoveNext
        Wend
        Rec.Close
        mSQL = "Select Sum(fltAmount) As CashBalance From faVouchers"
        mSQL = mSQL + " Where intInstrumentTypeID = 1"
        mSQL = mSQL + " And tnyCancelFlag <> 1"
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mCashBalance = IIf(IsNull(Rec!CashBalance), "", Rec!CashBalance)
        End If
        Rec.Close
        vsGrid.AddItem "Cash Balance"
        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = Format(mCashBalance - val(txtTotal.Text), "0.00")
        vsGrid.IsSubtotal(vsGrid.Rows - 1) = True
        vsGrid.RowOutlineLevel(vsGrid.Rows - 1) = 0
        vsGrid.Cell(flexcpFontBold, vsGrid.Rows - 1, 1) = True
        txtTotal.Text = Format(mCashBalance, "0.00")
        Exit Sub
err:
        MsgBox err.Description
    End Sub
   
   Private Sub cmdClose_Click()
        Unload Me
'        Call FillGrid
    End Sub

    Private Sub Form_Load()
        vsGrid.SelectionMode = flexSelectionByRow
        Call FillGrid
    End Sub
