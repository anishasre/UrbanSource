VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAlterTransactionType 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlterTransactionType.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Voucher Group"
      ForeColor       =   &H00000080&
      Height          =   885
      Left            =   8520
      TabIndex        =   11
      Top             =   4770
      Visible         =   0   'False
      Width           =   3075
      Begin VB.Label lblVoucherGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Group"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   900
         TabIndex        =   12
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Set To &Default"
      Height          =   405
      Left            =   150
      TabIndex        =   10
      Top             =   6150
      Width           =   1455
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3570
      Top             =   6420
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   10140
      TabIndex        =   9
      Top             =   6150
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   8640
      TabIndex        =   8
      Top             =   6150
      Width           =   1455
   End
   Begin VB.ComboBox cmbSection 
      Height          =   390
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5070
      Width           =   2955
   End
   Begin VB.TextBox txtBankHeadCode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1530
      MaxLength       =   100
      TabIndex        =   4
      Top             =   5505
      Width           =   4815
   End
   Begin VB.CommandButton cmdSearchBank 
      Caption         =   "..."
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
      Left            =   6390
      TabIndex        =   3
      Top             =   5520
      Width           =   315
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4185
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   11535
      _cx             =   20346
      _cy             =   7382
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
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   13
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAlterTransactionType.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   11655
      Y1              =   5985
      Y2              =   6000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to Change the Section Or Bank Account ?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   60
      TabIndex        =   7
      Top             =   4710
      Width           =   4695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Head Code"
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   270
      Left            =   60
      TabIndex        =   2
      Top             =   5100
      Width           =   630
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Change Section / Bank Account For Transaction Type"
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
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11505
   End
End
Attribute VB_Name = "frmAlterTransactionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Function FillGrid()
        On Error GoTo Err:
            Dim objDb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim mRowCnt As Integer
            
            If objDb.SetConnection(mCnn) Then
                mSql = "Select * from faTransactionType "
                mSql = mSql + " Inner Join faSection On faTransactionType.intSectionID = faSection.intSectionID "
                mSql = mSql + " Left Join faAccountHeads On faTransactionType.vchBankHeadCode = faAccountHeads.vchAccountHeadCode "
                mSql = mSql + " Order By vchTransactionType "
                Rec.Open mSql, mCnn
                vsGrid.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intSectionID), "", Rec!intSectionID)
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchGroup), "", Rec!vchGroup)
                    mRowCnt = mRowCnt + 1
                    vsGrid.Rows = vsGrid.Rows + 1
                    Rec.MoveNext
                Wend
            Else
                MsgBox "Connection to Finance does not exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdDefault_Click()
        If SetToDefault = True Then
            MsgBox "Restore Process Completed Successfully", vbInformation
            Call FillGrid
        End If
    End Sub

    Private Sub cmdSave_Click()
        If AlterValidations = True Then
            If AlterTrType = True Then
                MsgBox "Successfully Altered the Transaction Type", vbInformation
                Call FillGrid
                Call FormInitialize
            End If
        End If
    End Sub

    Private Sub cmdSearchBank_Click()
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = 2"
        frmSearchAccountHeads.Show 1
        If gbSearchID <> -1 Then
            txtBankHeadCode.Text = CStr(gbSearchStr)
            txtBankHeadCode.Tag = Left(gbSearchStr, 9)
        End If
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        WindowsXPC1.InitIDESubClassing
    End Sub

    Private Sub Form_Load()
        Call FormInitialize
        Call FillCombo
        Call FillGrid
    End Sub
    
    Private Sub FormInitialize()
        txtBankHeadCode.Text = ""
        txtBankHeadCode.Tag = ""
        cmbSection.ListIndex = -1
        Frame1.Visible = False
        vsGrid.Cell(flexcpBackColor, 1, 0, vsGrid.Rows - 1, 4) = vbWhite
    End Sub

    Private Sub vsGrid_Click()
        Frame1.Visible = False
        On Error GoTo Err:
            vsGrid.Cell(flexcpBackColor, 1, 0, vsGrid.Rows - 1, 4) = vbWhite
            vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 4) = &HC0C0FF
            cmbSection.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            txtBankHeadCode.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            txtBankHeadCode.Tag = vsGrid.TextMatrix(vsGrid.Row, 5)
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub FillCombo()
        On Error GoTo Err:
            PopulateList cmbSection, "Select vchSectionName,intSectionID from faSection Order By vchSectionName", , True, , True
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Function AlterValidations() As Boolean
        On Error GoTo Err:
            Dim mRowCnt As Integer
            Dim mCheckRowSelect As Boolean
            For mRowCnt = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 4) = &HC0C0FF Then
                    mCheckRowSelect = True
                End If
            Next
            If mCheckRowSelect = False Then
                AlterValidations = False
                MsgBox "Please Select the Transaction Type from the Grid", vbInformation
                vsGrid.SetFocus
                Exit Function
            End If
            
            If txtBankHeadCode.Text = "" And cmbSection.ListIndex = -1 Then
                MsgBox "Please Select the Section / Bank Account if you want to Alter.", vbInformation
                AlterValidations = False
                Exit Function
            End If
            
            AlterValidations = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function AlterTrType() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim objDb As New clsDB
            Dim mSql As String
            
            If objDb.SetConnection(mCnn) Then
                If cmbSection.ListIndex <> -1 And txtBankHeadCode.Text <> "" Then
                    mSql = "Update faTransactionType Set intSectionID = " & cmbSection.ItemData(cmbSection.ListIndex) & " , vchBankHeadCode = '" & txtBankHeadCode.Tag & "' Where intTransactionTypeID = " & vsGrid.TextMatrix(vsGrid.Row, 0)
                ElseIf cmbSection.ListIndex <> -1 Then
                    mSql = "Update faTransactionType Set intSectionID = " & cmbSection.ItemData(cmbSection.ListIndex) & " Where intTransactionTypeID = " & vsGrid.TextMatrix(vsGrid.Row, 0)
                ElseIf txtBankHeadCode.Text <> "" Then
                    mSql = "Update faTransactionType Set vchBankHeadCode = '" & Val(txtBankHeadCode.Tag) & "' Where intTransactionTypeID = " & vsGrid.TextMatrix(vsGrid.Row, 0)
                End If
                mCnn.Execute mSql
            Else
                MsgBox "Connection to Finance does not exist, Please Contact your System Administrator", vbInformation
            End If
            AlterTrType = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function SetToDefault() As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim mSql As String
            Dim objDb As New clsDB
            
            Dim mRowCnt As Integer
            Dim mCheckRowSelect As Boolean
            For mRowCnt = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpBackColor, mRowCnt, 0, mRowCnt, 4) = &HC0C0FF Then
                    mCheckRowSelect = True
                End If
            Next
            If mCheckRowSelect = False Then
                MsgBox "Please Select the Transaction Type from the Grid", vbInformation
                vsGrid.SetFocus
                SetToDefault = False
                Exit Function
            End If
            
            If objDb.SetConnection(mCnn) Then
                mSql = " Update faTransactionType Set intSectionID = intDefaultSectionID, vchBankHeadCode = vchDefaultBankHeadCode "
                mSql = mSql + " From faTransactionType Where intTransactionTypeID = " & vsGrid.TextMatrix(vsGrid.Row, 0)
                mCnn.Execute mSql
            Else
                MsgBox "Connection to Finance does not exist, Please Contact your System Administrator", vbInformation
            End If
            SetToDefault = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function

    Private Sub vsGrid_DblClick()
        Frame1.Visible = True
        Select Case (UCase(vsGrid.TextMatrix(vsGrid.Row, 6)))
            Case "R":
                lblVoucherGroup.Caption = "Receipt Voucher"
            Case "P":
                lblVoucherGroup.Caption = "Payment Voucher"
        End Select
        
    End Sub
