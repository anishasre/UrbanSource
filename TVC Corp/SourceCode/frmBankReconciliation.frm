VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmBankReconciliation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Entry Form"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   Begin VSFlex8LCtl.VSFlexGrid Grid 
      Height          =   255
      Left            =   9315
      TabIndex        =   23
      Top             =   6135
      Visible         =   0   'False
      Width           =   1695
      _cx             =   2990
      _cy             =   450
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBankReconciliation.frx":0000
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
   Begin MSComctlLib.ProgressBar pbImport 
      Height          =   285
      Left            =   1410
      TabIndex        =   22
      Top             =   5790
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExtractData 
      Caption         =   "Import Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   21
      Top             =   5745
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5940
      TabIndex        =   11
      Top             =   6195
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   6195
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin VB.CheckBox chkOpening 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Opening"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   10215
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Width           =   1440
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   60
         TabIndex        =   17
         Top             =   120
         Width           =   2775
         Begin VB.CommandButton cmdDateDown2 
            Caption         =   "u"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            TabIndex        =   9
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton cmdDateUp2 
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2010
            TabIndex        =   8
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton cmdDateUp1 
            Caption         =   "t"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2010
            TabIndex        =   5
            Top             =   210
            Width           =   315
         End
         Begin VB.CommandButton cmdDateDown1 
            Caption         =   "u"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            TabIndex        =   6
            Top             =   210
            Width           =   315
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   630
            TabIndex        =   7
            Top             =   660
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            Format          =   66125825
            CurrentDate     =   39612
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   630
            TabIndex        =   4
            Top             =   210
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            Format          =   66125825
            CurrentDate     =   39612
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   150
            TabIndex        =   19
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   330
            TabIndex        =   18
            Top             =   660
            Width           =   210
         End
      End
      Begin VB.TextBox txtBank 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   15
         Top             =   840
         Width           =   4725
      End
      Begin VB.CommandButton cmdSearchBank 
         Caption         =   "---"
         Height          =   285
         Left            =   11250
         TabIndex        =   1
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtAccountHeadID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   14
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblEntryType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Entry is Selected"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   330
         Left            =   4980
         TabIndex        =   20
         Top             =   330
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4950
         TabIndex        =   16
         Top             =   870
         Width           =   450
      End
   End
   Begin VB.TextBox txtOpeningBalance 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10110
      TabIndex        =   2
      Top             =   1350
      Width           =   1635
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -210
      Top             =   6660
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid fgReconsilationEntry 
      Height          =   3960
      Left            =   0
      TabIndex        =   3
      Top             =   1740
      Width           =   11805
      _cx             =   20823
      _cy             =   6985
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBankReconciliation.frx":002A
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
      SubtotalPosition=   0
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
      Editable        =   2
      ShowComboButton =   0
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
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3570
      Top             =   7950
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8640
      TabIndex        =   13
      Top             =   1380
      Width           =   1455
   End
End
Attribute VB_Name = "frmBankReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '=================================================================================='
    '                              Bank Reconsilation Entry                            '
    '=================================================================================='
    '
    '----------------------------------------------------------------------------------'
    '       Designed & Coaded By    :       Cijith Sreedharan                          '
    '       Date of Comletion       :       15/06/08                                   '
    '       Stored Procedure Used   :       spSaveBankReconsilation                    '
    '       Gloabal Variables Used  :       mSearchID,gbSearchID,gbSearchStr           '
    '----------------------------------------------------------------------------------'

Option Explicit
Dim mSearchID As Long

Private Sub chkOpening_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkOpening.Value = 1
    End If
End Sub

Private Sub chkOpening_LostFocus()
    If txtAccountHeadID.Text <> "" Then
        If chkOpening.Value = 1 Then
            MsgBox "You Have Selected for Opening Entry", vbOKOnly
            lblEntryType.Caption = "Opening Entry is Selected"
            lblEntryType.Visible = True
            fgReconsilationEntry.Clear 1, 1
            ShowOpeningEntry
        ElseIf chkOpening.Value = 0 Then
            'MsgBox "You Have Selected for Bank Entry", vbOKOnly
            'Call ClearAll
            'lblEntryType.Caption = "Bank Entry is Selected"
            'lblEntryType.Visible = True
            'fgReconsilationEntry.Clear 1, 1
            'fgReconsilationEntry.Rows = 2
        End If
    Else
        'MsgBox "Please Select the Bank", vbInformation
        'cmdSearchBank.SetFocus
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearAll
    fgReconsilationEntry.Clear 1, 1
    fgReconsilationEntry.Rows = 2
    Grid.Clear 1, 1
    Grid.Rows = 1
End Sub

Private Sub cmdDateDown1_Click()
'    If DTPicker1.Month < 12 Then
'        DTPicker1.Month = DTPicker1.Month + 1
'    End If
    DTPicker1.Value = DTPicker1.Value + 1
End Sub

Private Sub cmdDateDown2_Click()
'    If DTPicker2.Month < 12 Then
'        DTPicker2.Month = DTPicker2.Month + 1
'    End If
    DTPicker2.Value = DTPicker2.Value + 1
End Sub

Private Sub cmdDateUp1_Click()
'    If DTPicker1.Month <> 1 Then
'        If DTPicker1.Month > 0 Then
'            DTPicker1.Month = DTPicker1.Month - 1
'        End If
'    End If
    DTPicker1.Value = DTPicker1.Value - 1
End Sub

Private Sub cmdDateUp2_Click()
'    If DTPicker2.Month <> 1 Then
'        If DTPicker2.Month > 0 Then
'            DTPicker2.Month = DTPicker2.Month - 1
'        End If
'    End If
    DTPicker2.Value = DTPicker2.Value - 1
End Sub

Private Sub cmdExtractData_Click()
    Dim mCn             As ADODB.Connection
    Dim Rec             As New ADODB.Recordset
    Dim mCount          As Long
    Dim mRowCount       As Double
    Dim mString         As Variant
    Dim mSerialNo       As Double
    Dim mChequeNo       As String
    Dim mDescription    As String
    Dim mLength         As Long
    Dim mLoopCount      As Long
    Dim mArray          As Variant
    Dim mBalance        As Variant
    Dim mLoop           As Long
    
    txtAccountHeadID = Trim(txtAccountHeadID)
    '--------------------------------------------------------------------------'
    ' IF ITS TREASURY DATA CALL READTREASURY TO IMPORT
    ' ADDED BY AIBY ON 18-JAN-2008
    '--------------------------------------------------------------------------'
    Me.MousePointer = vbHourglass
    If Left(txtAccountHeadID, 6) = "450250" Or Left(txtAccountHeadID, 6) = "450450" Or Left(txtAccountHeadID, 6) = "450650" Then
        Call ReadTreasury(txtAccountHeadID)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
    '--------------------------------------------------------------------------'
    ' IF ITS Bank Account - State Bank Of India                                '
    '--------------------------------------------------------------------------'
        Grid.LoadGrid App.Path & "\CORP.xls", flexFileExcel
        fgReconsilationEntry.Rows = 1
        For mLoop = 2 To Grid.Rows - 1
            fgReconsilationEntry.Rows = mLoop
            fgReconsilationEntry.TextMatrix(mLoop - 1, 0) = DdMmmYy(Grid.Cell(flexcpText, mLoop, 1))  ' Bank Entry Date
            fgReconsilationEntry.TextMatrix(mLoop - 1, 1) = Grid.Cell(flexcpText, mLoop, 3)  ' Particulars
            fgReconsilationEntry.TextMatrix(mLoop - 1, 2) = ""                          ' Cheque No
            
            '------------------------------------------------------------------'
            ' Note:- Extracting Cheque Number                                  '
            '------------------------------------------------------------------'
            mChequeNo = ""
            mDescription = ""
            mArray = ""
            
            mString = IIf(IsNull(Grid.Cell(flexcpText, mLoop, 3)), "", CStr(Grid.Cell(flexcpText, mLoop, 3)))
            mLength = Len(mString)
            While mLength <> 0
                mArray = mID(RTrim(mString), mLength, 1)
                If IsNumeric(mArray) Or mArray = "" Then
                    mChequeNo = mArray + mChequeNo
                Else
                    GoTo PP2
                End If
                mLength = mLength - 1
            Wend
            
PP2:        If Len(mChequeNo) > 4 Then
                fgReconsilationEntry.TextMatrix(mLoop - 1, 2) = mChequeNo
            End If
            If mChequeNo <> "" Then
                mLength = InStr(1, mString, mChequeNo, vbTextCompare)
                mDescription = mID(mString, 1, mLength - 1)
                fgReconsilationEntry.TextMatrix(mLoop - 1, 1) = mDescription
            Else
                fgReconsilationEntry.TextMatrix(mLoop - 1, 1) = mString
            End If
            fgReconsilationEntry.TextMatrix(mLoop - 1, 3) = Grid.Cell(flexcpText, mLoop, 1)  ' Cheque Date
            fgReconsilationEntry.TextMatrix(mLoop - 1, 4) = Grid.Cell(flexcpText, mLoop, 6)  ' Debit Amount
            fgReconsilationEntry.TextMatrix(mLoop - 1, 5) = Grid.Cell(flexcpText, mLoop, 7)  ' Credit Amount
            fgReconsilationEntry.TextMatrix(mLoop - 1, 6) = Grid.Cell(flexcpText, mLoop, 8)
        Next
    End If
    Me.MousePointer = vbDefault
    '--------------------------------------------------------------------------'
    
    Exit Sub
    
    
    
    
    
    Set mCn = New ADODB.Connection
    With mCn
        .Provider = "MSDASQL"
        .ConnectionString = "Driver={Microsoft Excel Driver (*.xls)}; DBQ=" & App.Path & "\CORPO.xls; ReadOnly=False;"
        .Open
    End With
    mCount = -1
    
    
    'sXL = "c:\DaniWebExample.xls"
    'Set cn = New ADODB.Connection
    'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sXL & ";Extended Properties=Excel 8.0;Persist Security Info=False"
    'cn.ConnectionTimeout = 40
    'cn.Open

    
    GoTo Skip:
    
    
    Rec.CursorLocation = adUseClient
    'Rec.Open "Select * From `Sheet1$`", mCn, adOpenDynamic, adLockOptimistic
    Rec.Open "Select * From [CORPO$A1:G90]", mCn, adOpenDynamic, adLockOptimistic
    If Not (Rec.BOF And Rec.EOF) Then
        FileInitialize
        While Not Rec.EOF
            
            Print #gbFileNO, Rec.Fields(0).Value;
            Print #gbFileNO, Rec.Fields(1).Value;
            Print #gbFileNO, PadR(IIf(IsNull(Rec.Fields(3).Value), "", Rec.Fields(3).Value), 25); "  ";
            Print #gbFileNO, IIf(IsNull(Rec.Fields(4).Value), "", Rec.Fields(4).Value);
            Print #gbFileNO, IIf(IsNull(Rec.Fields(5).Value), "", Rec.Fields(5).Value);
            Print #gbFileNO, IIf(IsNull(Rec.Fields(6).Value), "", Rec.Fields(6).Value)
            Rec.MoveNext
        Wend
        Close #gbFileNO
        ShellPad
    End If
    Rec.Close
    ' "SELECT * FROM [Sheet1$A1:B10]"
Skip:
    Set Rec = Read_Excel(App.Path & "\CORP.xls")
    If Not (Rec.BOF And Rec.EOF) Then
        FileInitialize
        mRowCount = 1
        mSerialNo = 1
        pbImport.Max = Rec.RecordCount + 1
        pbImport.Value = 0
        While Not Rec.EOF
            If IsNull(Rec.Fields(1).Value) = False Then
                mChequeNo = ""
                mDescription = ""
                mArray = ""
                mString = IIf(IsNull(Rec.Fields(2).Value), "", CStr(Rec.Fields(2).Value))
                mLength = Len(mString)
                While mLength <> 0
                    mArray = mID(RTrim(mString), mLength, 1)
                    If IsNumeric(mArray) Or mArray = "" Then
                        mChequeNo = mArray + mChequeNo
                    Else
                        GoTo PP
                    End If
                    mLength = mLength - 1
                Wend
PP:             If Len(mChequeNo) > 4 Then
                    fgReconsilationEntry.TextMatrix(mRowCount, 2) = mChequeNo
                End If
                If mChequeNo <> "" Then
                    mLength = InStr(1, mString, mChequeNo, vbTextCompare)
                    mDescription = mID(mString, 1, mLength - 1)
                    fgReconsilationEntry.TextMatrix(mRowCount, 1) = mDescription
                Else
                    fgReconsilationEntry.TextMatrix(mRowCount, 1) = mString
                End If
            End If
            Print #gbFileNO, Rec.Fields(5).Value, Rec.Fields(6).Value
            fgReconsilationEntry.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec.Fields(0).Value), "", Rec.Fields(0).Value)
            fgReconsilationEntry.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec.Fields(0).Value), "", Rec.Fields(0).Value)
            fgReconsilationEntry.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec.Fields(5).Value), "", Rec.Fields(5).Value)
            fgReconsilationEntry.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec.Fields(5).Value), Abs(Rec.Fields(7).Value - mBalance), Rec.Fields(5).Value)
            fgReconsilationEntry.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec.Fields(6).Value), "", Rec.Fields(6).Value)
            fgReconsilationEntry.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec.Fields(7).Value), "", Rec.Fields(7).Value)
            fgReconsilationEntry.TextMatrix(mRowCount, 8) = mSerialNo
            mBalance = IIf(IsNull(Rec.Fields(7).Value), 0, Rec.Fields(7).Value)
  

            fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
            mRowCount = mRowCount + 1
            mSerialNo = mSerialNo + 1
            If pbImport.Value < pbImport.Max + 1 Then
                pbImport.Value = pbImport.Value + 1
            End If
            cmdExtractData.Enabled = False
            Rec.MoveNext
        Wend
        Close #gbFileNO
        ShellPad
    End If
    Rec.Close
    cmdExtractData.Enabled = True
End Sub
Private Sub cmdSave_Click()
    Call SaveReconsilation
    Call ClearAll
    fgReconsilationEntry.Clear 1, 1
    fgReconsilationEntry.Rows = 2
End Sub

Private Sub cmdSearchBank_Click()
    Dim mSql As String
    Dim mCount As Integer
    mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
    frmSearchAccountHeads.SQLString = mSql
    frmSearchAccountHeads.Show vbModal
    mCount = InStr(1, gbSearchStr, " ")
    mSearchID = gbSearchID
    txtAccountHeadID.Text = IIf(IsNull(Left(gbSearchStr, mCount)), "", Left(gbSearchStr, mCount))
    If mCount <> 0 Then
        txtBank.Text = IIf(IsNull(mID(gbSearchStr, mCount)), "", mID(gbSearchStr, mCount))
    End If
    gbSearchStr = ""
    gbSearchID = -1
End Sub

Private Sub ClearAll()
    txtAccountHeadID.Text = ""
    txtBank.Text = ""
    txtOpeningBalance.Text = ""
    chkOpening.Value = 0
    DTPicker1.Value = Date - (Day(Date) - 1)
    DTPicker2.Value = Date
End Sub

Private Sub DTPicker2_LostFocus()
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mSql As String
    Dim mRowCount
    Dim mSql2   As String
    Dim Rec2    As New ADODB.Recordset
        
    If txtAccountHeadID.Text <> "" Then
        objdb.SetConnection mCnn
        mSql = "Set Dateformat DMY Select * from faBankReconciliationEntries where dtBankEntryDate between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "' and  vchBankAccountHeadCode = '" & txtAccountHeadID.Text & "'"
        Rec.Open mSql, mCnn
''''        If Not Rec.EOF Then
''''            mSql2 = "Select vchBankName from faBanks where intAccountHeadID = " & Rec!intBankAccountHeadID
''''            Rec2.Open mSql2, mCnn
''''            txtBank.Text = Rec2!vchBankName
''''            txtAccountHeadID.Tag = Rec!intBankAccountHeadID
''''            'txtAccountHeadID.Text = Rec!vchBankAccountHeadCode
''''        End If
        mRowCount = 0
        fgReconsilationEntry.Rows = 1
        fgReconsilationEntry.Rows = 10
        While Not Rec.EOF
            mRowCount = mRowCount + 1
            fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
            'txtAccountHeadID.Text = IIf(IsNull(Rec!vchBankAccountHeadCode), "", Rec!vchBankAccountHeadCode)
            fgReconsilationEntry.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
            fgReconsilationEntry.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
            fgReconsilationEntry.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
            fgReconsilationEntry.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!dtChequeDate), "", Rec!dtChequeDate)
            fgReconsilationEntry.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
            fgReconsilationEntry.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount)
            fgReconsilationEntry.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intReconciliationID), 0, Rec!intReconciliationID)
            Rec.MoveNext
        Wend
    Else
        MsgBox "Please Select the Bank", vbInformation
        cmdSearchBank.SetFocus
    End If
End Sub

Private Sub fgReconsilationEntry_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case 4
    If fgReconsilationEntry.TextMatrix(Row, 4) <> "" Then
        If IsNumeric(fgReconsilationEntry.TextMatrix(Row, 4)) Then
            If IsNumeric(fgReconsilationEntry.TextMatrix(Row - 1, 6)) = True Then fgReconsilationEntry.TextMatrix(Row, 6) = fgReconsilationEntry.TextMatrix(Row - 1, 6)
                fgReconsilationEntry.TextMatrix(Row, 6) = val(fgReconsilationEntry.TextMatrix(Row, 6)) - val(fgReconsilationEntry.TextMatrix(Row, 4))
                If fgReconsilationEntry.TextMatrix(Row, 6) < 0 Then fgReconsilationEntry.TextMatrix(Row, 6) = fgReconsilationEntry.TextMatrix(Row, 6) * -1
                fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
        Else
            MsgBox "Enter a Numeric Value", vbCritical
            fgReconsilationEntry.Select Row, 3
        End If
    End If
    Case 5
        If IsNumeric(fgReconsilationEntry.TextMatrix(Row - 1, 6)) = True Then fgReconsilationEntry.TextMatrix(Row, 6) = fgReconsilationEntry.TextMatrix(Row - 1, 6)
        If IsNumeric(fgReconsilationEntry.TextMatrix(Row, 5)) Then
            fgReconsilationEntry.TextMatrix(Row, 6) = val(fgReconsilationEntry.TextMatrix(Row, 6)) + val(fgReconsilationEntry.TextMatrix(Row, 5))
            fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
        Else
            MsgBox "Enter a Numeric Value", vbCritical
            fgReconsilationEntry.Select Row, 4
        End If
    End Select
End Sub

Private Sub fgReconsilationEntry_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
    If fgReconsilationEntry.TextMatrix(Row, 0) <> "" Then
        fgReconsilationEntry.TextMatrix(Row, 0) = CheckDateInMMM(fgReconsilationEntry.TextMatrix(Row, 0))
    End If
    End If
    
    If Col = 3 Then
    If fgReconsilationEntry.TextMatrix(Row, 3) <> "" Then
        fgReconsilationEntry.TextMatrix(Row, 3) = CheckDateInMMM(fgReconsilationEntry.TextMatrix(Row, 3))
    End If
    End If
End Sub


Private Sub fgReconsilationEntry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case fgReconsilationEntry.Col
            Case 0
                fgReconsilationEntry.Select fgReconsilationEntry.Row, 1, , 1
                fgReconsilationEntry.EditCell
            Case 1
                fgReconsilationEntry.Select fgReconsilationEntry.Row, 2, , 2
                fgReconsilationEntry.EditCell
            Case 2
                fgReconsilationEntry.Select fgReconsilationEntry.Row, 3, , 3
                fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row, 3) = fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row, 0)
                fgReconsilationEntry.EditCell
            Case 3
                fgReconsilationEntry.Select fgReconsilationEntry.Row, 4, , 4
                fgReconsilationEntry.EditCell
            Case 4
                If fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row, 4) <> "" Then
                    'fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
                    fgReconsilationEntry.Select fgReconsilationEntry.Row + 1, 0, , 0
                    fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row, 0) = fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row - 1, 0)
                    fgReconsilationEntry.EditCell
                Else
                    'fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
                    fgReconsilationEntry.Select fgReconsilationEntry.Row, 5, , 5
                    fgReconsilationEntry.EditCell
                End If
            Case 5
                'fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
                fgReconsilationEntry.Select fgReconsilationEntry.Row + 1, 0, , 0
                fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row, 0) = fgReconsilationEntry.TextMatrix(fgReconsilationEntry.Row - 1, 0)
                fgReconsilationEntry.EditCell
            
        End Select
    End If
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitIDESubClassing
    ClearAll
    lblEntryType.Visible = False
End Sub

Private Sub lstBank_DblClick()
'        Dim mVarA As Variant
'        Dim mVarB As Variant
'        Dim mCode As Long
'
'        Dim objBank As New clsBank
'        Dim objAcc As New clsAccounts
'        If lstBank.ListIndex > 0 Then
'            txtBank.Text = lstBank.List(lstBank.ListIndex)
'            txtBank.Tag = lstBank.ItemData(lstBank.ListIndex)
'            objBank.SetBankInfo (lstBank.ItemData(lstBank.ListIndex))
'            mVarA = objBank.Branch
'            mVarB = objBank.BranchCode
'            objAcc.SetAccounts (objBank.BankAccountHeadID)
'            mCode = objAcc.AccountCode
'            txtAccountHeadID.Text = mCode
'        End If
'        lstBank.Visible = False
End Sub

Private Sub SaveReconsilation()
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim objdb As New clsDB
    Dim varin As Variant
    Dim mAmt As Variant
    Dim i As Integer
    Dim mMaxReconsilationID As Long
        objdb.SetConnection mCnn
    ReDim varin(9) As Variant
    For i = 1 To fgReconsilationEntry.Rows - 1
        varin(0) = IIf(IsNull(fgReconsilationEntry.TextMatrix(i, 7)), "", fgReconsilationEntry.TextMatrix(i, 7))
        'varin(1) = IIf(Val(txtAccountHeadID.Tag) = 0, mSearchID, txtAccountHeadID.Tag)
        varin(1) = mSearchID
        varin(2) = txtAccountHeadID.Text
        varin(3) = fgReconsilationEntry.TextMatrix(i, 0)
        varin(4) = Trim(fgReconsilationEntry.TextMatrix(i, 1))
        varin(5) = Trim(fgReconsilationEntry.TextMatrix(i, 2))
        varin(6) = fgReconsilationEntry.TextMatrix(i, 3)
        
        mAmt = fgReconsilationEntry.TextMatrix(i, 4)
        If IsNumeric(mAmt) Then
            If mAmt > 0 Then
                varin(7) = mAmt
            Else
                varin(7) = Null
            End If
        Else
            varin(7) = Null
        End If
        
        mAmt = Null
        mAmt = fgReconsilationEntry.TextMatrix(i, 5)
        If IsNumeric(mAmt) Then
            If mAmt > 0 Then
                varin(8) = mAmt
            Else
                varin(8) = Null
            End If
        Else
            varin(8) = Null
        End If
        
        varin(9) = chkOpening.Value
        If fgReconsilationEntry.TextMatrix(i, 1) <> "" Then
            objdb.ExecuteSP "spSaveBankReconsilation", varin, , , mCnn
        End If
    Next
End Sub

Private Sub txtOpeningBalance_LostFocus()
    Dim mRowCount As Integer
    Dim mCurrentOpening As Double
    If IsNumeric(txtOpeningBalance.Text) Then
        fgReconsilationEntry.TextMatrix(1, 6) = txtOpeningBalance.Text
        mCurrentOpening = val(txtOpeningBalance.Text)
        If fgReconsilationEntry.Rows > 2 Then
            For mRowCount = 1 To fgReconsilationEntry.Rows - 2
                mCurrentOpening = IIf(val(fgReconsilationEntry.TextMatrix(mRowCount, 4)) = 0, val(fgReconsilationEntry.TextMatrix(mRowCount, 5)) + val(mCurrentOpening), val(mCurrentOpening) - val(fgReconsilationEntry.TextMatrix(mRowCount, 4)))
                fgReconsilationEntry.TextMatrix(mRowCount, 6) = mCurrentOpening
            Next
        End If
    Else
        txtOpeningBalance.Text = "0"
        'MsgBox "Enter a Numeric Value", vbCritical
        'txtOpeningBalance.SetFocus
    End If
End Sub

Private Sub ShowOpeningEntry()
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mSql As String
    Dim mRowCount
    Dim mSql2   As String
    Dim Rec2    As New ADODB.Recordset
    objdb.SetConnection mCnn
    mSql = "Select * from faBankReconciliationEntries where tnyOpening = 1 and vchBankAccountHeadCode = '" & txtAccountHeadID.Text & "'"
    Rec.Open mSql, mCnn
'''    If Not Rec.EOF Then
'''        mSql2 = "Select vchBankName from faBanks where intAccountHeadID = " & Rec!intBankAccountHeadID
'''        Rec2.Open mSql2, mCnn
'''        txtBank.Text = Rec2!vchBankName
'''        txtAccountHeadID.Tag = Rec!intBankAccountHeadID
'''    End If
    While Not Rec.EOF
        mRowCount = mRowCount + 1
        fgReconsilationEntry.Rows = fgReconsilationEntry.Rows + 1
        'txtAccountHeadID.Text = Rec!vchBankAccountHeadCode
        fgReconsilationEntry.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
        fgReconsilationEntry.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
        fgReconsilationEntry.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
        fgReconsilationEntry.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!dtChequeDate), "", Rec!dtChequeDate)
        fgReconsilationEntry.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
        fgReconsilationEntry.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount)
        fgReconsilationEntry.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intReconciliationID), 0, Rec!intReconciliationID)
        Rec.MoveNext
    Wend
End Sub
 Public Function Read_Excel(ByVal sFile As String) As ADODB.Recordset
    
          On Error GoTo fix_err
          Dim Rec As ADODB.Recordset
          Set Rec = New ADODB.Recordset
          Dim mCn As String
    
          Rec.CursorLocation = adUseClient
          Rec.CursorType = adOpenKeyset
          Rec.LockType = adLockBatchOptimistic
    
          mCn = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & sFile
          'Rec.Open "SELECT * FROM [CORPO$]", mCn
          Rec.Open "SELECT * FROM [Sheet1$]", mCn
          Set Read_Excel = Rec
          Set Rec = Nothing
          Exit Function
fix_err:
          Debug.Print err.Description + " " + _
                      err.Source, vbCritical, "Import"
          err.Clear
    End Function
    Private Sub ReadTreasury(mAccountHeadCode As String)
        Dim mFileNo As Integer
        Dim mFileName As String
        Dim objLines As New Collection
        Dim mLine As String
        Dim mSkipTopRows As Integer
        Dim mLoop As Long
        Dim mRowForGrid As String
        Dim mSerialNo As Long
        
        ' Obtain the next free file descriptor.
        mFileNo = FreeFile
        
        ' Make sure a file is specified.
        If mAccountHeadCode = "" Then
            MsgBox "Please specify the Treasury Head Code..!", vbInformation
            Exit Sub
        End If
        
        mFileName = mAccountHeadCode & ".Txt"
        ' Make sure the file exists before trying to open it.
        If Dir(mFileName) = "" Then
            MsgBox "File not found in the Directory...", vbInformation
            Exit Sub
        End If
        
        ' Read the collection from a text file.
        Open mFileName For Input As mFileNo
        For mSkipTopRows = 1 To 9
            Line Input #mFileNo, mLine
        Next mSkipTopRows
        While Not EOF(mFileNo)
            Line Input #mFileNo, mLine
            objLines.Add mLine
        Wend
        Close mFileNo
        fgReconsilationEntry.Rows = 1
        mSerialNo = 1
         For mLoop = 1 To (objLines.count)
            mLine = Trim(objLines.Item(mLoop))
            If IsDate(mID(mLine, 1, 10)) Then
                mRowForGrid = Trim(mID(mLine, 1, 10)) & vbTab & Trim(mID(mLine, 35, 10)) & vbTab & Trim(mID(mLine, 23, 7)) & vbTab & Trim(mID(mLine, 12, 10)) & vbTab & Trim(mID(mLine, 48, 16)) & vbTab & Trim(mID(mLine, 65, 16)) & vbTab & Trim(mID(mLine, 82, 16)) & vbTab & vbTab & mSerialNo
                
                fgReconsilationEntry.AddItem mRowForGrid
                mSerialNo = mSerialNo + 1
            End If
        Next mLoop
        Set objLines = Nothing
    End Sub
