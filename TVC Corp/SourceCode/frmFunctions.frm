VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmFunctions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Functions"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.ListBox lstMajorFunctions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   4125
      Left            =   5910
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3210
      TabIndex        =   9
      Top             =   5880
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2955
      Left            =   420
      TabIndex        =   0
      Top             =   2730
      Width           =   8085
      _cx             =   14261
      _cy             =   5212
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   13
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFunctions.frx":0000
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
      ComboSearch     =   0
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
   Begin VB.Frame fraFunctions 
      Height          =   1995
      Left            =   0
      TabIndex        =   11
      Top             =   750
      Width           =   8925
      Begin VB.TextBox txtFunctionCode 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4230
         MaxLength       =   6
         TabIndex        =   2
         Top             =   510
         Width           =   1635
      End
      Begin VB.TextBox txtMajorFunctionCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   1560
         Width           =   2445
      End
      Begin VB.CommandButton cmdMajorFunction 
         Caption         =   "..."
         Height          =   285
         Left            =   7230
         TabIndex        =   5
         Top             =   1200
         Width           =   315
      End
      Begin VB.TextBox txtMajourFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   1200
         Width           =   3885
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   840
         Width           =   3885
      End
      Begin VB.TextBox txtPrefix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Major FunctionCode"
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   1590
         Width           =   1425
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Function"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   240
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   9630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Major Function"
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   1230
         Width           =   1050
      End
      Begin VB.Label lblFunctionName 
         AutoSize        =   -1  'True
         Caption         =   "&Function"
         Height          =   195
         Left            =   2610
         TabIndex        =   14
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblFunctionCode 
         AutoSize        =   -1  'True
         Caption         =   "Function &Code"
         Height          =   195
         Left            =   2190
         TabIndex        =   13
         Top             =   510
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8940
      TabIndex        =   17
      Top             =   0
      Width           =   8940
   End
End
Attribute VB_Name = "frmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mEditFlag As Boolean
    Dim mNewFlag As Boolean
        
    Private Sub DeleteSubFucntions(mMajorFunctionID As Long)
        Dim mLoopCount As Long
STEP1:
        For mLoopCount = 1 To vsGrid.Rows - 1
            If Val(vsGrid.TextMatrix(mLoopCount, 3)) = mMajorFunctionID And _
                Val(vsGrid.TextMatrix(mLoopCount, 2)) > 0 Then
                vsGrid.RemoveItem mLoopCount
                GoTo STEP1:
            End If
        Next mLoopCount
    End Sub
    Private Sub DisplaySubfunctions(mMajorFunctionID As Long, mRow As Long)
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        mSQL = " SELECT * FROM faFunctions WHERE intMajorFunctionID = " & mMajorFunctionID & " ORDER BY vchFunction"
        Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                    'vsGrid.AddItem ""
                    mRow = mRow + 1
                    vsGrid.AddItem "         " & Rec!vchFunction & vbTab & Rec!vchFunctionCode & vbTab & Rec!intFunctionID & vbTab & Rec!intMajorFunctionID & vbTab & "False", mRow
                    Rec.MoveNext
                Wend
            End If
        Rec.Close
    End Sub
        
    Private Sub cmdCancel_Click()
        If mEditFlag Then
            Call dispMajorFunctions
            Call FormInitialize
        Else
            Unload Me
        End If
    End Sub

    Private Sub cmdMajorFunction_Click()
        Dim mSQL As String
        mSQL = "Select vchMajorFunction, intMajorFunctionID From faMajorFunctions Order By vchMajorFunction"
        Call PopulateList(lstMajorFunctions, mSQL, , True, , True)
        lstMajorFunctions.Visible = True
        lstMajorFunctions.SetFocus
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitializeforNew
        mEditFlag = False
        mNewFlag = True
        txtFunctionCode.SetFocus
    End Sub

    Private Sub cmdSave_Click()
        Dim mintMjrFunctionID As Long
        Dim mintFunctionID As Long
        Dim mCon As New ADODB.Connection
        Dim arrInput As Variant
        
        Dim objDb As New clsDB
        
        '------------------------------------------
        '   Validations
        '------------------------------------------
        
        If Trim(txtFunctionCode.Text) = "" Then
            txtFunctionCode.SetFocus
            Exit Sub
        End If
        If Trim(txtFunction.Text) = "" Then
            txtFunction.SetFocus
            Exit Sub
        End If
        If mEditFlag And Val(txtFunctionCode.Tag) < 1 Then
            MsgBox "Error: Try again!", vbInformation
            Exit Sub
        ElseIf mEditFlag And Val(txtFunctionCode.Tag) > 0 Then
            MsgBox "Updating an Existing Function"
        ElseIf mEditFlag = False And Val(txtFunctionCode.Tag) = 0 Then
            'MsgBox "Creating a new Function"
        End If
      
        '------------------------------------------
        '   Saving a New Function
        '------------------------------------------

        objDb.SetConnection mCon
        mintMjrFunctionID = IIf(Val(txtMajorFunctionCode.Tag) > -1, Val(txtMajorFunctionCode.Tag), -1)
        mintFunctionID = IIf(Val(txtFunctionCode.Tag) > -1, Val(txtFunctionCode.Tag), -1)
       
            arrInput = Array((IIf(mEditFlag = True, mintFunctionID, -1)), _
                            Trim(txtMajorFunctionCode.Text) & Trim(txtFunctionCode.Text), _
                            Trim(txtFunction.Text), _
                            mintMjrFunctionID _
                            )
            objDb.ExecuteSP "spSaveFunctions", arrInput
            Call FormInitialize
'        End If
    End Sub

    Private Sub Form_Activate()
        Me.Top = 50
        frmFunctions.Left = (frmMenu.Width - Me.Width) / 2
    End Sub

    Private Sub Form_Load()
         Call dispMajorFunctions
        
    End Sub

    Private Sub lstMajorFunctions_DblClick()
        gbSearchStr = lstMajorFunctions.Text
        gbSearchID = lstMajorFunctions.ItemData(lstMajorFunctions.ListIndex)
        lstMajorFunctions.Visible = False
        txtMajourFunction.SetFocus
        txtMajourFunction.Text = gbSearchStr
        txtMajorFunctionCode.Tag = gbSearchID
        
        Dim objDb As New clsDB
        Dim mCon As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        objDb.SetConnection mCon
            Rec.Open "Select vchMajorFunctionCode From faMajorFunctions Where intMajorFunctionID=" & gbSearchID, mCon
            If Not (Rec.BOF And Rec.EOF) Then
                txtMajorFunctionCode.Text = Rec!vchMajorFunctionCode
            Else
                txtMajorFunctionCode.Tag = ""
                txtMajorFunctionCode.Text = ""
                txtMajourFunction.Text = ""
            End If
            gbSearchStr = ""
            gbSearchID = -1
    End Sub
    
    Private Sub lstMajorFunctions_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub dispMajorFunctions()
    
            Dim objDb As New clsDB
            Dim mCon As ADODB.Connection
            Dim objAcc As New clsAccounts
            Dim Rec As New ADODB.Recordset
            Dim i As Long
            
            vsGrid.Clear 1, 1    'clear all rows excluding fixed rows
            vsGrid.Rows = 2
            objDb.SetConnection mCon
            Rec.Open "Select * from faMajorFunctions", mCon
            txtMajorFunctionCode.Tag = Rec!intMajorFunctionID
            i = 0
            While Not (Rec.EOF = True)
                vsGrid.TextMatrix(i + 1, 0) = Rec!vchMajorFunction
                vsGrid.TextMatrix(i + 1, 1) = Rec!vchMajorFunctionCode
                vsGrid.TextMatrix(i + 1, 3) = Rec!intMajorFunctionID
                vsGrid.TextMatrix(i + 1, 4) = False
                
                Rec.MoveNext
                vsGrid.Rows = vsGrid.Rows + 1
                i = i + 1
            Wend
            vsGrid.Rows = vsGrid.Rows - 1
    End Sub

    Private Sub txtFunctionCode_GotFocus()
        txtFunctionCode.SelStart = 0
        txtFunctionCode.SelLength = Len(txtFunctionCode.Text)
    End Sub

    Private Sub txtFunctionCode_LostFocus()
       Call dispDetails
    End Sub

    Private Sub txtMajorFunctionCode_GotFocus()

    Dim objDb As New clsDB
    Dim mCon As ADODB.Connection
    Dim Rec As New ADODB.Recordset

    objDb.SetConnection mCon
    Rec.Open "select vchFunctionCode, vchFunction from faFunctions where faFunctions.intFunctionID=" & Val(txtFunctionCode.Tag), mCon
        If Not Rec.EOF Then
                txtFunction.Text = Rec!vchFunction
                txtFunctionCode.Text = Rec!vchFunctionCode
        Else
                txtFunction.Text = ""
                txtFunctionCode.Text = ""
        End If
    End Sub

    Private Sub txtMajourFunction_GotFocus()
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            mSQL = "       SELECT * FROM faMajorFunctions "
            mSQL = mSQL + " WHERE faMajorFunctions.intMajorFunctionID = " & Val(vsGrid.TextMatrix(vsGrid.Row, 3))
            
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                txtMajourFunction.Text = Rec!vchMajorFunction
                txtMajorFunctionCode.Text = Rec!vchMajorFunctionCode
                txtMajorFunctionCode.Tag = Rec!intMajorFunctionID
                txtPrefix.Text = Rec!vchMajorFunctionCode
            End If
    End Sub



    Private Sub vsGrid_Click()
        If vsGrid.Row > 0 Then
            If Val(vsGrid.TextMatrix(vsGrid.Row, 2)) = 0 Then
                Call FormInitialize
                 
                If vsGrid.TextMatrix(vsGrid.Row, 4) = False Then
                    vsGrid.TextMatrix(vsGrid.Row, 4) = True
                    'txtMajourFunction.SetFocus
                    Call txtMajourFunction_GotFocus
                    Call DisplaySubfunctions(Val(vsGrid.TextMatrix(vsGrid.Row, 3)), vsGrid.Row)
                Else
                    vsGrid.TextMatrix(vsGrid.Row, 4) = False
                    Call DeleteSubFucntions(Val(vsGrid.TextMatrix(vsGrid.Row, 3)))
                End If
            Else
                Call DisplayFunction(Val(vsGrid.TextMatrix(vsGrid.Row, 2)))
            End If
        End If
    End Sub
  
    Private Sub FormInitialize()
        txtFunction.Text = ""
        txtFunctionCode.Text = ""
        txtFunctionCode.Tag = 0
        txtMajourFunction.Text = ""
        txtMajorFunctionCode.Text = ""
        txtMajorFunctionCode.Tag = 0
        txtPrefix.Text = ""
        mEditFlag = False
    End Sub
    Private Sub DisplayFunction(mFunctionID As Long)
           Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            mEditFlag = True
            mSQL = "       SELECT * FROM faFunctions LEFT JOIN "
            mSQL = mSQL + "faMajorFunctions ON "
            mSQL = mSQL + "faMajorFunctions.intMajorFunctionID=faFunctions.intMajorFunctionID "
            mSQL = mSQL + " WHERE faFunctions.intFunctionID = " & mFunctionID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
            
                txtMajourFunction.Text = Rec!vchMajorFunction
                txtMajorFunctionCode.Text = Rec!vchMajorFunctionCode
                txtMajorFunctionCode.Tag = Rec!intMajorFunctionID
                
                txtFunctionCode.Text = Right(Rec!vchFunctionCode, Len(Trim(Rec!vchFunctionCode)) - Len(Trim(Rec!vchMajorFunctionCode)))
                txtPrefix.Tag = CStr(Rec!vchFunctionCode)
                txtPrefix.Text = Rec!vchMajorFunctionCode
                
                txtFunction.Text = Rec!vchFunction
                txtFunctionCode.Tag = Rec!intFunctionID
                
            End If
    End Sub

    Private Sub dispDetails()
        Dim objDb As New clsDB
        Dim mCon As ADODB.Connection
        Dim objFunction As New clsFunction
        
        txtFunctionCode.Text = Trim(txtFunctionCode.Text)
        If txtFunctionCode.Text <> "" Then
            objFunction.SetFunction (Trim(txtFunctionCode.Text))
            If mEditFlag And objFunction.FunctionID > -1 Then
                If Val(txtFunctionCode.Tag) = objFunction.FunctionID Then
                    Exit Sub
                End If
            ElseIf mEditFlag Then
                Exit Sub
            End If
        
            If objFunction.FunctionID > 0 Then
                mEditFlag = True
                txtMajourFunction.Text = objFunction.MajorFunctionName
                txtMajorFunctionCode.Text = objFunction.MajorFunctionCode
                txtMajorFunctionCode.Tag = objFunction.MajorFunctionID
                
                txtFunctionCode.Text = Right(objFunction.FunctionCode, Len(Trim(objFunction.FunctionCode)) - Len(Trim(objFunction.MajorFunctionCode)))
                txtPrefix.Tag = CStr(objFunction.FunctionCode)
                txtPrefix.Text = objFunction.MajorFunctionCode
                
                txtFunction.Text = objFunction.FunctionName
                txtFunctionCode.Tag = objFunction.FunctionID
                
            ElseIf mEditFlag = True Then
                mEditFlag = False

                txtMajourFunction.Text = ""
                txtMajorFunctionCode.Text = ""
                txtMajorFunctionCode.Tag = ""

                txtFunctionCode.Text = txtFunctionCode.Text
                txtPrefix.Tag = ""
                txtPrefix.Text = ""

                txtFunction.Text = ""
                txtFunctionCode.Tag = ""
            End If
        End If
    End Sub
    
    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call vsGrid_Click
        End If
    End Sub
Private Sub FormInitializeforNew()
        txtFunction.Text = ""
        txtFunctionCode.Text = ""
        txtFunctionCode.Tag = 0
        mEditFlag = False
    End Sub
