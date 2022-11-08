VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmFields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ~ Field ~"
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
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3210
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2955
      Left            =   420
      TabIndex        =   0
      Top             =   2790
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
      FormatString    =   $"frmFields.frx":0000
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
      TabIndex        =   1
      Top             =   780
      Width           =   8925
      Begin VB.TextBox txtFieldCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Top             =   750
         Width           =   2085
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   1110
         Width           =   3885
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Field"
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
         TabIndex        =   9
         Top             =   0
         Width           =   9630
      End
      Begin VB.Label lblFunctionName 
         AutoSize        =   -1  'True
         Caption         =   "&Field"
         Height          =   195
         Left            =   3060
         TabIndex        =   4
         Top             =   1110
         Width           =   330
      End
      Begin VB.Label lblFunctionaryCode 
         AutoSize        =   -1  'True
         Caption         =   "Field &Code"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   750
         Width           =   750
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
      TabIndex        =   10
      Top             =   0
      Width           =   8940
   End
End
Attribute VB_Name = "frmFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Option Explicit
    Dim mEditFlag As Boolean
    Dim mNewFlag As Boolean

    Private Sub cmdCancel_Click()
        If mEditFlag Then
            FormInitialize
        Else
            Unload Me
        End If
    End Sub
    
    Private Sub cmdNew_Click()
        Call FormInitialize
        mEditFlag = False
        txtFieldCode.SetFocus
    End Sub

    Private Sub cmdSave_Click()
        Dim mintFieldID As Long
        Dim mCon As New ADODB.Connection
        Dim arrInput As Variant
        Dim objDB As New clsDB
        '------------------------------------------'
        '   Validations                            '
        '------------------------------------------'
        If Trim(txtFieldCode.Text) = "" Then
            txtFieldCode.SetFocus
            Exit Sub
        End If
        If Trim(txtField.Text) = "" Then
            txtField.SetFocus
            Exit Sub
        End If
        If mEditFlag And Val(txtFieldCode.Tag) < 1 Then
            MsgBox "Error: Try again!", vbInformation
            Exit Sub
        ElseIf mEditFlag And Val(txtFieldCode.Tag) > 0 Then
            MsgBox "Updating an Existing Field"
        ElseIf mEditFlag = False And Val(txtFieldCode.Tag) = 0 Then
            MsgBox "Creating a new Field"
        End If
        '------------------------------------------'
        '   Saving a New Function                  '
        '------------------------------------------'
        mintFieldID = IIf(Val(txtFieldCode.Tag) > -1, Val(txtFieldCode.Tag), -1)
        arrInput = Array(IIf(mEditFlag, mintFieldID, -1), _
                    CStr(Trim(txtFieldCode.Text)), _
                    Trim(txtField.Text), _
                    gbLocalBodyID)
        objDB.ExecuteSP "spSaveField", arrInput
        Call FormInitialize
        Call FillGrid
    End Sub

   Private Sub Form_Activate()
         Me.Top = 50
        frmFields.Left = (frmMenu.Width - Me.Width) / 2
    End Sub

    Private Sub Form_Load()
         Call FillGrid
         Call FormInitialize
    End Sub

    Private Sub FillGrid()
        Dim mCon As New ADODB.Connection
        Dim mLoopCount As Long
        Dim Rec As New ADODB.Recordset
        vsGrid.Rows = 2
        Set Rec = GetRecordSet("Select * From faFields Order By vchField")
        mLoopCount = 0
        While Not (Rec.EOF = True)
            vsGrid.TextMatrix(mLoopCount + 1, 0) = Rec!vchField
            vsGrid.TextMatrix(mLoopCount + 1, 1) = Rec!vchFieldCode
            vsGrid.TextMatrix(mLoopCount + 1, 2) = Rec!intFieldID
            Rec.MoveNext
            vsGrid.Rows = vsGrid.Rows + 1
            mLoopCount = mLoopCount + 1
        Wend
        If vsGrid.Rows < 15 Then
            vsGrid.Rows = vsGrid.Rows + (16 - mLoopCount)
            vsGrid.Row = mLoopCount - 1
        End If
    End Sub

    Private Sub FormInitialize()
        txtFieldCode.Text = ""
        txtFieldCode.Tag = ""
        txtField.Text = ""
        txtField.Tag = ""
        mEditFlag = False
    End Sub
    
    Private Sub Display(mID As Long)
        Dim Obj As New clsField
        Obj.SetField (mID)
        If Obj.FieldID > 0 Then
            mEditFlag = True
            txtFieldCode.Text = Obj.FieldCode
            txtField.Text = Obj.FieldName
            txtFieldCode.Tag = Obj.FieldID
        End If
        Set Obj = Nothing
    End Sub
    


    Private Sub txtField_GotFocus()
        Call dispDetails
    End Sub
   
    Private Sub dispDetails()
        Dim objDB As New clsDB
        Dim mCon As ADODB.Connection
        Dim objField As New clsField
        
        txtFieldCode.Text = Trim(txtFieldCode.Text)
        If txtFieldCode.Text <> "" Then
            txtField.SetFocus
            objField.SetFieldByCode (txtFieldCode.Text)
            If objField.FieldID > -1 Then
                
                If mEditFlag And Val(txtFieldCode.Tag) = objField.FieldID Then
                    Exit Sub
                Else
                    mEditFlag = True
                    txtField.Text = objField.FieldName
                    txtFieldCode.Text = objField.FieldCode
                    txtFieldCode.Tag = objField.FieldID
                End If
            ElseIf mEditFlag Then
                Exit Sub
            End If
        End If
    End Sub

    Private Sub vsGrid_Click()
        If Val(vsGrid.TextMatrix(vsGrid.Row, 2)) > 0 Then
            Call Display(Val(vsGrid.TextMatrix(vsGrid.Row, 2)))
        End If
    End Sub

    Private Sub vsGrid_RowColChange()
        Call vsGrid_Click
    End Sub
