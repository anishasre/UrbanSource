VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmFunctionary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Functionary"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
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
      TabIndex        =   11
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3210
      TabIndex        =   12
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
      FormatString    =   $"frmFunctionary.frx":0000
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
      Begin VB.TextBox txtInstitutionCode 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4290
         MaxLength       =   4
         TabIndex        =   4
         Top             =   510
         Width           =   1635
      End
      Begin VB.TextBox txDepartmentCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         TabIndex        =   10
         Top             =   1500
         Width           =   2445
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         TabIndex        =   8
         Top             =   1170
         Width           =   3885
      End
      Begin VB.TextBox txtInstitution 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         TabIndex        =   6
         Top             =   840
         Width           =   3885
      End
      Begin VB.TextBox txtPrefix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Department Code"
         Height          =   195
         Left            =   2130
         TabIndex        =   9
         Top             =   1530
         Width           =   1245
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Functionary"
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
         TabIndex        =   14
         Top             =   0
         Width           =   9630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Department"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblFunctionName 
         AutoSize        =   -1  'True
         Caption         =   "&Institution"
         Height          =   195
         Left            =   2670
         TabIndex        =   5
         Top             =   870
         Width           =   675
      End
      Begin VB.Label lblFunctionaryCode 
         AutoSize        =   -1  'True
         Caption         =   "Institition &Code"
         Height          =   195
         Left            =   2310
         TabIndex        =   2
         Top             =   540
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
      TabIndex        =   15
      Top             =   0
      Width           =   8940
   End
End
Attribute VB_Name = "frmFunctionary"
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
    
    Private Sub DisplaySubfunctions(mMajorFunctionaryID As Long, mRow As Long)
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        mSQL = " SELECT * FROM faFunctionaries WHERE intMajorFunctionaryID = " & mMajorFunctionaryID & " ORDER BY vchFunctionary"
        Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                    'vsGrid.AddItem ""
                    mRow = mRow + 1
                    vsGrid.AddItem "         " & Rec!vchfunctionary & vbTab & Rec!vchFunctionaryCode & vbTab & Rec!intFunctionaryID & vbTab & Rec!intMajorFunctionaryID & vbTab & "False", mRow
                    Rec.MoveNext
                Wend
            End If
        Rec.Close
    End Sub
        
    Private Sub cmdCancel_Click()
        If mEditFlag Then
            Call dispMajorFunctionary
            Call FormInitialize
            Else
                Unload Me
            End If
        
    End Sub
    
    Private Sub cmdNew_Click()
        Call FormInitializeForNew
        mEditFlag = False
        mNewFlag = True
        txtInstitutionCode.SetFocus
    End Sub
        
    Private Sub cmdSave_Click()
        
        Dim mintMjrFunctionaryID As Long
        Dim mintFunctionaryID As Long
        Dim mCon As New ADODB.Connection
        Dim arrInput As Variant
        
        Dim objDB As New clsDB
        
        '------------------------------------------
        '   Validations
        '------------------------------------------
        If Trim(txtInstitutionCode.Text) = "" Then
            txtInstitutionCode.SetFocus
            Exit Sub
        End If
        If Trim(txtInstitution.Text) = "" Then
            txtInstitution.SetFocus
            Exit Sub
        End If
        If mEditFlag And Val(txtInstitutionCode.Tag) < 1 Then
            MsgBox "Error: Try again!", vbInformation
            Exit Sub
        ElseIf mEditFlag And Val(txtInstitutionCode.Tag) > 0 Then
            MsgBox "Updating an Existing Institution"
        ElseIf mEditFlag = False And Val(txtInstitutionCode.Tag) = 0 Then
            MsgBox "Creating a new Institution"
        End If
                
        '------------------------------------------
        '   Saving a New Functionary
        '------------------------------------------
        
        objDB.SetConnection mCon
        mintMjrFunctionaryID = IIf(Val(txDepartmentCode.Tag) > -1, Val(txDepartmentCode.Tag), -1)
        mintFunctionaryID = IIf(Val(txtInstitutionCode.Tag) > -1, Val(txtInstitutionCode.Tag), -1)
            arrInput = Array((IIf(mEditFlag = True, mintFunctionaryID, -1)), _
                            Trim(txDepartmentCode.Text) & Trim(txtInstitutionCode.Text), _
                            Trim(txtInstitution.Text), _
                            mintMjrFunctionaryID _
                            )
            objDB.ExecuteSP "spSaveFunctionaries", arrInput
            Call FormInitialize
    End Sub

    Private Sub Form_Activate()
        Me.Top = 50
        frmFunctionary.Left = (frmMenu.Width - Me.Width) / 2
    End Sub

    Private Sub Form_Load()
         Call dispMajorFunctionary
    End Sub

    
    Private Sub dispMajorFunctionary()
    
            Dim objDB As New clsDB
            Dim mCon As ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim i As Long
            
            vsGrid.Clear 1, 1    'clear all rows excluding fixed rows
            vsGrid.Rows = 2
            objDB.SetConnection mCon
            Rec.Open "Select * from faMajorFunctionaries", mCon
            txDepartmentCode.Tag = Rec!intDepartmentID
            i = 0
            While Not (Rec.EOF = True)
                vsGrid.TextMatrix(i + 1, 0) = Rec!vchDepartment
                vsGrid.TextMatrix(i + 1, 1) = Rec!vchDepartmentCode
                vsGrid.TextMatrix(i + 1, 3) = Rec!intDepartmentID
                vsGrid.TextMatrix(i + 1, 4) = False
                
                Rec.MoveNext
                vsGrid.Rows = vsGrid.Rows + 1
                i = i + 1
            Wend
            vsGrid.Rows = vsGrid.Rows - 1
    End Sub

    Private Sub txtDepartmentCode_GotFocus()

        Dim objDB As New clsDB
        Dim mCon As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        objDB.SetConnection mCon
        Rec.Open "select vchFunctionaryCode, vchFunctionary from faFunctionaries where faFunctionaries.intFunctionaryID=" & Val(txtInstitutionCode.Tag), mCon
            If Not Rec.EOF Then
                    txtInstitution.Text = Rec!vchfunctionary
                    txtInstitutionCode.Text = Rec!vchFunctionaryCode
            Else
                    txtInstitution.Text = ""
                    txtInstitutionCode.Text = ""
            End If
    End Sub


    Private Sub txtDepartment_GotFocus()
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            mSQL = "       SELECT * FROM faMajorFunctionaries "
            mSQL = mSQL + " WHERE faMajorFunctionaries.intDepartmentID = " & Val(vsGrid.TextMatrix(vsGrid.Row, 3))
    
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                txtDepartment.Text = Rec!vchDepartment
                txDepartmentCode.Text = Rec!vchDepartmentCode
                txDepartmentCode.Tag = Rec!intDepartmentID
                txtPrefix.Text = Rec!vchDepartmentCode
            End If
    End Sub

    Private Sub txtInstitutionCode_LostFocus()
        Call dispDetails
        
    End Sub

    Private Sub vsGrid_Click()
        If vsGrid.Row > 0 Then
            If Val(vsGrid.TextMatrix(vsGrid.Row, 2)) = 0 Then
                Call FormInitialize
                 
                If vsGrid.TextMatrix(vsGrid.Row, 4) = False Then
                    vsGrid.TextMatrix(vsGrid.Row, 4) = True
                    'txtDepartment.SetFocus
                    Call txtDepartment_GotFocus
                    Call DisplaySubfunctions(Val(vsGrid.TextMatrix(vsGrid.Row, 3)), vsGrid.Row)
                Else
                    vsGrid.TextMatrix(vsGrid.Row, 4) = False
                    Call DeleteSubFucntions(Val(vsGrid.TextMatrix(vsGrid.Row, 3)))
                End If
            Else
                Call DisplayFunction(Val(vsGrid.TextMatrix(vsGrid.Row, 2)))
            End If
            Else
        End If
    End Sub
  
    Private Sub FormInitialize()
        txtInstitution.Text = ""
        txtInstitutionCode.Text = ""
        txtInstitutionCode.Tag = 0
        txtDepartment.Text = ""
        txDepartmentCode.Text = ""
        txDepartmentCode.Tag = 0
        txtPrefix.Text = ""
        mEditFlag = False
    End Sub
    
    Private Sub DisplayFunction(mFunctionaryID As Long)
           Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            mEditFlag = True
            mSQL = "       SELECT * FROM faFunctionaries LEFT JOIN "
            mSQL = mSQL + "faMajorFunctionaries ON "
            mSQL = mSQL + "faMajorFunctionaries.intDepartmentID=faFunctionaries.intMajorFunctionaryID "
            mSQL = mSQL + " WHERE faFunctionaries.intFunctionaryID = " & mFunctionaryID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.BOF And Rec.EOF) Then
                txtDepartment.Text = Rec!vchDepartment
                txDepartmentCode.Text = Rec!vchDepartmentCode
                txDepartmentCode.Tag = Rec!intDepartmentID
                txtPrefix.Text = Rec!vchDepartmentCode
                txtInstitutionCode.Text = Right(Rec!vchFunctionaryCode, Len(Trim(Rec!vchFunctionaryCode)) - Len(Trim(Rec!vchDepartmentCode)))
                txtPrefix.Tag = CStr(Rec!vchFunctionaryCode)
                txtInstitution.Text = Rec!vchfunctionary
                txtInstitutionCode.Tag = Rec!intFunctionaryID
            End If
    End Sub

    Private Sub dispDetails()
            Dim objDB As New clsDB
            Dim mCon As ADODB.Connection
            Dim objFunctionary As New clsFunctionary
            
            txtInstitutionCode.Text = Trim(txtInstitutionCode.Text)
            If txtInstitutionCode.Text <> "" Then
                objFunctionary.SetFunctionary (Trim(txtPrefix.Text) + (txtInstitutionCode.Text))
                If mEditFlag And objFunctionary.FunctionaryID > -1 Then
                    If Val(txtInstitutionCode.Tag) = objFunctionary.FunctionaryID Then
                        Exit Sub
                    End If
                ElseIf mEditFlag Then
                    Exit Sub
                End If
            
                If objFunctionary.FunctionaryID > 0 And objFunctionary.FunctionaryCode <> "" Then
                
                    mEditFlag = True
                    txtDepartment.Text = objFunctionary.MajorFunctionaryName
                    txDepartmentCode.Text = objFunctionary.MajorFunctionaryCode
                    txDepartmentCode.Tag = objFunctionary.MajorFunctionaryID
                    
                    txtInstitutionCode.Text = Right(objFunctionary.FunctionaryCode, Len(Trim(objFunctionary.FunctionaryCode)) - Len(Trim(objFunctionary.MajorFunctionaryCode)))
                    txtInstitution.Text = objFunctionary.FunctionaryName
                    txtInstitutionCode.Tag = objFunctionary.FunctionaryCode
                    
                    txtPrefix.Tag = CStr(objFunctionary.FunctionaryCode)
                    txtPrefix.Text = objFunctionary.MajorFunctionaryCode
                    
                ElseIf objFunctionary.FunctionaryID < 0 And objFunctionary.FunctionaryCode <> "" Then
                    mEditFlag = False
                    txtDepartment.Text = ""
                    txDepartmentCode.Text = ""
                    txDepartmentCode.Tag = ""
                    
                    txtInstitutionCode.Text = ""
                    txtInstitution.Text = ""
                    txtInstitutionCode.Tag = ""
                    
                    txtPrefix.Tag = ""
                    txtPrefix.Text = ""
                    
                End If
            End If
        End Sub

        Private Sub vsGrid_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                Call vsGrid_Click
            End If
        End Sub
Private Sub FormInitializeForNew()
        txtInstitution.Text = ""
        txtInstitutionCode.Text = ""
        txtInstitutionCode.Tag = 0
        mEditFlag = False
    End Sub
