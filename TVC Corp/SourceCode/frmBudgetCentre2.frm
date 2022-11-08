VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmBudgetCentre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "B u d g e t    C e n t r e"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2205
      Left            =   30
      TabIndex        =   19
      Top             =   2940
      Width           =   8205
      _cx             =   14473
      _cy             =   3889
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBudgetCentre2.frx":0000
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8250
      TabIndex        =   9
      Top             =   0
      Width           =   8250
   End
   Begin VB.Frame Frame2 
      Height          =   2235
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   8265
      Begin VB.TextBox txtBudgetCentreCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   18
         Text            =   "000000"
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox txtBudgetCentreName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3540
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   3315
      End
      Begin VB.CommandButton cmdField 
         Caption         =   "..."
         Height          =   315
         Left            =   6900
         TabIndex        =   15
         Top             =   1590
         Width           =   375
      End
      Begin VB.CommandButton cmdFunctionary 
         Caption         =   "..."
         Height          =   315
         Left            =   6900
         TabIndex        =   14
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton cmdFunction 
         Caption         =   "..."
         Height          =   315
         Left            =   6900
         TabIndex        =   13
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3540
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1590
         Width           =   3315
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3540
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3540
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   930
         Width           =   3315
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1590
         Width           =   2115
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1260
         Width           =   2115
      End
      Begin VB.TextBox txtFunctions 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   930
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Budget Centre"
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
         TabIndex        =   17
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Field"
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
         Left            =   930
         TabIndex        =   7
         Top             =   1620
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Functionary"
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
         Left            =   390
         TabIndex        =   5
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Function"
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
         Left            =   630
         TabIndex        =   3
         Top             =   990
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   4050
      TabIndex        =   1
      Top             =   5340
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   5340
      Width           =   1215
   End
End
Attribute VB_Name = "frmBudgetCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
            
    Dim mCon As New ADODB.Connection
    Dim mCom As New ADODB.Command
    Dim RecFunction As New ADODB.Recordset
    Dim RecFunctionary As New ADODB.Recordset
    Dim RecField As New ADODB.Recordset
    Dim Rec As New ADODB.Recordset
    Dim RecBudgetCentre As New ADODB.Recordset
    Private marrFunctionCode() As String
    Private marrFunctionaryCode() As String
    Private marrFieldCode() As String

    Private Sub cmbField_Click()
'        If Val(txtBudgetCentreCode.Tag) < 1 Then
           txtField.Text = ""
'        Else
            txtField.Text = marrFieldCode(cmbField.ListIndex, 0)
        'End If
    
    End Sub
    
    Private Sub cmbFunctionary_Click()
'        If Val(txtBudgetCentreCode.Tag) < 1 Then
'            Exit Sub
'        Else
            txtFunctionary.Text = marrFunctionaryCode(cmbFunctionary.ListIndex, 0)
        'End If
            
    End Sub
    
    Private Sub cmbFunctions_Click()
'        If Val(txtBudgetCentreCode.Tag) < 1 Then
'            Exit Sub
'        Else
            txtFunctions.Text = marrFunctionCode(cmbFunctions.ListIndex, 0)
        'End If
    End Sub
    
    Private Sub cmdSave_Click()
    
            Dim mintFunctionaryID As Long
            Dim mintFunctionID As Long
            Dim mintFieldID As Long
            Dim arrInput As Variant
            
            Dim cntFunction As Integer
            Dim cntFunctionary As Integer
            Dim cntFeild As Integer
                        
            Dim objDb As New clsDB
            Set Rec = New ADODB.Recordset
             
        '------------------------------------------
        '   Validations
        '------------------------------------------
        If txtBudgetCentreCode.Text = "" Or txtBudgetCentreName.Text = "" Then
            
            MsgBox "Sorry!!!please provide a new Code and Name for the new Budget Centre"
            Exit Sub
         End If
         
        '------------------------------------------
        '   Saving a New BudgetCentre
        '------------------------------------------
                   
            objDb.SetConnection mCon
            'Rec.Open "Select Count(*) from faBudgetCentres where faBudgetCentres.vchBudgetCentreCode='" & txtBudgetCentreCode.Text & "'", mCon
            'If (Rec(0)) = 0 Then
            If Val(txtBudgetCentreCode.Tag) < 1 Then
                mintFunctionaryID = IIf(cmbFunctionary.ListIndex > -1, cmbFunctionary.ItemData(cmbFunctionary.ListIndex), -1)
                mintFunctionID = IIf(cmbFunctions.ListIndex > -1, cmbFunctions.ItemData(cmbFunctions.ListIndex), -1)
                mintFieldID = IIf(cmbField.ListIndex > -1, cmbField.ItemData(cmbField.ListIndex), -1)
                
                arrInput = Array(-1, _
                                    Trim(txtBudgetCentreCode.Text), _
                                    Trim(txtBudgetCentreName.Text), _
                                    mintFunctionaryID, _
                                    mintFunctionID, _
                                    mintFieldID, _
                                    gbFinancialYearID, _
                                    gbLocalBodyID _
                                    )
                
                objDb.ExecuteSP "spSaveBudgetCentre", arrInput
                MsgBox "you have created a new  B u d g e t  C e n t r e"
                Call FormInitialize
                        
        Else
            MsgBox " Sorry!!This Budget Centre already exist"
            Call FormInitialize
                                   
        End If
        
        
        
'        '..................copied.....................
'
'
'        '-------------------------------------------------'
'
'                        '--------------------------------------------'
'                        ' Creating a new Cheque Book
'                        '--------------------------------------------'
'
'                            '--------------------------------------------'
'                            ' Updating an existing Cheque Book
'                            '--------------------------------------------'
'                        Else
'                            arrInput = Array(Val(txtBookNo.Tag), _
'                                                Val(Trim(txtBookNo.Text)), _
'                                                Trim(txtPrefix.Text), _
'                                                Val(Trim(txtSerialStartNo.Text)), _
'                                                Val(Trim(txtSerialLastNo.Text)), _
'                                                Format(dtBookDate.Value, "dd/mmm/yyyy"))
'                            objDb.ExecuteSP "spUpdateChequeBook", arrInput
'                            arrInput = Array((objAcc.AccountHeadID))
'                            Call objDb.ExecuteSP("spGetPreviousChequeBooks", arrInput, arrOutBookCnt)
'                        End If
'                        Call ClearChequeBookForm
'                    End If
'                End Sub
'
'
    End Sub
    
    Private Sub cmdCancel_Click()

        Unload Me

    End Sub

    Private Sub Form_Load()
    
    Dim cntFunction As Integer
    Dim cntFunctionary As Integer
    Dim cntFeild As Integer
    Dim objDb As New clsDB
    objDb.SetConnection mCon
    Set RecFunction = New ADODB.Recordset
    Set RecFunctionary = New ADODB.Recordset
    Set RecField = New ADODB.Recordset
       
       RecFunction.Open "Select * from faFunctions order by faFunctions.vchFunction", mCon, adOpenStatic, adLockOptimistic
        cntFunction = 0
        txtFunctions.Text = ""
        cmbFunctions.Clear

        While RecFunction.EOF <> True
            ReDim Preserve marrFunctionCode(152, 0)

            cmbFunctions.AddItem RecFunction!vchFunction
            cmbFunctions.ItemData(cmbFunctions.NewIndex) = RecFunction!intFunctionID
            marrFunctionCode(cntFunction, 0) = RecFunction!vchFunctionCode

            cntFunction = cntFunction + 1
            RecFunction.MoveNext

        Wend
        RecFunction.Close


        RecFunctionary.Open "Select * from faFunctionaries order by faFunctionaries.vchFunctionary", mCon, adOpenStatic, adLockOptimistic

            cntFunctionary = 0
            txtFunctionary.Text = ""
            cmbFunctionary.Clear
            While RecFunctionary.EOF <> True
                ReDim Preserve marrFunctionaryCode(30, 0)
                cmbFunctionary.AddItem RecFunctionary!vchfunctionary
                cmbFunctionary.ItemData(cmbFunctionary.NewIndex) = RecFunctionary!intFunctionaryID
                marrFunctionaryCode(cntFunctionary, 0) = RecFunctionary!vchFunctionaryCode

                cntFunctionary = cntFunctionary + 1
                RecFunctionary.MoveNext
            Wend
        RecFunctionary.Close


        RecField.Open "Select * from faFields order by faFields.vchField", mCon, adOpenStatic, adLockOptimistic
            cntField = 0
            txtField.Text = ""
            cmbField.Clear
            While RecField.EOF <> True
                ReDim Preserve marrFieldCode(19, 0)

                cmbField.AddItem RecField!vchField
                cmbField.ItemData(cmbField.NewIndex) = RecField!intFieldID
                marrFieldCode(cntField, 0) = RecField!vchFieldCode

                cntField = cntField + 1
                RecField.MoveNext
            Wend
        RecField.Close
    End Sub

    Private Sub txtBudgetCentreCode_LostFocus()
    
        Dim objDb As New clsDB
        Dim mCon As ADODB.Connection
        Dim Rec As ADODB.Recordset
                                                        
        Set Rec = New ADODB.Recordset
        
            objDb.SetConnection mCon
            txtBudgetCentreCode = Trim(txtBudgetCentreCode)
            
            Rec.Open "Select * from faBudgetCentres where faBudgetCentres.vchBudgetCentreCode = " & Val(txtBudgetCentreCode.Text), mCon
                If Not (Rec.EOF Or Rec.BOF) Then
                        
                        txtBudgetCentreCode.Tag = Rec!intBudgetCentreID
                        txtBudgetCentreName = Rec!vchBudgetCentre
                End If
                
            If Len(Trim(txtBudgetCentreCode)) Then
                    Dim objBudCen As New clsBudgetCentre
                    objBudCen.SetBudgetCentre (txtBudgetCentreCode.Text)
                    If objBudCen.BudgetCentreID < 0 Then
                        Call FormInitialize
                        Else
                        cmbFunctions.Text = objBudCen.FunctionName
                        cmbFunctionary.Text = objBudCen.FunctionaryName
                        cmbField.Text = objBudCen.FieldName
                    End If
            End If
    End Sub
    
    Private Function FormInitialize()
    
            txtBudgetCentreCode.Tag = 0
            txtBudgetCentreCode = txtBudgetCentreCode.Text
            txtBudgetCentreName.Text = ""
            txtFunctions = ""
            cmbFunctions.ListIndex = -1
            txtFunctionary.Text = ""
            cmbFunctionary.ListIndex = -1
            txtField.Text = ""
            cmbField.ListIndex = -1
            
    End Function
