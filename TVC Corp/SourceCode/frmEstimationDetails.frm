VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmEstimationDetails 
   BackColor       =   &H00DAECFA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimation Details"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   Icon            =   "frmEstimationDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProjectNameEng 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5393
      TabIndex        =   5
      Top             =   165
      Width           =   1530
   End
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "ML-TTRevathi"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1418
      TabIndex        =   4
      Top             =   495
      Width           =   5505
   End
   Begin VB.TextBox txtProjectNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1418
      TabIndex        =   3
      Top             =   165
      Width           =   1530
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2925
      Left            =   75
      TabIndex        =   0
      Top             =   885
      Width           =   7005
      _cx             =   12356
      _cy             =   5159
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14347514
      ForeColor       =   -2147483640
      BackColorFixed  =   10736893
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14347514
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEstimationDetails.frx":1CCA
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
   Begin VB.Label lblProjectNameEng 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name in English"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3270
      TabIndex        =   6
      Top             =   165
      Width           =   2085
   End
   Begin VB.Label lblProjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   218
      TabIndex        =   2
      Top             =   495
      Width           =   1185
   End
   Begin VB.Label lblProjectNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   488
      TabIndex        =   1
      Top             =   165
      Width           =   915
   End
End
Attribute VB_Name = "frmEstimationDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mAllotment  As Integer
    
    '*********************************************************************************************'
    '                   Form to View the Fundwise Details of a Project                            '
    '*********************************************************************************************'
    
    Private Sub Form_Activate()
        'Me.Top = 2000
        'Me.Left = 600
    End Sub

    Private Sub Form_Load()
        Me.Height = 4035
        Me.Width = 7230
        vsGrid.SelectionMode = flexSelectionByRow
    End Sub
    
    Private Sub vsGrid_DblClick()
        Dim mCnn    As New ADODB.Connection
        Dim mSQL    As String
        Dim objdb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        'Note:- Added by Aiby
        If vsGrid.Row <> 0 Then
            With gbProject
                .decProjectID = Null
                .intLBID = Null
                .intYearID = Null
                .intProjectSlNo = Null
                .chvProjectSlNo = Null
                .chvProjectName = Null
                .chvProjectnameEnglish = Null
                .intProjCatID = Null
                .chvDPCOrderNo = Null
                .dtDPCOrderDate = Null
                .intSectorTypeID = Null
                .intPlanID = Null
                .intSourceOfFundID = Null
                .fltEstSourceAmt = Null
            End With
            'Block End
            If mAllotment = 1 Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" And vsGrid.Row <> 0 Then
                    frmAllotmentLetter.txtProjectNo.Text = ""
                    frmAllotmentLetter.txtProjectNo.Tag = ""
                    frmAllotmentLetter.txtProjectName.Text = ""
                    frmAllotmentLetter.txtProjectName.Tag = ""
                    frmAllotmentLetter.txtProjectNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 4) 'Project ID
                    frmAllotmentLetter.txtProjectName.Tag = vsGrid.TextMatrix(vsGrid.Row, 5) 'Fund ID
                    frmAllotmentLetter.txtAmountInFigures.Tag = val(vsGrid.TextMatrix(vsGrid.Row, 1))
                    
                    If frmAllotmentLetter.cmbCategory.ListIndex = -1 Then
                        MsgBox "Please select the Category", vbInformation
                    Else
                        If frmAllotmentLetter.cmbCategory.ItemData(frmAllotmentLetter.cmbCategory.ListIndex) <> 0 Then
                            If frmAllotmentLetter.txtProjectName.Tag = frmAllotmentLetter.cmbCategory.ItemData(frmAllotmentLetter.cmbCategory.ListIndex) Then
                            'If frmAllotmentLetter.txtProjectName.Tag = 1 Or frmAllotmentLetter.txtProjectName.Tag = 3 Or frmAllotmentLetter.txtProjectName.Tag = 16 Or frmAllotmentLetter.txtProjectName.Tag = 17 Then
                                If frmAllotmentLetter.txtProjectNo.Tag <> "" Then
                                    mSQL = "Select * From suProjectDetails"
                                    mSQL = mSQL + " Where decProjectID =" & frmAllotmentLetter.txtProjectNo.Tag
                                    Rec.Open mSQL, mCnn
                                    If Not (Rec.EOF And Rec.BOF) Then
                                        frmAllotmentLetter.txtProjectNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
                                        frmAllotmentLetter.txtProjectName.Text = IIf(IsNull(Rec!chvProjectName), "", Rec!chvProjectName)
                                    End If
                                    Rec.Close
                                End If
                                Me.Mode = 0
                                Unload Me
                                Unload frmSulekhaIntegration
                            End If
                        Else
                            MsgBox "Please select the Category", vbInformation
                        End If
                    End If
                End If
            Else ' Note:- Else Part Added By Aiby
                With gbProject
                    .decProjectID = val(txtProjectNo.Tag)
                    .intProjectSlNo = Null
                    .chvProjectSlNo = txtProjectNo
                    .chvProjectName = txtProjectName.Text
                    .chvProjectnameEnglish = txtProjectNameEng
                    .intSourceOfFundID = val(vsGrid.TextMatrix(vsGrid.Row, 5))
                    .fltEstSourceAmt = val(vsGrid.TextMatrix(vsGrid.Row, 1))
                End With
                Unload Me
                Unload frmSulekhaIntegration
            End If
        End If
    End Sub
    Public Property Let Mode(mMode As Integer)
        mAllotment = mMode
    End Property
