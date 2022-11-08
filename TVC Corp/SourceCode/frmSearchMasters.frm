VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchMasters 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchMasters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4515
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   8085
      _cx             =   14261
      _cy             =   7964
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
      BackColorFixed  =   16761024
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
      Rows            =   14
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchMasters.frx":1CCA
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Search Masters"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
   End
End
Attribute VB_Name = "frmSearchMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private mQry As String
    Private enmDataBase As enuSourceString
    Private enmQryOrSP As QryOrSp
    
    Public Enum QryOrSp
        Qyery = 1
        StoredProcedure = 2
    End Enum

    Public Property Let SQLQry(mData As String)
        mQry = mData
    End Property
    
    Public Property Let Connection(mData As enuSourceString)
        enmDataBase = mData
    End Property
    
    Public Property Let QrySP(mData As QryOrSp)
        enmQryOrSP = mData
    End Property
    
    
    Private Function LoadValidation() As Boolean
        On Error GoTo Err:
            If mQry = "" Then
                MsgBox "Please Give the SQL Query, Selecting 2 Parameters, First the Index Part & Second the String Part  (*** as Form Property ***)", vbInformation
                LoadValidation = False
                Exit Function
            End If
            If enmDataBase = 0 Then
                MsgBox "Please Give the enuSourceString to Which the Query Should be Executed (*** as Form Property ***)", vbInformation
                LoadValidation = False
                Exit Function
            End If
            
            LoadValidation = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function FillGrid() As Boolean
        On Error GoTo Err:
            Dim Rec As New ADODB.Recordset
            Dim mRowCnt As Integer
            Dim mData As String
            Dim mCnn As New ADODB.Connection
            Dim objDb As New clsDB
            
            mRowCnt = 1
            vsGrid.Rows = 10
            If objDb.CreateNewConnection(mCnn, enmDataBase) Then
                Rec.CursorLocation = adUseClient
                If enmQryOrSP = Qyery Then
                    Set Rec = objDb.ExecuteSP(mQry, , , , mCnn, adCmdText)
                Else
                    Set Rec = objDb.ExecuteSP(mQry, , , , mCnn, adCmdStoredProc)
                End If
                mRowCnt = 1
                vsGrid.Rows = 2
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = Rec(0)
                    vsGrid.TextMatrix(mRowCnt, 1) = Rec(1)
'''                    If Rec.Fields(2) Is Nothing Then
'''                        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec(2)), "", Rec(2)) 'ADDED ON 31-08-2011
'''                    End If
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
                '''If Not (Rec.EOF Or Rec.BOF) Then
                '''    vsGrid.Rows = Rec.RecordCount + 1
                '''    vsGrid.Col = 0
                '''    vsGrid.Row = 1
                '''    vsGrid.ColHidden(0) = True
                '''    vsGrid.ColSel = 1
                '''    vsGrid.RowSel = vsGrid.Rows - 1
                '''    mData = Rec.GetString(, , vbTab, Chr(13))
                '''    vsGrid.Clip = mData
                '''End If
                FillGrid = True
            Else
                MsgBox "Connection Cannot be Established, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    
    Private Sub Form_Load()
    
        gbSearchID = -1
        gbSearchCode = ""
        gbSearchStr = ""
        
        If LoadValidation = True Then
            Call FillGrid
        End If
    End Sub

    Private Sub vsGrid_Click()
        On Error GoTo Err:
            vsGrid.Cell(flexcpBackColor, 1, 0, vsGrid.Rows - 1, 1) = vbWhite
            If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
                vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 1) = &HC0C0FF
            Else
                vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 1) = vbWhite
            End If
        Exit Sub
Err:
        MsgBox (Error$)
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
                    gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 1)
                    gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 0)
                    'gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 2)
                    Unload Me
                End If
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
