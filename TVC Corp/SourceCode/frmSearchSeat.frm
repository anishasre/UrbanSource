VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchSeat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seat Search"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearchKey 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   4800
      Width           =   3405
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   315
      Left            =   3450
      TabIndex        =   1
      Top             =   4800
      Width           =   345
   End
   Begin VB.ListBox lstSeatName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF7&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   3645
      TabIndex        =   0
      Top             =   5175
      Visible         =   0   'False
      Width           =   3765
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   30
      Top             =   5190
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4245
      Left            =   45
      TabIndex        =   5
      Top             =   315
      Width           =   3750
      _cx             =   6615
      _cy             =   7488
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   -2147483632
      TreeColor       =   16777215
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchSeat.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "  Seats "
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
      Left            =   -30
      TabIndex        =   4
      Top             =   0
      Width           =   11625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seat Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   30
      TabIndex        =   3
      Top             =   4530
      Width           =   870
   End
End
Attribute VB_Name = "frmSearchSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private mQry As String

    

   
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vsGrid.Row > 0 And vsGrid.Text <> "" Then
            Call vsGrid_DblClick
        End If
    End If
End Sub

''''''''''''    Dim mSeatPrefix As String
''''''''''''
''''''''''''    Private Sub cmdSearch_Click()
''''''''''''        Call FillSeats(txtSearchKey.Text)
''''''''''''    End Sub
''''''''''''
''''''''''''    Private Sub Form_Load()
''''''''''''        WindowsXPC1.InitIDESubClassing
''''''''''''        If mQry <> "" Then
''''''''''''            Call FillSeatUsingQueryPassed
''''''''''''        Else
''''''''''''            Call FillSeats(txtSearchKey.Text)
''''''''''''        End If
''''''''''''    End Sub
''''''''''''    Private Sub FillSeatUsingQueryPassed()
''''''''''''        On Error GoTo Err:
''''''''''''            If objDb.CreateNewConnection(mCnn, enuSourceString.DBMaster) Then
''''''''''''                Call PopulateList(lstSeatName, mQry, , True, True, True, enuSourceString.DBMaster)
''''''''''''            Else
''''''''''''                    MsgBox "Connection To Master Does not Exist, Please Contact your System Administrator", vbInformation
''''''''''''            End If
''''''''''''            Exit Sub
''''''''''''Err:
''''''''''''            MsgBox (Error$)
''''''''''''    End Sub
''''''''''''    Private Function FillSeats(ByVal strSeat As String)
''''''''''''        On Error GoTo Err:
''''''''''''            Dim mSql As String
''''''''''''            Dim Rec As New ADODB.Recordset
''''''''''''            Dim mCnn As New ADODB.Connection
''''''''''''            Dim objDb As New clsDB
''''''''''''            Dim mQuery As String
''''''''''''
''''''''''''            If objDb.CreateNewConnection(mCnn, enuSourceString.DBMaster) Then
''''''''''''                If IsNull(strSeat) Then
''''''''''''                    mQuery = "Select Left(Convert( VarChar(10),numSeatID), 6) As Prefix From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
''''''''''''                Else
''''''''''''                    mQuery = "Select Left(Convert( VarChar(10),numSeatID), 6) As Prefix From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " and chvSeatTitle Like '%" & strSeat & "%' Order By chvSeatTitle"
''''''''''''                End If
''''''''''''                Rec.Open mQuery, mCnn
''''''''''''                If Not (Rec.EOF And Rec.BOF) Then
''''''''''''                    mSeatPrefix = IIf(IsNull(Rec!Prefix), "", Rec!Prefix)
''''''''''''                End If
''''''''''''                If Rec.State = 1 Then Rec.Close
''''''''''''                If IsNull(strSeat) Then
''''''''''''                    mSql = "Select chvSeatTitle, Right(Convert( VarChar(10),numSeatID), 5) From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
''''''''''''                Else
''''''''''''                    mSql = "Select chvSeatTitle, Right(Convert( VarChar(10),numSeatID), 5) From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " and chvSeatTitle Like '%" & strSeat & "%' Order By chvSeatTitle"
''''''''''''                End If
''''''''''''                Call PopulateList(lstSeatName, mSql, , True, True, True, enuSourceString.DBMaster)
''''''''''''            Else
''''''''''''                MsgBox "Connection To Master Does not Exist, Please Contact your System Administrator", vbInformation
''''''''''''            End If
''''''''''''        Exit Function
''''''''''''Err:
''''''''''''        MsgBox (Error$)
''''''''''''    End Function
''''''''''''
''''''''''''    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''''''''''        gbSearchStr = ""
''''''''''''        gbSearchID = -1
''''''''''''        If KeyCode = vbKeyEscape Then
''''''''''''            Unload Me
''''''''''''        ElseIf KeyCode = 13 Then
''''''''''''            If lstSeatName.ListIndex > -1 Then
''''''''''''                gbSearchStr = lstSeatName.Text
''''''''''''                gbSearchID = CDbl(mSeatPrefix + CStr(lstSeatName.ItemData(lstSeatName.ListIndex)))
''''''''''''                Unload Me
''''''''''''            End If
''''''''''''        End If
''''''''''''    End Sub
''''''''''''
''''''''''''    Private Sub lstSeatName_DblClick()
''''''''''''        Call Form_KeyDown(13, 0)
''''''''''''    End Sub
''''''''''''
''''''''''''    Private Sub txtSearchKey_Change()
''''''''''''        On Error GoTo Err:
''''''''''''            Dim mIndex As Long
''''''''''''            Dim mStr As String
''''''''''''
''''''''''''            mStr = txtSearchKey.Text
''''''''''''            mIndex = SendMyMessage(lstSeatName.hwnd, LB_FINDSTRING, -1, ByVal mStr)
''''''''''''            If mIndex > -1 Then
''''''''''''                lstSeatName.ListIndex = mIndex
''''''''''''            End If
''''''''''''        Exit Sub
''''''''''''Err:
''''''''''''        MsgBox (Error$)
''''''''''''    End Sub
''''''''''''    Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
''''''''''''        On Error GoTo Err:
''''''''''''            '38 Up Arrow
''''''''''''            If (KeyCode = 38 Or KeyCode = 40) Then
''''''''''''                Dim objAcc As New clsAccounts
''''''''''''                If KeyCode = 38 And lstSeatName.ListIndex > 0 Then
''''''''''''                    lstSeatName.ListIndex = lstSeatName.ListIndex - 1
''''''''''''                End If
''''''''''''                '40 = Down Arrow
''''''''''''                If KeyCode = 40 And lstSeatName.ListIndex < (lstSeatName.ListCount - 1) Then
''''''''''''                    lstSeatName.ListIndex = lstSeatName.ListIndex + 1
''''''''''''                    'Debug.Print lstSeatName.ListCount - 1, lstSeatName.ListIndex
''''''''''''                End If
''''''''''''                If lstSeatName.ListIndex > -1 Then
''''''''''''                    objAcc.SetAccounts (lstSeatName.ItemData(lstSeatName.ListIndex))
''''''''''''                    txtSearchKey.Text = objAcc.AccountCode
''''''''''''                End If
''''''''''''            End If
''''''''''''        Exit Sub
''''''''''''Err:
''''''''''''        MsgBox (Error$)
''''''''''''    End Sub
''''''''''''
    Private Sub Form_Load()
       Call FillGrid
'        If mQry <> "" Then
'
'        End If
    End Sub
    Private Sub cmdSearch_Click()
        Call FillGrid(txtSearchKey.Text)
    End Sub
    Public Property Let SQLString(mSql As String)
        mQry = mSql
    End Property

    Private Sub Form_Unload(Cancel As Integer)
        mQry = ""
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row <> 0 Then
            gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 0))
            gbSearchID = val(vsGrid.TextMatrix(vsGrid.Row, 1))
            mQry = ""
            Unload Me
        End If
    End Sub
    Private Sub FillGrid(Optional mString As String)
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb   As New clsDB
        Dim mSql    As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.DBMaster
        vsGrid.SelectionMode = flexSelectionByRow
        If mQry = "" Then
            If IsNull(mString) Then
                mSql = "Select chvSeatTitle, numSeatID From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
            Else
                mSql = "Select chvSeatTitle,numSeatID From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " and chvSeatTitle Like '%" & mString & "%' Order By chvSeatTitle"
            End If
        Else
            mSql = mQry
        End If
        Rec.Open mSql, mCnn
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        If Not (Rec.BOF And Rec.BOF) Then
            While Not Rec.EOF
                vsGrid.AddItem ""
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
                Rec.MoveNext
            Wend
        End If
    End Sub
    
