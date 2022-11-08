VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCounter 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Counter List"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   30
      Top             =   4410
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2565
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   6765
      _cx             =   11933
      _cy             =   4524
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
      BackColorAlternate=   -2147483624
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCounter.frx":0000
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
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6765
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   5400
         TabIndex        =   12
         Top             =   1110
         Width           =   1095
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   11
         Top             =   690
         Width           =   3345
      End
      Begin VB.TextBox txtCounterNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5820
         TabIndex        =   10
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtIP4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3060
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtIP3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2460
         TabIndex        =   8
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtIP2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1860
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtIP1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkDeactivate 
         Appearance      =   0  'Flat
         Caption         =   "Deactivate"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4770
         TabIndex        =   4
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   735
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Counter No"
         Height          =   195
         Left            =   4770
         TabIndex        =   2
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP Address"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   285
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Dim mIPAddress As String
    Dim mLen As Integer
    

    Private Sub cmdSave_Click()
        If Trim(txtIP1.Text) <> "" And Trim(txtIP2.Text) <> "" And Trim(txtIP3.Text) <> "" And Trim(txtIP4.Text) <> "" Then
            mIPAddress = Trim(txtIP1.Text) & "." & Trim(txtIP2.Text) & "." & Trim(txtIP3.Text) & "." & Trim(txtIP4.Text)
            lSubSaveCounter
        End If
        
        
    End Sub


    Private Sub Form_Load()
    ViewDetails
        Dim mLoopCtrl As Integer
        Dim mcount As Integer
        Dim mIpSub As String
        vsGrid.AutoResize = True
        WindowsXPC.InitIDESubClassing
        mIPAddress = GetIPAddress()
        mLen = Len(mIPAddress)
        mIpSub = ""
        mcount = 1
        For mLoopCtrl = 1 To mLen
            If mID(mIPAddress, mLoopCtrl, 1) = "." Then
                If mcount = 1 Then
                    txtIP1.Text = mIpSub
                    mIpSub = ""
                    mcount = mcount + 1
                ElseIf mcount = 2 Then
                    txtIP2.Text = mIpSub
                    mIpSub = ""
                    mcount = mcount + 1
                ElseIf mcount = 3 Then
                    txtIP3.Text = mIpSub
                    mIpSub = ""
                    mcount = mcount + 1
                End If
                mIpSub = ""
            Else
                mIpSub = mIpSub & mID(mIPAddress, mLoopCtrl, 1)
            End If
                txtIP4.Text = mIpSub
        Next mLoopCtrl
'        vsGrid.MergeCells = flexMergeFree
'        vsGrid.MergeRow(2) = True
        
    End Sub
     Private Sub lSubSaveCounter()
        Dim mVarrIn(5) As Variant
        Dim objDb As New clsDB
        Dim Rec As New ADODB.Recordset
        mVarrIn(0) = txtCounterNo.Text
        mVarrIn(1) = gbLocalBodyID
        mVarrIn(2) = txtDescription.Text
        mVarrIn(3) = mIPAddress
        mVarrIn(4) = chkDeactivate.Value
        mVarrIn(5) = 0
    If Trim(txtCounterNo.Text) <> "" And Trim(txtDescription.Text) <> "" Then
       Set Rec = objDb.ExecuteSP("spSaveCounterdetails", mVarrIn)
    End If
    ViewDetails
End Sub
Private Sub ViewDetails()
    Dim objDb As New clsDB
    Dim mCon As New ADODB.Connection
    Dim mVarrOut As Variant
    Dim Display As New ADODB.Recordset
    Dim i As Integer
    
       Set Display = objDb.ExecuteSP("spSelectCounter", , mVarrOut, , mCon, adCmdStoredProc)
'        If (objDb.SetConnection(mCon)) Then
'            rec1.Open "exec spSelectAll", mCon
'            mVarrOut = rec1.GetRows
'        End If
    
        If IsArray(mVarrOut) Then
            vsGrid.Rows = UBound(mVarrOut, 2) + 2
        
            For i = 0 To UBound(mVarrOut, 2)
            vsGrid.TextMatrix(i + 1, 1) = mVarrOut(0, i)
            vsGrid.TextMatrix(i + 1, 2) = mVarrOut(1, i)
            vsGrid.TextMatrix(i + 1, 3) = mVarrOut(2, i)
            vsGrid.TextMatrix(i + 1, 4) = mVarrOut(3, i)
            Next i
        End If
End Sub
    Private Sub FillGrid()
        Dim mRow As Integer
        vsGrid.Clear
        For mRow = 1 To 5
            vsGrid.Cell(flexcpText, mRow, 1) = Trim(txtCounterNo)
            vsGrid.Cell(flexcpText, mRow, 2) = Trim(txtDescription)
            vsGrid.Cell(flexcpText, mRow, 3) = mIPAddress
            If chkDeactivate.Value = 1 Then
                vsGrid.Cell(flexcpChecked, mRow, 4) = flexcpChecked
            Else
                vsGrid.Cell(flexcpChecked, mRow, 4) = 2
            End If
        Next mRow
    End Sub
Private Sub txtCounterNo_LostFocus()
        Dim objDb As New clsDB
        Dim mRec As New ADODB.Recordset
        Dim mCon As New ADODB.Connection
        Dim mVarrIn(0) As Variant
        Dim mVarrOut As Variant
        mVarrIn(0) = txtCounterNo.Text
       If (objDb.SetConnection(mCon)) Then
            If Trim(txtCounterNo.Text) <> "" Then
                mRec.Open "Select count(*) from faCounters where intCounterNo=" & txtCounterNo.Text, mCon
                    If mRec(0) <> 0 Then
                        MsgBox "Counter already exists"
                    End If
            End If
    End If
End Sub

Private Sub vsGrid_Click()
    txtCounterNo.Text = vsGrid.TextMatrix(vsGrid.RowSel, 1)
    txtDescription.Text = vsGrid.TextMatrix(vsGrid.RowSel, 2)
    mIPAddress = vsGrid.TextMatrix(vsGrid.RowSel, 3)
    chkDeactivate.Value = vsGrid.TextMatrix(vsGrid.RowSel, 4)
End Sub
