VERSION 5.00
Begin VB.Form frmSelectAccountHeads 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Account Heads"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4965
      Left            =   0
      TabIndex        =   2
      Top             =   -90
      Width           =   9900
      Begin VB.OptionButton optAssets 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Assets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optLiabilities 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Liabilities"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   300
         TabIndex        =   10
         Top             =   480
         Width           =   945
      End
      Begin VB.ListBox lstSelected 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3885
         Left            =   5250
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1050
         Width           =   4605
      End
      Begin VB.ListBox lstAccountHeads 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3885
         Left            =   30
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1050
         Width           =   4605
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4710
         TabIndex        =   4
         Top             =   1875
         Width           =   450
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4710
         TabIndex        =   3
         Top             =   2475
         Width           =   450
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "(Account Heads)"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   9840
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "  Account Heads"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   30
         TabIndex        =   7
         Top             =   765
         Width           =   4920
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "      Selected Account Heads"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   4995
         TabIndex        =   6
         Top             =   765
         Width           =   4845
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
      ScaleWidth      =   9900
      TabIndex        =   1
      Top             =   0
      Width           =   9900
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   7170
      TabIndex        =   0
      Top             =   5010
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    
    Option Explicit

    Private Sub cmdAdd_Click()
        
        Dim mLoop As Long
        Dim mCount As Long
        mCount = lstAccountHeads.ListCount
        For mLoop = 0 To mCount
            If mLoop = lstAccountHeads.ListCount Then Exit For
            If lstAccountHeads.Selected(mLoop) Then
                lstSelected.AddItem lstAccountHeads.List(mLoop)
                lstSelected.ItemData(lstSelected.NewIndex) = lstAccountHeads.ItemData(mLoop)
                lstAccountHeads.RemoveItem (mLoop)
                mCount = mCount - 1
                mLoop = mLoop - 1
            End If
        Next mLoop
        
    End Sub
    
    Private Sub cmdRemove_Click()
        Dim mLoop As Long
        Dim mCount As Long
        
        mCount = lstSelected.ListCount
        For mLoop = 0 To mCount
            If mLoop = lstSelected.ListCount Then Exit For
            If lstSelected.Selected(mLoop) Then
                lstAccountHeads.AddItem lstSelected.List(mLoop)
                lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = lstSelected.ItemData(mLoop)
                lstSelected.RemoveItem (mLoop)
                mCount = mCount - 1
                mLoop = mLoop - 1
            End If
        Next mLoop
    End Sub

    Private Sub cmdSelect_Click()
        Dim vsAcc       As VSFlexGrid
        Dim mRow        As Long
        Dim mLoop       As Long
        Dim mGridLoop   As Long
        Dim mFoundRow   As Long
        Dim mTmpRow     As Long
        
        If lstSelected.ListCount > 0 Then
            If optAssets.Value Then
                Set vsAcc = frmOpeningBalace.vsGridR
            ElseIf optLiabilities.Value Then
                Set vsAcc = frmOpeningBalace.vsGridL
            End If
            
            For mLoop = 0 To lstSelected.ListCount - 1
                mFoundRow = vsAcc.FindRow(lstSelected.List(mLoop), , 0)
                If mFoundRow = -1 Then
StepBk:
                    mRow = mRow + 1
                    vsAcc.AddItem lstSelected.List(mLoop) & vbTab & vbTab & lstSelected.ItemData(mLoop), mRow
                Else
                    For mGridLoop = 1 To vsAcc.Rows - 1
                        If vsAcc.TextMatrix(mGridLoop, 0) = "" And Val(vsAcc.TextMatrix(mGridLoop, 2)) = 0 Then
                            ' Not Found
                            GoTo StepBk:
                        ElseIf vsAcc.TextMatrix(mGridLoop, 0) = lstSelected.List(mLoop) And Val(vsAcc.TextMatrix(mGridLoop, 2)) = lstSelected.ItemData(mLoop) Then
                            ' Found
                            Exit For
                        End If
                    Next mGridLoop
                End If
            Next mLoop
            
            If optAssets.Value Then
                frmOpeningBalace.RRows = frmOpeningBalace.RRows + mRow
                vsAcc.Select 1, 0, frmOpeningBalace.RRows, 0
            ElseIf optLiabilities.Value Then
                frmOpeningBalace.LRows = frmOpeningBalace.LRows + mRow
                vsAcc.Select 1, 0, frmOpeningBalace.LRows, 0
            End If
            vsAcc.Sort = flexSortStringAscending
            vsAcc.Select 1, 0
            Unload Me
        End If
    End Sub

    Private Sub Form_Activate()
        Me.Top = 2115
        Me.Left = 1800
    End Sub
    
    Private Sub FillHeads()
        lstSelected.Clear
        Dim mSQL As String
        If optLiabilities.Value Then
            mSQL = "Select vchAccountHead, intAccountHeadID From faAccountHeads Where tinType = 3 Order By vchAccountHead"
            Call PopulateList(lstAccountHeads, mSQL, , , True, True)
        Else
            mSQL = "Select vchAccountHead, intAccountHeadID From faAccountHeads Where tinType = 4 Order By vchAccountHead"
            Call PopulateList(lstAccountHeads, mSQL, , , True, True)
        End If
    End Sub
    
    Private Sub lstAccountHeads_DblClick()
        If lstAccountHeads.ListIndex > -1 Then
            lstSelected.AddItem lstAccountHeads.Text
            lstSelected.ItemData(lstSelected.NewIndex) = lstAccountHeads.ItemData(lstAccountHeads.ListIndex)
            lstAccountHeads.RemoveItem (lstAccountHeads.ListIndex)
        End If
    End Sub
    
    Private Sub lstSelected_DblClick()
        If lstSelected.ListIndex > -1 Then
            lstAccountHeads.AddItem lstSelected.Text
            lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = lstSelected.ItemData(lstSelected.ListIndex)
            lstSelected.RemoveItem (lstSelected.ListIndex)
        End If
    End Sub

    Private Sub optAssets_Click()
        Call FillHeads
    End Sub
    
    Private Sub optLiabilities_Click()
        Call FillHeads
    End Sub
