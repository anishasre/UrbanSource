VERSION 5.00
Begin VB.Form frmDatabaseConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONNECT TO DATABASE "
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdConnect 
      Caption         =   "CONNECT"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   5055
      Begin VB.TextBox txtDBName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtServerName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cmbLBName 
         Height          =   390
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmDatabaseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub cmbLBName_Click()
        Call FillDB
    End Sub

    Private Sub cmdConnect_Click()
         Dim mCnnServer  As New ADODB.Connection
         Dim mCnnDB  As New ADODB.Connection
         Call ServerConnection(mCnnServer)
         Call DatabaseConnection(mCnnDB)
    End Sub

    Private Sub Form_Load()
        Call FillCombo
        txtServerName.Text = "DB_Finance"
    End Sub
    
    Private Sub ServerConnection(mCnnServer As ADODB.Connection)
        mCnnServer.ConnectionString = "PROVIDER=MSDASQL;dsn=dsnFa;uid=FAUser;pwd=FAUser;database=" + Trim(txtServerName.Text) + ";"
        mCnnServer.Open
    End Sub
    
    Private Sub DatabaseConnection(mCnnClient As ADODB.Connection)
        mCnnClient.ConnectionString = "PROVIDER=MSDASQL;dsn=DSNFinance_01;uid=FAUser;pwd=FAUser;database=" + Trim(txtDBName.Text) + ";"
        mCnnClient.Open
    End Sub
    Private Sub FillDB()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL  As String
        Dim Rec   As New ADODB.Recordset
        
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSQL = "SELECT *  FROM tmpMergedLBs WHERE intLBID=" & (val(cmbLBName.ItemData(cmbLBName.ListIndex)))
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                txtDBName.Text = Rec!vchDBName
                txtDBName.Enabled = False
            End If
            Rec.Close
        End If
        
        
    End Sub
    Private Sub FillCombo()
        Dim mSQL As String
           
        mSQL = "SELECT vchLBName,intLBID FROM tmpMergedLBs"
        PopulateList cmbLBName, mSQL, , , True, True, enuSourceString.Saankhya

    End Sub
        
