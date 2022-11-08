VERSION 5.00
Begin VB.Form frmDCBReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View DCB Report"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   Icon            =   "frmDCBReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbMonth 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1035
      TabIndex        =   2
      Top             =   1485
      Width           =   2055
   End
   Begin VB.ComboBox cmbFinancialYear 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblItem 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblFinancialyear 
      Caption         =   "Financial Year"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "frmDCBReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub FillCombo()
'''        Dim mCnn  As New ADODB.Connection
'''        Dim objDb   As New clsDB
'''        Dim mSql    As String
'''        Dim Rec     As New ADODB.Recordset
'''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'''        mSql = "Select  intFinancialYear,intFinancialYearID from faFinancialYear"
'''        PopulateList cmbFinancialYear, mSql, , True, True, True, enuSourceString.Saankhya
        
        cmbFinancialYear.AddItem 2013
        cmbFinancialYear.ItemData(cmbFinancialYear.NewIndex) = 2013
        cmbFinancialYear.AddItem 2014
        cmbFinancialYear.ItemData(cmbFinancialYear.NewIndex) = 2014
                
    End Sub
    
    Private Sub cmbFinancialYear_Click()
        cmbMonth.Enabled = True
        'Call ExtractDCB
    End Sub

    Private Sub cmbMonth_Click()
       'Call ExtractDCB
    End Sub

    Private Sub cmdReport_Click()
        Dim objDb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        Dim mYearID As Integer
        If cmbFinancialYear.ListIndex > -1 Then
            mYearID = cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)
        Else
            mYearID = gbFinancialYearID
        End If
        
        If cmbMonth.ListIndex < 0 Then
            MsgBox "Please select the Item", vbInformation
            cmbMonth.SetFocus
            Exit Sub
        End If
        
        arInput = Array(mYearID, cmbMonth.ItemData(cmbMonth.ListIndex))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptDCB.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
        Set frmNewViewer = Nothing
        Set frmNewRpt = Nothing

    End Sub
    
    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
    End Sub
    
    Private Sub Form_Load()
        Call FillCombo
        Call fillMonthCombo
    End Sub
    
    
    Private Function ExtractDCB()
        Dim mCnn    As New ADODB.Connection
        Dim mSql    As String
        Dim objDb   As New clsDB
        Dim arInput As Variant
        
        arInput = Array(cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex), cmbMonth.ItemData(cmbMonth.ListIndex))
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCnn.Execute " Delete  from faMonthlyDCB Where intYearID =" & val(cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)) & " And intMonthID = " & cmbMonth.ItemData(cmbMonth.NewIndex)
        objDb.ExecuteSP "spExtactDCBHeadWise", arInput, , , mCnn, adCmdStoredProc
        mCnn.Close
        
End Function
Private Sub fillMonthCombo()
        cmbMonth.AddItem "April"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 4
        cmbMonth.AddItem "May"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 5
        cmbMonth.AddItem "June"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 6
        cmbMonth.AddItem "July"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 7
        cmbMonth.AddItem "August"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 8
        cmbMonth.AddItem "September"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 9
        cmbMonth.AddItem "October"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 10
        cmbMonth.AddItem "November"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 11
        cmbMonth.AddItem "December"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 12
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
    End Sub
