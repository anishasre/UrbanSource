VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChequeBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheque Book"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9870
      TabIndex        =   32
      Top             =   0
      Width           =   9870
   End
   Begin VB.Frame fraBankInfo 
      Height          =   6165
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   9870
      Begin VB.ListBox lstBanks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         ItemData        =   "frmCheckBook.frx":0000
         Left            =   6150
         List            =   "frmCheckBook.frx":0002
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   6660
         TabIndex        =   34
         Top             =   660
         Width           =   345
      End
      Begin VB.TextBox txtBranchCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3330
         TabIndex        =   9
         Top             =   1650
         Width           =   1905
      End
      Begin VB.TextBox txtBankCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3330
         TabIndex        =   5
         Top             =   990
         Width           =   1905
      End
      Begin VB.TextBox txtBankHeadName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   4710
         TabIndex        =   14
         Top             =   2310
         Width           =   3765
      End
      Begin VB.TextBox txtAcNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3330
         TabIndex        =   11
         Top             =   1980
         Width           =   3285
      End
      Begin VB.TextBox txtBranch 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3330
         TabIndex        =   7
         Top             =   1320
         Width           =   3285
      End
      Begin VB.TextBox txtBankName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3330
         TabIndex        =   3
         Top             =   660
         Width           =   3285
      End
      Begin VB.TextBox txtBankHead 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   3330
         TabIndex        =   13
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Frame fraChequeInfo 
         Height          =   3195
         Left            =   30
         TabIndex        =   15
         Top             =   2970
         Width           =   9810
         Begin VB.CommandButton cmdNew 
            Caption         =   "New"
            Height          =   315
            Left            =   2760
            TabIndex        =   29
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtBookNo 
            Height          =   285
            Left            =   3360
            TabIndex        =   18
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancelBook 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   5160
            TabIndex        =   31
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton cmdSaveBook 
            Caption         =   "&Save"
            Height          =   315
            Left            =   3960
            TabIndex        =   30
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   ">>"
            Height          =   315
            Left            =   5520
            TabIndex        =   20
            Top             =   600
            Width           =   585
         End
         Begin VB.TextBox txtSerialLastNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5250
            MaxLength       =   10
            TabIndex        =   28
            Top             =   1650
            Width           =   1725
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "<<"
            Height          =   315
            Left            =   4920
            TabIndex        =   19
            Top             =   600
            Width           =   585
         End
         Begin VB.TextBox txtSerialStartNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   26
            Top             =   1650
            Width           =   1725
         End
         Begin VB.TextBox txtPrefix 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            TabIndex        =   24
            Top             =   1320
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtBookDate 
            Height          =   345
            Left            =   3360
            TabIndex        =   22
            Top             =   930
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-mmm-yyyy"
            Format          =   59179009
            CurrentDate     =   39302
         End
         Begin VB.Label lblBookNo 
            Caption         =   "Book No"
            Height          =   255
            Left            =   2670
            TabIndex        =   17
            Top             =   630
            Width           =   675
         End
         Begin VB.Label lblSeparator 
            Caption         =   "-- "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5100
            TabIndex        =   27
            Top             =   1680
            Width           =   180
         End
         Begin VB.Label lblBookDate 
            AutoSize        =   -1  'True
            Caption         =   "Book Issued Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1770
            TabIndex        =   21
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label lblSerialNo 
            AutoSize        =   -1  'True
            Caption         =   "Serial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2790
            TabIndex        =   25
            Top             =   1650
            Width           =   495
         End
         Begin VB.Label lblPrefix 
            AutoSize        =   -1  'True
            Caption         =   "Prefix"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2820
            TabIndex        =   23
            Top             =   1290
            Width           =   495
         End
         Begin VB.Label lblChequeInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "  &Cheque Info"
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
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   9780
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2145
         TabIndex        =   8
         Top             =   1620
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2310
         TabIndex        =   4
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblBankInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "  & Bank Info"
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
         Height          =   255
         Left            =   15
         TabIndex        =   1
         Top             =   120
         Width           =   9825
      End
      Begin VB.Label lblBranch 
         AutoSize        =   -1  'True
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2655
         TabIndex        =   6
         Top             =   1290
         Width           =   615
      End
      Begin VB.Label lblBankName 
         AutoSize        =   -1  'True
         Caption         =   "Name of Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   2
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label lblAcNo 
         AutoSize        =   -1  'True
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1815
         TabIndex        =   10
         Top             =   1950
         Width           =   1455
      End
      Begin VB.Label lblBankHead 
         AutoSize        =   -1  'True
         Caption         =   "Bank Account Head"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1530
         TabIndex        =   12
         Top             =   2280
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmChequeBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        Option Explicit
        
            Dim ChequeBookCnt As Integer
            Dim ChequeLeafCnt As Integer
            Dim arrOutBookCnt As Variant
            Dim arrOutLeafCnt As Variant
            Dim mEditFlag     As Boolean
           
        
        Private Sub FormInitialize()
            Set arrOutBookCnt = Nothing
            ChequeBookCnt = -1
            txtBankHead.Text = ""
            txtBankHeadName.Text = ""
            txtBankName.Text = ""
            txtBranch.Text = ""
            txtBranchCode.Text = ""
            txtBankCode.Text = ""
            txtAcNo.Text = ""
            txtBookNo.Text = ""
            Call ClearChequeBookForm
        End Sub
        
        Private Sub DisplayCheckBook(intBookID As Long)
            Dim ObjDb       As New clsDB
            Dim arrInput    As Variant
            Dim Rec         As New ADODB.Recordset
                       
            arrInput = Array(intBookID)
            Set Rec = ObjDb.ExecuteSP("spGetChequeBookInfo", arrInput)
            If Not (Rec.BOF And Rec.EOF) Then
                dtBookDate.value = Rec!dtissuedDate
                txtBookNo.Tag = Rec!intChequeBookID
                txtBookNo.Text = Rec!intBookNo
                txtPrefix.Text = IIf(IsNull(Rec!vchPrefix), "", Rec!vchPrefix)
                txtSerialStartNo.Text = Rec!intStartingNo
                txtSerialLastNo.Text = Rec!intEndingNo
            Else
                txtBookNo.Tag = ""
                txtBookNo.Text = ""
                txtPrefix.Text = ""
                txtSerialStartNo.Text = ""
                txtSerialLastNo.Text = ""
                dtBookDate.value = Date
            End If
            Rec.Close
        End Sub

        Private Sub cmdCancelBook_Click()
           Unload Me
        End Sub
        
        Private Sub cmdClear_Click()
             Call FormInitialize
        End Sub

        Private Sub cmdNew_Click()
            Call ClearChequeBookForm
            Call ClearBankDetails
            txtBookNo.SetFocus
        End Sub

        Private Sub cmdNext_Click()
            If Not IsArray(arrOutBookCnt) Then
                Exit Sub
            End If
            ChequeBookCnt = ChequeBookCnt + 1
            If ChequeBookCnt <= UBound(arrOutBookCnt, 2) And ChequeBookCnt > -1 Then
STEP1:
                dtBookDate.value = arrOutBookCnt(5, ChequeBookCnt)
                txtBookNo.Text = arrOutBookCnt(1, ChequeBookCnt)
                txtBookNo.Tag = val(arrOutBookCnt(0, ChequeBookCnt))
                txtPrefix.Text = IIf(IsNull((arrOutBookCnt(2, ChequeBookCnt))), "", (arrOutBookCnt(2, ChequeBookCnt))) 'arrOutBookCnt(2, ChequeBookCnt)
                txtSerialStartNo.Text = arrOutBookCnt(3, ChequeBookCnt)
                txtSerialLastNo.Text = arrOutBookCnt(4, ChequeBookCnt)
            Else
                ChequeBookCnt = LBound(arrOutBookCnt, 2)
                cmdPrevious.Enabled = True
                If ChequeBookCnt > -1 Then GoTo STEP1
            End If
        End Sub
        
        Private Sub cmdPrevious_Click()
            If Not IsArray(arrOutBookCnt) Then
                Exit Sub
            End If
            ChequeBookCnt = ChequeBookCnt - 1
            If ChequeBookCnt <= UBound(arrOutBookCnt, 2) And ChequeBookCnt > -1 Then
STEP1:
                dtBookDate.value = arrOutBookCnt(5, ChequeBookCnt)
                txtBookNo.Text = arrOutBookCnt(1, ChequeBookCnt)
                txtBookNo.Tag = val(arrOutBookCnt(0, ChequeBookCnt))
                txtPrefix.Text = IIf(IsNull((arrOutBookCnt(2, ChequeBookCnt))), "", (arrOutBookCnt(2, ChequeBookCnt))) 'arrOutBookCnt(2, ChequeBookCnt)
                txtSerialStartNo.Text = arrOutBookCnt(3, ChequeBookCnt)
                txtSerialLastNo.Text = arrOutBookCnt(4, ChequeBookCnt)
            Else
                ChequeBookCnt = UBound(arrOutBookCnt, 2)
                cmdNext.Enabled = True
                If ChequeBookCnt > -1 Then
                    GoTo STEP1
                End If
            End If
        End Sub
        
       Private Sub dtBookDate_LostFocus()
                If mID$(dtBookDate.value, 7, 4) > gbFinancialYearID Then
                    MsgBox "Book Issued Date is greater than Current Financial year", vbInformation
                ElseIf dtBookDate.value > gbTransactionDate Then
                    MsgBox "Book Issued Date is greater than Current date", vbInformation
                End If
        End Sub

        Private Sub Form_Activate()
                ChequeBookCnt = 1
                Me.Top = 0
                frmChequeBook.Left = (frmMenu.Width - Me.Width) / 2
        End Sub
        
        Private Sub cmdSaveBook_Click()
                Dim objAcc  As New clsAccounts
                Dim ObjDb   As New clsDB
                Dim objBk   As New clsBank
                '----------------------------------------------------'
                ' Validations
                '----------------------------------------------------'
                objAcc.SetAccountCode ((Trim(txtBankHead.Text)))
                If objAcc.AccountHeadID < 0 Then
                    MsgBox "Select a Cash or Bank Account Head!", vbInformation
                    'txtBankHead.SetFocus
                    Exit Sub
                End If

                objBk.SetBankInfo (val(txtBankName.Tag))
                If objBk.BankID < 0 Then
                    txtBankName.SetFocus
                    Exit Sub
                End If
                
                If txtSerialStartNo.Text = "" Or txtSerialLastNo.Text = "" Or txtBookNo.Text = "" Then 'txtPrefix.Text = "" Or
                    MsgBox "Enter ChequeBook Details!", vbInformation
                    txtBookNo.SetFocus
                    Exit Sub
                End If
                
                If val(txtBookNo.Text) = -1 Then
                    MsgBox "Please Check the Book Number Entered", vbInformation
                    txtBookNo.SetFocus
                End If
                
                If val(txtSerialStartNo.Text) = -1 Then
                    MsgBox "Please Check the Serial Number Entered", vbInformation
                    txtSerialStartNo.SetFocus
                End If
                
                If val(txtSerialLastNo.Text) = -1 Then
                    MsgBox "Please Check the Serial Number Entered", vbInformation
                    txtSerialLastNo.SetFocus
                End If
                
                If Trim(txtSerialLastNo.Text) < Trim(txtSerialStartNo.Text) Then
                    MsgBox "Serial End Number should be Greater than the Start Number", vbInformation
                    txtSerialStartNo.SetFocus
                End If
                '-------------------------------------------------'
                ' Saving Data
                '-------------------------------------------------'
                Dim arrInput            As Variant
                Dim mCon                As ADODB.Connection
                Dim mCom                As ADODB.Command
                Dim Rec                 As ADODB.Recordset
                Dim RecBook             As ADODB.Recordset
                Dim RecBookUpdate       As ADODB.Recordset
                Dim mAccountID          As Double
                Dim mBankID             As Long
                Dim mCurrentBookFlag    As Integer
                Dim msQl                As String
                
                Set Rec = New ADODB.Recordset
                Set RecBook = New ADODB.Recordset
                Set RecBookUpdate = New ADODB.Recordset
                Set mCom = New ADODB.Command

                  If objBk.BankID > 0 Then
                         mBankID = objBk.BankID
                         mAccountID = objAcc.AccountHeadID
                        ObjDb.SetConnection mCon
                        '--------------------------------------------'
                        ' Creating a new Cheque Book
                        '--------------------------------------------'
                        If val(txtBookNo.Tag) < 1 Then
                            arrInput = Array(-1, _
                                                mAccountID, _
                                                ((Trim(txtAcNo.Text))), _
                                                val(Trim(txtBookNo.Text)), _
                                                Trim(txtPrefix.Text), _
                                                val(Trim(txtSerialStartNo.Text)), _
                                                val(Trim(txtSerialLastNo.Text)), _
                                                Format(dtBookDate.value, "dd/mmm/yyyy"), _
                                                mCurrentBookFlag, _
                                                mBankID)
                            ObjDb.ExecuteSP "spSaveNewChequeBook", arrInput
                            
                            arrInput = Array((objBk.BankID))
                            Call ObjDb.ExecuteSP("spGetPreviousChequeBooks", arrInput, arrOutBookCnt)
                            '--------------------------------------------'
                            ' Updating an existing Cheque Book
                            '--------------------------------------------'
                        Else
                            arrInput = Array(val(txtBookNo.Tag), _
                                                val(Trim(txtBookNo.Text)), _
                                                Trim(txtPrefix.Text), _
                                                val(Trim(txtSerialStartNo.Text)), _
                                                val(Trim(txtSerialLastNo.Text)), _
                                                Format(dtBookDate.value, "dd/mmm/yyyy"))
                            ObjDb.ExecuteSP "spUpdateChequeBook", arrInput
                            arrInput = Array((objBk.BankID))
                            Call ObjDb.ExecuteSP("spGetPreviousChequeBooks", arrInput, arrOutBookCnt)
                        End If
                        Call ClearChequeBookForm
                    End If
                End Sub
        





            Private Sub txtBankHead_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then PressTabKey
            End Sub
             
            Private Sub txtBankName_GotFocus()
                Call txtBankName_LostFocus
            End Sub

            Private Sub txtBankName_LostFocus()
                    Dim ObjDb               As New clsDB
                    Dim arrInput            As Variant
                    Dim Rec                 As New ADODB.Recordset
                    Dim objAcc              As New clsAccounts
                    Dim objBk               As New clsBank
                    Dim arInput             As Variant
                    Dim RecPrevChequeBook   As ADODB.Recordset
                    Dim msQl                As String
                    Dim RecChequeBook       As ADODB.Recordset
                    Dim RecChequeLeaf       As ADODB.Recordset 'Checkleaf recordset
                    Dim mLoop               As Long
                    Dim objCheque           As New clsChequeBook

                    objBk.SetBankInfo (val(txtBankName.Tag))
                    If objBk.BankID > -1 Then
                        arrInput = Array(objBk.BankID)
                        Set Rec = ObjDb.ExecuteSP("spGetBankInfo", arrInput)

                        If Not (Rec.BOF And Rec.EOF) Then
                            txtBankName.Text = Rec!vchBankName
                            txtBankName.Tag = Rec!intBankID
                            txtBranch.Text = IIf(IsNull(Rec!vchBranch), "", Rec!vchBranch)
                            txtAcNo.Text = Rec!vchAccountNumber
                            Set arrOutBookCnt = Nothing
                            Call ClearChequeBookForm

                            arInput = Array((objBk.BankID))
                            Call ObjDb.ExecuteSP("spGetPreviousChequeBooks", arInput, arrOutBookCnt)
                            msQl = "Select * From faChequeBook WHERE intBankID  = " & objBk.BankID & " AND tinCurrentBookFlag = 1"
                            Set RecChequeBook = GetRecordSet(msQl)
                             If Not (RecChequeBook.EOF Or RecChequeBook.BOF) Then
                                dtBookDate.value = RecChequeBook!dtissuedDate
                                txtBookNo.Tag = RecChequeBook!intChequeBookID
                                txtBookNo.Text = IIf(IsNull(RecChequeBook!intBookNo), "", RecChequeBook!intBookNo)
                                txtPrefix.Text = IIf(IsNull(RecChequeBook!vchPrefix), "", RecChequeBook!vchPrefix)
                                txtSerialStartNo.Text = RecChequeBook!intStartingNo
                                txtSerialLastNo.Text = RecChequeBook!intEndingNo
                                For mLoop = LBound(arrOutBookCnt, 2) To UBound(arrOutBookCnt, 2)
                                    If arrOutBookCnt(0, mLoop) = RecChequeBook!intChequeBookID Then
                                        ChequeBookCnt = mLoop
                                    End If
                                Next mLoop
                            End If
                            RecChequeBook.Close
                    Else
                        txtBankName.Text = ""
                        txtBranch.Text = ""
                        txtAcNo.Text = ""

                        dtBookDate.value = Date
                        txtBookNo.Tag = ""
                        txtBookNo.Text = ""
                        txtPrefix.Text = ""
                        txtSerialStartNo.Text = ""
                        txtSerialLastNo.Text = ""
                    End If
                    End If
                End Sub
        
                Private Function ClearChequeBookForm()
                    dtBookDate.value = gbTransactionDate
                    txtPrefix.Text = ""
                    txtSerialStartNo.Text = ""
                    txtSerialLastNo.Text = ""
                    txtBookNo.Text = ""
                    txtBookNo.Tag = ""
                End Function
                Private Function ClearBankDetails()
                    txtBankName.Text = ""
                    txtBankCode.Text = ""
                    txtBranch.Text = ""
                    txtBranchCode.Text = ""
                    txtAcNo.Text = ""
                    txtBankHead.Text = ""
                    txtBankHeadName.Text = ""
                End Function
                
                Private Sub txtBookNo_LostFocus()
                    Dim objAcc      As New clsAccounts
                    Dim ObjDb       As New clsDB
                    Dim objBk       As New clsBank
                    Dim mCon        As ADODB.Connection
                    Dim Rec         As ADODB.Recordset
                    Dim mBankID     As Long
                    Dim msQl        As String
                    Dim temp        As Double
                                                                    
                   
                                                                    
                    Set Rec = New ADODB.Recordset
                    objBk.SetBankInfo (val(txtBankName.Tag))
                     If objBk.BankID > -1 Then
                        mBankID = objBk.BankID
                        
                        ObjDb.SetConnection mCon
                        Rec.Open "Select * from faChequeBook where faChequeBook.intBookNo = " & val(txtBookNo.Text) & " AND faChequeBook.intBankID  = " & mBankID, mCon
                            If Not (Rec.EOF Or Rec.BOF) Then
                                dtBookDate.value = Rec!dtissuedDate
                                txtBookNo.Tag = Rec!intChequeBookID
                                txtBookNo.Text = Rec!intBookNo
                                txtPrefix.Text = IIf(IsNull(Rec!vchPrefix), "", Rec!vchPrefix)
                                txtSerialStartNo.Text = Rec!intStartingNo
                                txtSerialLastNo.Text = Rec!intEndingNo
                            End If
                    End If
                End Sub

                Private Sub DisplayBank(mID As Double)
                    Dim objAcc          As New clsAccounts
                    Dim objBk           As New clsBank
                    Dim objchequeBook   As New clsChequeBook
            
                    objBk.SetBankInfo mID
                    If objBk.BankID > -1 Then
                        mEditFlag = True
                        objAcc.SetAccountID objBk.BankAccountHeadID
                        objchequeBook.SetChequeBookInfo objBk.BankID
                        
                        txtBankName.Text = objBk.BankName
                        txtBankName.Tag = objBk.BankID
                        txtBankCode.Text = objBk.BankCode
                        txtBranch.Text = objBk.Branch
                        txtBranchCode.Text = objBk.BranchCode
                        txtAcNo.Text = objBk.AccountNumber
                        txtBankHead.Text = objAcc.AccountCode
                        txtBankHeadName.Text = objAcc.AccountHead
                        txtBankHeadName.Tag = objAcc.AccountHeadID
                        txtBookNo = objchequeBook.BookNo
                        dtBookDate = objchequeBook.IssuedDate
                        txtPrefix = objchequeBook.Prefix
                        txtSerialStartNo = objchequeBook.StartingNo
                        txtSerialLastNo = objchequeBook.EndingNo
                    End If
                    Set objBk = Nothing
                End Sub
                Private Sub cmdSearch_Click()
                    Call PopulateList(lstBanks, "Select vchBankName, intBankID from faBanks Order By vchBankName", , , , True)
                    lstBanks.Visible = True
                    lstBanks.SetFocus
                End Sub
    
                Private Sub lstBanks_DblClick()
                    Call lstBanks_KeyDown(13, 0)
                End Sub
            
                Private Sub lstBanks_KeyDown(KeyCode As Integer, Shift As Integer)
                    If KeyCode = 13 Then
                        gbSearchStr = lstBanks.Text
                        gbSearchID = lstBanks.ItemData(lstBanks.ListIndex)
                        Call DisplayBank(gbSearchID)
                        lstBanks.Visible = False
                    End If
                End Sub
                        
                Private Sub lstBanks_LostFocus()
                    lstBanks.Visible = False
                End Sub
                Private Sub txtBookNo_GotFocus()
                    If gbSearchStr <> "" Then
                        Dim mStr As String
                        txtBankName.Text = Trim(gbSearchStr)
                        txtBankName.Tag = gbSearchID
                        gbSearchStr = ""
                        gbSearchID = -1
                    End If
                    txtBankName.SelStart = 0
                    txtBankName.SelLength = Len(txtBankName)
                End Sub

                Private Sub txtSerialLastNo_KeyPress(KeyAscii As Integer)
                    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                        KeyAscii = 0
                    End If
                End Sub
                Private Sub txtSerialLastNo_LostFocus()
                    Dim ObjDb       As New clsDB
                    Dim mCon        As ADODB.Connection
                    Dim Rec         As ADODB.Recordset
                    Dim msQl        As String
                    
                  
                    Set Rec = New ADODB.Recordset
                    ObjDb.SetConnection mCon
                       If txtSerialLastNo.Text <> "" Then
                            If txtBankName.Text <> "" Then
                                'msQl = "SELECT * From faChequeBook WHERE " & Trim(txtSerialLastNo.Text) & " BETWEEN intStartingNo AND intEndingNo and intBankID=" & Trim(txtBankName.Tag) & ""
                                msQl = "SELECT * From faChequeBook WHERE intStartingNo <= " & Trim(txtSerialLastNo.Text) & "  And intEndingNo>=" & Trim(txtSerialLastNo.Text) & " and intBankID=" & Trim(txtBankName.Tag) & ""
                                Rec.Open msQl, mCon
                                If Not (Rec.EOF Or Rec.BOF) Then
                                    MsgBox "Serial Number Already Entered", vbInformation
                                    txtSerialStartNo.SetFocus
                                End If
                            End If
                        End If
                End Sub

                Private Sub txtSerialStartNo_KeyPress(KeyAscii As Integer)
                    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                        KeyAscii = 0
                    End If
                End Sub

                Private Sub txtSerialStartNo_LostFocus()
                    Dim ObjDb       As New clsDB
                    Dim mCon        As ADODB.Connection
                    Dim Rec         As ADODB.Recordset
                    Dim msQl        As String
                    
                    Set Rec = New ADODB.Recordset
                                                                                       
                    ObjDb.SetConnection mCon
                    
                    If txtSerialStartNo.Text <> "" Then
                        If txtBankName.Text <> "" Then
                            'msQl = "SELECT * From faChequeBook WHERE " & Trim(txtSerialStartNo.Text) & " BETWEEN intStartingNo AND intEndingNo and intBankID=" & Trim(txtBankName.Tag) & ""
                            msQl = "SELECT * From faChequeBook WHERE intStartingNo <= " & Trim(txtSerialStartNo.Text) & "  And intEndingNo>=" & Trim(txtSerialStartNo.Text) & " and intBankID=" & Trim(txtBankName.Tag) & ""
                            Rec.Open msQl, mCon
                            If Not (Rec.EOF Or Rec.BOF) Then
                                MsgBox "Serial Number Already Entered", vbInformation
                                txtSerialStartNo.SetFocus
                            End If
                        End If
                    End If
                End Sub















