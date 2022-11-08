VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStartUp.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   8850
      Top             =   4560
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   5355
      TabIndex        =   0
      Top             =   1530
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   8910
      Picture         =   "frmStartUp.frx":18E15
      Top             =   4050
      Visible         =   0   'False
      Width           =   9030
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim mLBTypeID As Variant
'''''''''''''''''''' For making the Background of the form as Transparent ''''''''''''''''''''''
'    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
'                ByVal hwnd As Long, _
'                ByVal nIndex As Long) As Long
'
'    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
'                    ByVal hwnd As Long, _
'                    ByVal nIndex As Long, _*
'                    ByVal dwNewLong As Long) As Long
'
'    Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
'                    ByVal hwnd As Long, _
'                    ByVal crKey As Long, _
'                    ByVal bAlpha As Byte, _
'                    ByVal dwFlags As Long) As Long
'
'    Private Const GWL_STYLE = (-16)
'    Private Const GWL_EXSTYLE = (-20)
'    Private Const WS_EX_LAYERED = &H80000
'    Private Const LWA_COLORKEY = &H1
'    Private Const LWA_ALPHA = &H2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''' '''''''''''''For load the form in Fade mode & Make the form Transparent ''''''''''''''''''''''''''''''''''''''''

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_Defaut         As Long = &H2
Private Const WS_EX_LAYERED     As Long = &H80000
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''' For round corners of the Form ''''''''''''''''''''''''''
Private Declare Function CreateRoundRectRgn _
    Lib "gdi32" (ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal X3 As Long, _
    ByVal Y3 As Long) As Long
    
    Private Declare Function SetWindowRgn _
    Lib "user32" (ByVal hwnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''' For round corners of the Form '''''''''''''''''''''''''''''''
    Public Sub CreateRoundRectFromWindow(ByRef oWindow As Object)
    
        Dim lRight As Long
        Dim lBottom As Long
        Dim hRgn As Long
        
        With oWindow
            lRight = .Width / Screen.TwipsPerPixelX
            lBottom = .Height / Screen.TwipsPerPixelY
            hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, 40, 40)
            SetWindowRgn .hwnd, hRgn, True
        End With
    End Sub
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Unload Me
'        If mLBTypeID = 3 Or mLBTypeID = 4 Then
'            frmSplash.Show vbModal
'        Else
'            frmSplashForPanchayat.Show vbModal
'        End If
        frmSplashForPanchayat.Show vbModal
    End Sub
    
    Private Sub Form_Load()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        
        On Error GoTo err
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select tnyLBTypeID From faLBSettings"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mLBTypeID = IIf(IsNull(Rec!tnyLBTypeID), "", Rec!tnyLBTypeID)
                If mLBTypeID = 3 Or mLBTypeID = 4 Then
                    lblCompanyProduct.Caption = "Kerala Municipal Accounting System"
                    lblCompanyProduct.Left = 5355
                Else
                    lblCompanyProduct.Caption = "Kerala Panchayat Raj Accounting System"
                    lblCompanyProduct.Left = 4995
                End If
            End If
            Rec.Close
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        CreateRoundRectFromWindow Me
        'Timer1_Timer
               
        '''''''''''''''''''' '''''''''''''For load the form in Fade mode '''''''''''''''''''''''''''''''
        Dim i As Integer
        'Ex: all transparent at ratio 140/255
        'ActiveTransparency Me, True, False, 140, Me.BackColor
        'Ex: Form transparent, visible component at ratio 140/255
        'ActiveTransparency Me, True, True, 140, Me.BackColor
         
        'Example display the form transparency degradation
'        ActiveTransparency Me, True, False, 0
'        Me.Show
'        For i = 0 To 255 Step 3
'            ActiveTransparency Me, True, False, i
'            Me.Refresh
'        Next i
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''' For making the Background of the form as Transparent ''''''''''''''''''''''
'                'Me.BackColor = vbCyan
'                SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
'                SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        On Error GoTo lblSkip:
        Unload Me

        'If Date > "20-Aug-2012" And Date < "5-Sep-2012" Then
        '    frmGreetings.Show vbModal
        'Else
        '    frmSplashForPanchayat.Show vbModal
        'End If
lblSkip:

        frmSplashForPanchayat.Show vbModal
    End Sub

    ''''''''''''''''''' '''''''''''''For load the form in Fade mode '''''''''''''''''''''''''''''''
    Public Function Transparency(ByVal hwnd As Long, Optional ByVal Col As Long = vbBlack, _
        Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
    ' Return : True if there is no error.
    ' hWnd   : hWnd of the window to make transparent
    ' Col : Color to make transparent if TrMode=False
    ' PcTransp  : 0 Ã  255 >> 0 = transparent  -:- 255 = Opaque
    Dim DisplayStyle As Long
        On Error GoTo Ex
        VoirStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
            DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
            Call SetWindowLong(hwnd, GWL_EXSTYLE, DisplayStyle)
        End If
        Transparency = (SetLayeredWindowAttributes(hwnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)
         
Ex:
        If Not err.Number = 0 Then err.Clear
    End Function
    
    Public Sub ActiveTransparency(M As Form, d As Boolean, F As Boolean, _
        T_Transparency As Integer, Optional Color As Long)
        Dim b As Boolean
            If d And F Then
            'Makes color (here the background color of the shape) transparent
            'upon value of T_Transparency
                b = Transparency(M.hwnd, Color, T_Transparency, False)
            ElseIf d Then
                'Makes form, including all components, transparent
                'upon value of T_Transparency
                b = Transparency(M.hwnd, 0, T_Transparency, True)
            Else
                'Restores the form opaque.
                b = Transparency(M.hwnd, , 255, True)
            End If
    End Sub

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub Timer1_Timer()
        On Error GoTo lblExitSub:
        Unload Me

        'If Date > "20-Aug-2012" And Date < "5-Aug-2012" Then
        '    frmGreetings.Show vbModal
        'Else
            frmSplashForPanchayat.Show vbModal
        'End If
lblExitSub:
    End Sub
