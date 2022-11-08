VERSION 5.00
Begin VB.Form frmGreetings 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   2100
      Left            =   15
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   5475
      Left            =   -465
      Picture         =   "frmGreetings.frx":0000
      Stretch         =   -1  'True
      Top             =   -840
      Width           =   9465
   End
End
Attribute VB_Name = "frmGreetings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_Defaut         As Long = &H2
Private Const WS_EX_LAYERED     As Long = &H80000

Private Sub Form_Click()
    Unload Me
    frmSplashForPanchayat.Show vbModal
End Sub

Private Sub ActiveTransparency(M As Form, d As Boolean, F As Boolean, _
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
        If Not Err.Number = 0 Then Err.Clear
    End Function

Private Sub Form_Load()
    Dim i As Integer
    ActiveTransparency Me, True, False, 0
        Me.Show
        For i = 0 To 255 Step 3
            ActiveTransparency Me, True, False, i
            Me.Refresh
        Next i
End Sub

Private Sub Image1_Click()
    On Error Resume Next
    Unload Me
    frmSplashForPanchayat.Show vbModal
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Unload Me
    frmSplashForPanchayat.Show vbModal
End Sub
