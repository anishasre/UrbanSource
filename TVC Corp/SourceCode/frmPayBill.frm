VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPayBill 
   BackColor       =   &H00F2FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5475
      Left            =   8865
      TabIndex        =   9
      Top             =   2940
      Width           =   9420
      Begin VB.ComboBox cmbMonth 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2295
         TabIndex        =   11
         Text            =   "cmbMonth"
         Top             =   540
         Width           =   1560
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         TabIndex        =   10
         Text            =   "2009"
         Top             =   555
         Width           =   615
      End
      Begin VSFlex8LCtl.VSFlexGrid GridAbstract 
         Height          =   4065
         Left            =   345
         TabIndex        =   12
         Top             =   990
         Width           =   8625
         _cx             =   15214
         _cy             =   7170
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16053492
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPayBill.frx":0000
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
         FillStyle       =   1
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1785
         TabIndex        =   14
         Top             =   570
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   13
         Top             =   585
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6300
      Left            =   60
      TabIndex        =   2
      Top             =   -45
      Width           =   9345
      Begin VB.CommandButton cmdSeat 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         Height          =   285
         Left            =   5595
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5940
         Width           =   315
      End
      Begin VB.TextBox txtForward2Seat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5955
         Width           =   1725
      End
      Begin VB.TextBox txtGrossAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2025
         TabIndex        =   35
         Top             =   5250
         Width           =   1110
      End
      Begin VB.TextBox txtTotalDeduction 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4515
         TabIndex        =   34
         Top             =   5250
         Width           =   1110
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   105
         TabIndex        =   15
         Top             =   105
         Width           =   9135
         Begin VB.TextBox txtSourceOfFund 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   47
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CommandButton cmdSearchSourceOfFund 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1920
            Width           =   315
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   45
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox txtBillNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4410
            TabIndex        =   41
            Top             =   840
            Width           =   585
         End
         Begin VB.TextBox txtPayOrderNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   1305
            TabIndex        =   37
            Top             =   180
            Width           =   3105
         End
         Begin VB.TextBox txtSection 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1305
            TabIndex        =   23
            Top             =   840
            Width           =   3105
         End
         Begin VB.TextBox txtMonth 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2805
            TabIndex        =   22
            Text            =   "2009"
            Top             =   510
            Width           =   1605
         End
         Begin VB.TextBox txtYearID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1305
            TabIndex        =   21
            Text            =   "2009"
            Top             =   510
            Width           =   915
         End
         Begin VB.TextBox txtDueDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6960
            TabIndex        =   20
            Text            =   "1-Apr-2009"
            Top             =   585
            Width           =   1260
         End
         Begin VB.TextBox txtID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6960
            TabIndex        =   19
            Text            =   "100001"
            Top             =   255
            Width           =   1260
         End
         Begin VB.TextBox txtFunctionary 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1305
            TabIndex        =   18
            Top             =   1185
            Width           =   3105
         End
         Begin VB.TextBox txtFunction 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5130
            TabIndex        =   17
            Top             =   1185
            Width           =   3105
         End
         Begin VB.ComboBox cmbAccountHeads 
            Height          =   315
            Left            =   1305
            TabIndex        =   16
            Text            =   "Combo1"
            Top             =   1500
            Width           =   7185
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Of Fund"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Name Of Payee"
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
            Height          =   255
            Left            =   4920
            TabIndex        =   48
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PayOrder No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   345
            TabIndex        =   38
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   915
            TabIndex        =   31
            Top             =   540
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2370
            TabIndex        =   30
            Top             =   540
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6270
            TabIndex        =   29
            Top             =   615
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6795
            TabIndex        =   28
            Top             =   285
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   720
            TabIndex        =   27
            Top             =   885
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Functionary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   26
            Top             =   1215
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Function"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4485
            TabIndex        =   25
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary A/c Head"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   24
            Top             =   1560
            Width           =   1185
         End
      End
      Begin VB.TextBox txtPensionContribution 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1425
         TabIndex        =   6
         Text            =   "Pension Contribution"
         Top             =   5640
         Width           =   3735
      End
      Begin VB.TextBox txtPensionCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5175
         TabIndex        =   5
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox txtNetSalary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6885
         TabIndex        =   4
         Top             =   5250
         Width           =   1230
      End
      Begin VB.TextBox txtAmtPension 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6900
         TabIndex        =   3
         Top             =   5640
         Width           =   1215
      End
      Begin VSFlex8LCtl.VSFlexGrid Grid 
         Height          =   2625
         Left            =   840
         TabIndex        =   7
         Top             =   2580
         Width           =   7605
         _cx             =   13414
         _cy             =   4630
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16053492
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
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPayBill.frx":0142
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
         FillStyle       =   1
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded to Seat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2445
         TabIndex        =   44
         Top             =   6000
         Width           =   1365
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press Escape Key to Go Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   6030
         TabIndex        =   40
         Top             =   6000
         Width           =   2700
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   945
         TabIndex        =   36
         Top             =   5280
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deduction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3390
         TabIndex        =   33
         Top             =   5280
         Width           =   1110
      End
      Begin VB.Label lblPaymentOrderNo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00FF8080&
         Height          =   210
         Left            =   165
         TabIndex        =   32
         Top             =   5985
         Width           =   1680
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Salary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6105
         TabIndex        =   8
         Top             =   5280
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   5235
         Width           =   7590
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   570
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   9390
      TabIndex        =   0
      Top             =   6180
      Width           =   9450
      Begin WinXPC_Engine.WindowsXPC WindowsXPC 
         Left            =   7440
         Top             =   135
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   45
         Top             =   45
      End
      Begin VB.CommandButton cmdPaymentOrder 
         Caption         =   "Payment Order"
         Height          =   390
         Left            =   3480
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   4200
      TabIndex        =   39
      Top             =   5835
      Width           =   60
   End
End
Attribute VB_Name = "frmPayBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mTimeDelay As Integer
    Dim mSalaryAccountHeadCode As Variant
    Dim mMode As Integer '1=PaymentOrder; 2=Approval

Private Sub Calculate()
    Dim mLoop As Integer
    Dim mAmt As Double
    Dim objAc As New clsAccounts
    
    For mLoop = 1 To Grid.Rows - 1
        If Trim(Grid.TextMatrix(mLoop, 2)) <> "" And val(Grid.TextMatrix(mLoop, 3)) Then
            mAmt = mAmt + val(Grid.TextMatrix(mLoop, 3))
        End If
    Next mLoop
'''    txtTotalDeduction.Text = Format(mAmt, "0.00")
'''    txtNetSalary.Text = Format(val(txtGrossAmount) - val(mAmt), "0.00")
'''    txtAmtPension.Text = Format(val(txtGrossAmount) * 15 / 100, "0.00")
    
    Select Case mSalaryAccountHeadCode
        Case Is = "210100101" ' Salaries - Secretary-->Contribution to Pension Fund - Regular employees-Secretary
            objAc.SetAccountCode "210300101"
            If objAc.AccountHeadID > 0 Then
                txtPensionCode.Text = objAc.AccountCode
                txtPensionContribution.Text = objAc.AccountHead
            End If
        Case Is = "210100106" ' Salaries - Contingent Staff -->Contribution to Pension Fund - Contingent Staff
            objAc.SetAccountCode "210300201"
            If objAc.AccountHeadID > 0 Then
                txtPensionCode.Text = objAc.AccountCode
                txtPensionContribution.Text = objAc.AccountHead
            End If
''''''''''''        Case Is = "210100105"   'Temperory Staff
''''''''''''            txtPensionCode.Text = ""
''''''''''''            txtPensionContribution.Text = ""
''''''''''''            Exit Sub
        Case Else
            objAc.SetAccountCode "210300100"     'Contribution to Pension Fund - Regular employees
            If objAc.AccountHeadID > 0 Then
                txtPensionCode.Text = objAc.AccountCode
                txtPensionContribution.Text = objAc.AccountHead
            End If
    End Select
End Sub

Private Sub FormInitialize()
    Dim mCrl As Control
    For Each mCrl In Me.Controls
        If TypeOf mCrl Is TextBox Then
            mCrl.Text = ""
            mCrl.Tag = ""
        ElseIf TypeOf mCrl Is OptionButton Then
            mCrl.value = False
        ElseIf TypeOf mCrl Is ComboBox Then
            If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
        ElseIf TypeOf mCrl Is ComboBox Then
            mCrl.ListIndex = -1
        End If
    Next
    
    If gbUserTypeID = 3 Then
        mMode = 1
    ElseIf gbUserTypeID = 2 Then
        mMode = 2
    Else
        mMode = 0
    End If
    
    lblStatus.Caption = "*"
    mSalaryAccountHeadCode = ""
    Frame2.Visible = False
    Frame1.Visible = True
End Sub

Private Sub ShowAbstract(mID As Long, mFunctionID As Variant)
    '----------------------------------------------------------------------------'
    ' Note:- Fetch Bill Abstracts Details From Sthapana And
    '        Display in Grid
    '        Where mID is Section ID, mFunctionID is Function ID
    '----------------------------------------------------------------------------'
    Frame1.Visible = False
    Frame2.Visible = True
    
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mRow As Integer
    Dim arrInput As Variant
    Dim objFnry As New clsFunctionary
    Dim objFn As New clsFunction
    Dim mDueDate As Date
    Dim mSql As String
    
    cmdPaymentOrder.Enabled = True
    txtPayOrderNo.Text = ""
    If IsNumeric(mID) Then
        arrInput = Array(mID, mFunctionID)
    Else
        arrInput = Array(0)
    End If
    objdb.CreateNewConnection mCnn, enuSourceString.Sthapana
    If mCnn.State Then
        Set Rec = objdb.ExecuteSP("spGetAbstractsNew", arrInput, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
            Me.MousePointer = vbHourglass
            If gbUserTypeID = UserType.Approver Then
                cmdPaymentOrder.Caption = "&Approve"
            Else
                cmdPaymentOrder.Caption = "&PaymentOrder"
            End If
            cmdPaymentOrder.Visible = True
            
            txtPayOrderNo.Text = GridAbstract.TextMatrix(GridAbstract.Row, 3)
            txtSection.Text = IIf(IsNull(Rec!chvSectionName), "", Rec!chvSectionName)
            txtSection.Tag = IIf(IsNull(Rec!intSecID), "", Rec!intSecID)
            
            objFnry.SetFunctionary (Rec!chvFunctionaryCode)
            If objFnry.FunctionaryID > 0 Then
                txtFunctionary.Text = objFnry.FunctionaryName
                txtFunctionary.Tag = objFnry.FunctionaryID
            End If
            objFn.SetFunction (Rec!chvFunctionCode)
            If objFn.FunctionID > 0 Then
                txtFunction.Text = objFn.FunctionName
                txtFunction.Tag = objFn.FunctionID
            End If
            txtBillNo.Text = Rec!BillID 'Modified By Anisha
            txtID.Text = Rec!intSecID
            txtYearID.Text = Rec!intYearID
            txtMonth.Text = Format(DateSerial(Rec!intYearID, Rec!intMonthID, 1), "mmmm")
            txtMonth.Tag = Rec!intMonthID
            mDueDate = "1-" & txtMonth & "-" & txtYear 'DateSerial(Val(txtYear), txtMonth, 1)
            mDueDate = DateAdd("d", -1, DateAdd("m", 1, mDueDate))
            txtDueDate.Text = DdMmmYy(mDueDate)
            
            txtGrossAmount.Text = Format(Rec!Gross, "0.00")
            Grid.Rows = 50
            mRow = 0
            If Not IsNull(Rec!LeaveAdj) And Rec!LeaveAdj > 0 Then   ' (210100101,210100102,210100103,210100104,210100105,210100106)
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Leave Adjustment"
                Grid.TextMatrix(mRow, 2) = mSalaryAccountHeadCode
                Grid.TextMatrix(mRow, 3) = Rec!LeaveAdj
            End If
            If Not IsNull(Rec!QuartersRent) And Rec!QuartersRent > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Quarters Rent"
                Grid.TextMatrix(mRow, 2) = "130200100"
                Grid.TextMatrix(mRow, 3) = Rec!QuartersRent
            End If
            If Not IsNull(Rec!ExcessPay) And Rec!ExcessPay > 0 Then ' (210100101,210100102,210100103,210100104,210100105,210100106)
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Recovery Of Excess Salary Paid"
                Grid.TextMatrix(mRow, 2) = mSalaryAccountHeadCode
                Grid.TextMatrix(mRow, 3) = Rec!ExcessPay
            End If
            If Not IsNull(Rec!Family) And Rec!Family > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Family Benefit Scheme"
                Grid.TextMatrix(mRow, 2) = "350200128"
                Grid.TextMatrix(mRow, 3) = Rec!Family
            End If
            If Not IsNull(Rec!DAArrearToPF) And Rec!DAArrearToPF > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "DA Arrear to PF"
                Grid.TextMatrix(mRow, 2) = "350200101"
                Grid.TextMatrix(mRow, 3) = Rec!DAArrearToPF
            End If
            If Not IsNull(Rec!PRA) And Rec!PRA > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Pay Revision Arrear to PF"
                Grid.TextMatrix(mRow, 2) = "350200101"
                Grid.TextMatrix(mRow, 3) = Rec!PRA
            End If
            If Not IsNull(Rec!PFSubscription) And Rec!PFSubscription > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "PF Subscription"
                Grid.TextMatrix(mRow, 2) = "350200101"
                Grid.TextMatrix(mRow, 3) = Rec!PFSubscription
            End If
            If Not IsNull(Rec!PFLoan) And Rec!PFLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "PF Loan"
                Grid.TextMatrix(mRow, 2) = "350200101"
                Grid.TextMatrix(mRow, 3) = Rec!PFLoan
            End If
            If Not IsNull(Rec!PFArrear) And Rec!PFArrear > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "PF Subscription Arrear"
                Grid.TextMatrix(mRow, 2) = "350200101"
                Grid.TextMatrix(mRow, 3) = Rec!PFArrear
            End If
            If Not IsNull(Rec!Income) And Rec!Income > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Income Tax"
                Grid.TextMatrix(mRow, 2) = "350200109"
                Grid.TextMatrix(mRow, 3) = Rec!Income
            End If
            If Not IsNull(Rec!ProfessionTax) And Rec!ProfessionTax > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Profession Tax"
                Grid.TextMatrix(mRow, 2) = "110100200"
                Grid.TextMatrix(mRow, 3) = Rec!ProfessionTax
            End If
            If Not IsNull(Rec!LICEmp) And Rec!LICEmp > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "LIC Premia/LIC Arrear/LIC Loan"
                Grid.TextMatrix(mRow, 2) = "350200104"
                Grid.TextMatrix(mRow, 3) = Rec!LICEmp
            End If
            If Not IsNull(Rec!SLIAmt) And Rec!SLIAmt > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "State Life Insurance/ Arrear of SLI"
                Grid.TextMatrix(mRow, 2) = "350200116"
                Grid.TextMatrix(mRow, 3) = Rec!SLIAmt
            End If
            If Not IsNull(Rec!GSLIAmt) And Rec!GSLIAmt > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Group Saving Life Insurane/ Arrear of GSLI"
                Grid.TextMatrix(mRow, 2) = "350200117"
                Grid.TextMatrix(mRow, 3) = Rec!GSLIAmt
            End If
            If Not IsNull(Rec!GIAmt) And Rec!GIAmt > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Group Insurarnce/ Arrear of GIS"
                Grid.TextMatrix(mRow, 2) = "350200118"
                Grid.TextMatrix(mRow, 3) = Rec!GIAmt
            End If
            If Not IsNull(Rec!PostalLifeInsurance) And Rec!PostalLifeInsurance > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Recurring Deposit/Postal Life Insurance"
                Grid.TextMatrix(mRow, 2) = "350200104"
                Grid.TextMatrix(mRow, 3) = Rec!PostalLifeInsurance
            End If
            If Not IsNull(Rec!HouseLoan) And Rec!HouseLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "House Loan"
                Grid.TextMatrix(mRow, 2) = "460100100"
                Grid.TextMatrix(mRow, 3) = Rec!HouseLoan
            End If
            If Not IsNull(Rec!MarriageLoan) And Rec!MarriageLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Marriage Loan"
                Grid.TextMatrix(mRow, 2) = "460100800"
                Grid.TextMatrix(mRow, 3) = Rec!MarriageLoan
            End If
            If Not IsNull(Rec!VehicleAdvance) And Rec!VehicleAdvance > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Vehicle Advance (Scooter/Cycle Etc)"
                Grid.TextMatrix(mRow, 2) = "460100200"
                Grid.TextMatrix(mRow, 3) = Rec!VehicleAdvance
            End If
            If Not IsNull(Rec!FestivalAdvance) And Rec!FestivalAdvance > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Festival Advance"
                Grid.TextMatrix(mRow, 2) = "460100400"
                Grid.TextMatrix(mRow, 3) = Rec!FestivalAdvance
            End If
            If Not IsNull(Rec!WelfareFund) And Rec!WelfareFund > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Welfare Subscription"
                Grid.TextMatrix(mRow, 2) = "350200120"
                Grid.TextMatrix(mRow, 3) = Rec!WelfareFund
            End If
            If Not IsNull(Rec!WelfareLoan) And Rec!WelfareLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Welfare Loan"
                Grid.TextMatrix(mRow, 2) = "350200199"
                Grid.TextMatrix(mRow, 3) = Rec!WelfareLoan
            End If
            If Not IsNull(Rec!BankLoan) And Rec!BankLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Bank Loan"
                Grid.TextMatrix(mRow, 2) = "350200103"
                Grid.TextMatrix(mRow, 3) = Rec!BankLoan
            End If               'PTA WithOut Vehicle
            If Not IsNull(Rec!KSFE) And Rec!KSFE > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "KSFE Loan"
                Grid.TextMatrix(mRow, 2) = "350200107"
                Grid.TextMatrix(mRow, 3) = Rec!KSFE
            End If                 '--Photostat Allowance
            If Not IsNull(Rec!CSLoan) And Rec!CSLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Co Op. Socity Loan"
                Grid.TextMatrix(mRow, 2) = "350200106"
                Grid.TextMatrix(mRow, 3) = Rec!CSLoan
            End If                ',--Theatre Checking  Allowance
            If Not IsNull(Rec!CourtAttachment) And Rec!CourtAttachment > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Court Attachment"
                Grid.TextMatrix(mRow, 2) = "350200105"
                Grid.TextMatrix(mRow, 3) = Rec!CourtAttachment
            End If                 '--charge Allowance
            If Not IsNull(Rec!AccidentCompensationRecovery) And Rec!AccidentCompensationRecovery > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Gr. Personal Accident Insurance Scheme/ACR"
                Grid.TextMatrix(mRow, 2) = "350200199"
                Grid.TextMatrix(mRow, 3) = Rec!AccidentCompensationRecovery
            End If            '
            If Not IsNull(Rec!ElecBill) And Rec!ElecBill > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Electricity Bill/Farewell Fund"
                Grid.TextMatrix(mRow, 2) = "350200123"
                Grid.TextMatrix(mRow, 3) = Rec!ElecBill
            End If
            If Not IsNull(Rec!CostOfLand) And Rec!CostOfLand > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Cost of Land"
                Grid.TextMatrix(mRow, 2) = "350200199"
                Grid.TextMatrix(mRow, 3) = Rec!CostOfLand
            End If
            If Not IsNull(Rec!AuditRecovery) And Rec!AuditRecovery > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Audit Recovery"
                Grid.TextMatrix(mRow, 2) = "180400100"
                Grid.TextMatrix(mRow, 3) = Rec!AuditRecovery
            End If
            If Not IsNull(Rec!MedicalLoan) And Rec!MedicalLoan > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Medical Loan"
                Grid.TextMatrix(mRow, 2) = "460109900"
                Grid.TextMatrix(mRow, 3) = Rec!MedicalLoan
            End If
            If Not IsNull(Rec!StampRecovery) And Rec!StampRecovery > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Stamp Recovery"
                Grid.TextMatrix(mRow, 2) = "350200199"
                Grid.TextMatrix(mRow, 3) = Rec!StampRecovery
            End If
            If Not IsNull(Rec!KBDC) And Rec!KBDC > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Recovery to Other LSGIs"
                Grid.TextMatrix(mRow, 2) = "350200108"
                Grid.TextMatrix(mRow, 3) = Rec!KBDC
            End If
            If Not IsNull(Rec!PORD) And Rec!PORD > 0 Then
                mRow = mRow + 1
                Grid.TextMatrix(mRow, 0) = mRow
                Grid.TextMatrix(mRow, 1) = "Post Office Recurring Deposit"
                Grid.TextMatrix(mRow, 2) = "350209900"
                Grid.TextMatrix(mRow, 3) = Rec!PORD
            End If
            Grid.Rows = mRow + 1
            
            
            '-----Modified on 27.09.2011 by Minu----------------------------
            txtTotalDeduction.Text = Format(Rec!DeductionTotal, "0.00") 'Format(mAmt, "0.00")
            txtNetSalary.Text = Format(Rec!NetTotal, "0.00") 'Format(val(txtGrossAmount) - val(mAmt), "0.00")
            txtAmtPension.Text = Format(Rec!PC, "0.00") 'Format(val(txtGrossAmount) * 15 / 100, "0.00")
            
            
            Call Calculate
            Me.MousePointer = vbDefault
            mCnn.Close
            'End of fetching Displaying Data From Sthapana
            
            'Note:- Check Wether PayOrder is Approved Or Not
            If gbUserTypeID = UserType.Approver Then
                mSql = "Select * From faPayOrder Where vchPayOrderNo = " & val(txtPayOrderNo) & " And tnyStatus = 0"
                objdb.SetConnection mCnn
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                If (Rec.BOF And Rec.EOF) Then
                   cmdPaymentOrder.Enabled = False
                   lblStatus.Caption = " ( Approved and Journalized for Payment ) "
                End If
                Rec.Close
            End If
            
        End If
        
    End If
       
End Sub

Private Sub FillCombo()
    Dim mSql As String
    
    cmbMonth.AddItem "", 0
    cmbMonth.ItemData(0) = 0
    
    cmbMonth.AddItem "April", 1
    cmbMonth.ItemData(1) = 4
    
    cmbMonth.AddItem "May", 2
    cmbMonth.ItemData(2) = 5
    
    cmbMonth.AddItem "June", 3
    cmbMonth.ItemData(3) = 6
    
    cmbMonth.AddItem "July", 4
    cmbMonth.ItemData(4) = 7
    
    cmbMonth.AddItem "August", 5
    cmbMonth.ItemData(5) = 8
    
    cmbMonth.AddItem "September", 6
    cmbMonth.ItemData(6) = 9
    
    cmbMonth.AddItem "October", 7
    cmbMonth.ItemData(7) = 10
    
    cmbMonth.AddItem "November", 8
    cmbMonth.ItemData(8) = 11
    
    cmbMonth.AddItem "December", 9
    cmbMonth.ItemData(9) = 12
    
    cmbMonth.AddItem "January", 10
    cmbMonth.ItemData(10) = 1
    
    cmbMonth.AddItem "February", 11
    cmbMonth.ItemData(11) = 2
    
    cmbMonth.AddItem "March", 12
    cmbMonth.ItemData(12) = 3
    
    mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) As AccHead, intAccountHeadID From faAccountHeads Where vchAccountHeadCode Between 210100101 And 210100109"
    PopulateList cmbAccountHeads, mSql, "210100104  Salaries - Permanent Staff", True, , True
    If Len(cmbAccountHeads) > 9 Then
        mSalaryAccountHeadCode = Left(cmbAccountHeads.Text, 9)
    Else
        mSalaryAccountHeadCode = ""
    End If
End Sub

Private Sub FetchAbstracts()
    '----------------------------------------------------------------------------'
    'Note:- Fetch Bill Abstracts Details From Sthapana And
    '        Display in Grid
    '----------------------------------------------------------------------------'
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mRow As Integer
    Dim mYearID As Integer
    Dim mMonthID As Integer
    Dim mSql As String
    
    objdb.CreateNewConnection mCnn, enuSourceString.Sthapana
    If mCnn.State Then
        Rec.CursorLocation = adUseClient
        Set Rec = objdb.ExecuteSP("spGetAbstractsNew", , , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
            
            Me.MousePointer = vbHourglass
            mRow = 0
            mYearID = Rec!intYearID
            mMonthID = Rec!intMonthID
            txtYear.Text = mYearID
            GridAbstract.Rows = 1
            cmbMonth.Text = Format(DateSerial(gbFinancialYearID, mMonthID, 25), "MMMM")
            While Not Rec.EOF
                mRow = mRow + 1
                GridAbstract.Rows = mRow + 1
                GridAbstract.TextMatrix(mRow, 0) = mRow
                GridAbstract.TextMatrix(mRow, 1) = Rec!chvSectionName
                GridAbstract.TextMatrix(mRow, 2) = Rec!NetTotal
                GridAbstract.TextMatrix(mRow, 3) = "" ' Payment Order
                GridAbstract.TextMatrix(mRow, 4) = "" ' Payment Voucher No
                GridAbstract.TextMatrix(mRow, 5) = "" ' Check No
                GridAbstract.TextMatrix(mRow, 6) = Rec!intSecID
                GridAbstract.TextMatrix(mRow, 8) = Rec!BillID 'Added On 9/03/10
                GridAbstract.TextMatrix(mRow, 9) = Rec!chvFunctionCode
                Rec.MoveNext
            Wend
            Rec.Close
            mCnn.Close
            GridAbstract.Select 1, 6
            GridAbstract.Sort = flexSortGenericAscending
            
            mSql = "Select faPayOrder.vchPayOrderNo,faPayOrder.intVoucherNo,faPayOrder.tnyStatus,faPayOrder.intKeyID,faPayOrder.vchBillNo,faFunctions.vchFunctionCode,faVouchers.vchInstrumentNo  From faPayOrder inner join faFunctions on faPayOrder.intFunctionID=faFunctions.intFunctionID"
            mSql = mSql + "  left join faVouchers on faVouchers.intVoucherNo=faPayOrder.intVoucherNo "
            mSql = mSql + " Where faPayOrder.intTransactionTypeID = " & gbTransactionTypePayBills & " "
            mSql = mSql + " And dtKeyDate = '" & DdMmmYy(DateSerial(mYearID, mMonthID, 1)) & "' Order By intKeyID"
            Dim mLoop As Integer
            mLoop = 1
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
            For mLoop = mLoop To GridAbstract.Rows - 1
              If Not (Rec.EOF Or Rec.BOF) Then
                 While Not Rec.EOF  '**************TO MODIFY************'MODIFIED BY MINU ON 28-04-2011
                    If val(GridAbstract.Cell(flexcpText, mLoop, 6)) = Rec!intKeyID And val(GridAbstract.Cell(flexcpText, mLoop, 8)) = Rec!vchBillNo And val(GridAbstract.Cell(flexcpText, mLoop, 9)) = Rec!vchFunctionCode Then
                       GridAbstract.Cell(flexcpText, mLoop, 3) = Rec!vchPayOrderNo
                       GridAbstract.Cell(flexcpText, mLoop, 7) = Rec!tnyStatus
                       If GridAbstract.Cell(flexcpText, mLoop, 7) = 1 Then
                         GridAbstract.Cell(flexcpText, mLoop, 4) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                        GridAbstract.Cell(flexcpText, mLoop, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                       End If

                    End If
                    Rec.MoveNext
                  Wend
                  Rec.MoveFirst
               End If
            Next
            
            GridAbstract.Select 1, 0
            GridAbstract.Sort = flexSortGenericAscending
            Rec.Close
            
            Me.MousePointer = vbDefault
        Else
            'Note:- Not found any generated Bills
            Exit Sub
        End If
    Else 'Note:- Connection Failed to STHAPANA
        MsgBox "Didn't able to connect Sthapana Database", vbInformation
        Exit Sub
    End If
End Sub

Private Sub cmbAccountHeads_LostFocus()
    If Len(cmbAccountHeads.Text) > 9 Then
        mSalaryAccountHeadCode = Left(cmbAccountHeads.Text, 9)
    Else
        mSalaryAccountHeadCode = ""
    End If
    Call Calculate
End Sub

Private Sub cmdPaymentOrder_Click()

    Dim PaymentOrder As uPaymentOrder
    Dim PaymentOrderChild As uPaymentOrderChild
    Dim PaymentOrderAddress As uPaymentOrderAddress
    Dim objdb As New clsDB
    Dim objAc As New clsAccounts
    Dim mCnn As New ADODB.Connection
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim mPaymentOrderID As Variant
    Dim mPaymentOrderNo As Variant
    Dim mGrossAmt As Variant
    Dim mLoopCount As Integer
    Dim mSLNo As Integer
    Dim mSalaryHeadID As Integer
    
    If cmbAccountHeads.ListIndex > -1 Then
        mSalaryHeadID = cmbAccountHeads.ItemData(cmbAccountHeads.ListIndex)
        objAc.SetAccountID mSalaryHeadID
        If objAc.AccountHeadID <= 0 Then
            MsgBox "Please select a Salary Acount Head!", vbInformation
            cmbAccountHeads.SetFocus
            Exit Sub
        End If
    End If
    If txtForward2Seat.Tag = "" Then
        MsgBox "Please select the Forwarded Seat", vbInformation
        cmdSeat.SetFocus
        Exit Sub
    End If
    If txtName.Text = "" Then
        MsgBox "Please enter the Name Of Payee", vbInformation
        txtName.SetFocus
        Exit Sub
    End If
    If txtSourceOfFund.Tag = "" Then
        MsgBox "Please select the Source Of Fund", vbInformation
        txtSourceOfFund.SetFocus
        Exit Sub
    End If

    If mMode = 1 Then
                'Note:- Saving Payment Order
                txtPayOrderNo.Text = ""
                objdb.SetConnection mCnn
                With PaymentOrder
                    .intPayOrderID = Null
                    .vchPayOrderNo = Null
                    .dtPayOrderDate = gbDate
                    .dtDueDate = txtDueDate.Text
                    .intFunctionaryID = val(txtFunctionary.Tag)
                    .intFunctionID = val(txtFunction.Tag)
                    .intTransactionTypeID = gbTransactionTypePayBills
                    .vchBillNo = Trim(txtBillNo.Text) 'added on 9/03/10
                    .numBillAmount = Null
                    .dtBillDate = Null
                    .intInstrumentTypeID = gbInstrumentCheque
                    .intCashOrBankHeadID = mSalaryHeadID
                    '.vchDescription = "Salary credited in section " & txtSection & " for the month " & txtMonth & ", " & txtYear
                    '.vchTitle = "Salary for the month " & txtMonth & ", " & txtYear & "( Section : " & txtSection & " )"
                    
                    .vchDescription = "Salary credited in section " & txtSection & " for the month " & txtMonth & " " & txtYear
                    .vchTitle = "Salary for the month " & txtMonth & " " & txtYear & "( Section : " & txtSection & " )"
                   
                    .intSubLedgerTypeID = 0
                    .intPayToSubLedgerID = 0
                    .intSubsidiaryCashBookID = 0
                    .intImplementingOfficerID = 0
                    .numProjectNo = 0
                    .intStockRegisterID = Null
                    .vchStockRefNo = Null
                    .intAssetTypeID = Null
                    .intAssetID = Null
                    .numFwdSeatID = val(txtForward2Seat.Tag)
                    .intLocalBodyID = gbLocalBodyID
                    .intZonalID = gbLocationID
                    .intFinancialYearID = gbFinancialYearID
                    .numUserID = gbUserID
                    .numSeatID = gbSeatID
                    .numApprovingOfficerID = Null
                    .numApprovingSeatID = Null
                    .dtApprovingDate = Null
                    .intVoucherID = Null
                    .intVoucherNo = Null
                    .dtVoucherDate = Null
                    .intKeyID = val(txtID)
                    .numKeyID = Null
                    .dtKeyDate = DateSerial(val(txtYear), val(txtMonth.Tag), 1)
                    .tnyStatus = 0
                    .tnyCancelled = Null
                    .intAppID = 115
                    .intModuleID = 1001
                    .intSourceOfFundID = val(txtSourceOfFund.Tag)
                    arrInput = Array(.intPayOrderID, _
                    .vchPayOrderNo, _
                    .dtPayOrderDate, _
                    .dtDueDate, _
                    .intFunctionaryID, _
                    .intFunctionID, _
                    .intTransactionTypeID, _
                    .vchBillNo, .numBillAmount, .dtBillDate, _
                    .intInstrumentTypeID, .intCashOrBankHeadID, .vchDescription, .vchTitle, .intSubLedgerTypeID, _
                    .intPayToSubLedgerID, .intSubsidiaryCashBookID, .intImplementingOfficerID, _
                    .numProjectNo, .intStockRegisterID, .vchStockRefNo, _
                    .intAssetTypeID, .intAssetID, .numFwdSeatID, _
                    .intLocalBodyID, .intZonalID, .intFinancialYearID, _
                    .numUserID, .numSeatID, .numApprovingOfficerID, _
                    .numApprovingSeatID, .dtApprovingDate, .intVoucherID, _
                    .intVoucherNo, .dtVoucherDate, .tnyStatus, _
                    .intKeyID, .numKeyID, .dtKeyDate, .tnyCancelled, .intAppID, .intModuleID, .intSourceOfFundID _
                    )
                End With
                mCnn.BeginTrans
                objdb.ExecuteSP "spSavePayOrder", arrInput, arrOutPut, , mCnn, adCmdStoredProc
                If IsNumeric(arrOutPut(0, 0)) Then
                    mPaymentOrderID = arrOutPut(0, 0)
                    mPaymentOrderNo = arrOutPut(1, 0)
                    
                    'Note:- Gross Salary Payable to PaymentOrderChild Table
                    With PaymentOrderChild
                        .intPayOrderID = mPaymentOrderID
                        .intSlNo = 1
                        objAc.SetAccountID mSalaryHeadID
                        If objAc.AccountHeadID > 0 Then
                            .intAccountHeadID = mSalaryHeadID       'gbAcHeadIDGrossSalaryPayable
                            .vchAccountHeadCode = objAc.AccountCode 'gbAcHeadCodeGrossSalaryPayable
                        Else
                            MsgBox "Error: Salary Head Not Found! :( ", vbInformation
                            GoTo ErrRollBank:
                        End If
                        .numAmount = val(txtGrossAmount)
                        .tnyCategoryFlag = 1
                        .tnyDebitOrCreditFlag = 1
                        .vchDescription = Null
                        
                        
                        arrInput = Array(.intPayOrderID, _
                        .intSlNo, _
                        .intAccountHeadID, _
                        .vchAccountHeadCode, _
                        .numAmount, _
                        .tnyCategoryFlag, _
                        .tnyDebitOrCreditFlag, _
                        .vchDescription)
                        objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
                    End With
                
                    'Note:- All Deduction Heads to PaymentOrderChild (From Grid)
                    For mLoopCount = 1 To Grid.Rows - 1
                        If val(Grid.TextMatrix(mLoopCount, 3)) > 0 And Trim(Grid.TextMatrix(mLoopCount, 2)) <> "" Then
                                With PaymentOrderChild
                                    .intPayOrderID = mPaymentOrderID
                                    .intSlNo = mSLNo + 1
                                    objAc.SetAccountCode Trim(Grid.TextMatrix(mLoopCount, 2))
                                    If objAc.AccountHeadID < 1 Then
                                        MsgBox "Account Head not found for " & Grid.TextMatrix(mLoopCount, 1)
                                        GoTo ErrRollBank:
                                    End If
                                    .intAccountHeadID = objAc.AccountHeadID
                                    .vchAccountHeadCode = objAc.AccountCode
                                    .numAmount = val(Grid.TextMatrix(mLoopCount, 3))
                                    .tnyCategoryFlag = 2
                                    .tnyDebitOrCreditFlag = 0
                                    
                                    arrInput = Array(.intPayOrderID, _
                                    .intSlNo, _
                                    .intAccountHeadID, _
                                    .vchAccountHeadCode, _
                                    .numAmount, _
                                    .tnyCategoryFlag, _
                                    .tnyDebitOrCreditFlag, _
                                    .vchDescription)
                                    
                                    objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
                                End With
                        End If
                    Next
        
                    'Note:- Net Salary Payable to PaymentOrderChild
                    With PaymentOrderChild
                        .intPayOrderID = mPaymentOrderID
                        .intSlNo = mSLNo + 1
                        .intAccountHeadID = gbAcHeadIDNetSalaryPayable
                        .vchAccountHeadCode = gbAcHeadCodeNetSalaryPayable
                        .numAmount = val(txtNetSalary)
                        .tnyCategoryFlag = 3
                        .tnyDebitOrCreditFlag = 0
                        
                        arrInput = Array(.intPayOrderID, _
                                .intSlNo, _
                                .intAccountHeadID, _
                                .vchAccountHeadCode, _
                                .numAmount, _
                                .tnyCategoryFlag, _
                                .tnyDebitOrCreditFlag, _
                                .vchDescription)
                        objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
                    End With
                    
                    If val(txtAmtPension.Text) > 0 Then
                        mSLNo = mSLNo + 1
                        With PaymentOrderChild
                            .intPayOrderID = mPaymentOrderID
                            .intSlNo = mSLNo
                            .intAccountHeadID = 0
                            .vchAccountHeadCode = 0
                            .numAmount = val(txtAmtPension)
                            .tnyCategoryFlag = 5
                            .tnyDebitOrCreditFlag = 0
                            .vchDescription = "Pension Contribution Amount"
                            
                            arrInput = Array(.intPayOrderID, _
                            .intSlNo, _
                            .intAccountHeadID, _
                            .vchAccountHeadCode, _
                            .numAmount, _
                            .tnyCategoryFlag, _
                            .tnyDebitOrCreditFlag, _
                            .vchDescription)
                            objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
                        End With
                    End If
                        
                    'Note:- Saving Payment Order Address Table
                    With PaymentOrderAddress
                        .intPayOrderID = mPaymentOrderID
                        .intSubsidiaryAccountHeadID = Null
                        .intSubLegerTypeID = Null
                        .vchSubLedgerCode = Null
                        .vchName = txtName.Text
                        .vchHouseName = Null
                        .vchStreet = Null
                        .vchLocalPlace = Null
                        .vchMainPlace = Null
                        .vchPost = Null
                        .vchPinCode = Null
                        .vchPhone = Null
                        
                        arrInput = Array(.intPayOrderID, _
                                    .intSubsidiaryAccountHeadID, _
                                    .intSubLegerTypeID, _
                                    .vchSubLedgerCode, _
                                    .vchName, _
                                    .vchHouseName, _
                                    .vchStreet, _
                                    .vchLocalPlace, _
                                    .vchMainPlace, _
                                    .vchPost, _
                                    .vchPinCode, _
                                    .vchPhone)
                                    
                        objdb.ExecuteSP "spSavePayOrderAddress", arrInput, , , mCnn, adCmdStoredProc
                    End With
                    mCnn.CommitTrans
                    txtPayOrderNo.Text = mPaymentOrderNo
                Else ' If IsNumeric(arrOutPut(0, 0)) Then
                    GoTo Err:
                End If
                
                GridAbstract.TextMatrix(GridAbstract.Row, 3) = txtPayOrderNo.Text
                cmdPaymentOrder.Enabled = False
                Exit Sub
ErrRollBank:
                mCnn.RollbackTrans
                Set mCnn = Nothing
    ElseIf mMode = 2 Then  ' Mode = 2 Approving Payment Order
        'Note:- Added By Aiby
         cmdPaymentOrder.Enabled = False
         Call GeneratePayBillJournals(val(txtPayOrderNo))
         
    Else
    
    End If
Err:

End Sub

    Private Sub cmdSearchSourceofFund_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.Show vbModal
        'txtSourceOfFund.SetFocus
        If gbSearchID <> -1 Then
            txtSourceOfFund.Text = gbSearchStr
            txtSourceOfFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID <> -1 Then
            txtForward2Seat.Text = gbSearchStr
            txtForward2Seat.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Frame2.Visible Then
            Frame2.Visible = False
            Frame1.Visible = True
            cmdPaymentOrder.Visible = False
            lblStatus.Caption = "*"
        Else
            Unload Me
        End If
    End If
End Sub
Private Sub Form_Load()
    Call FormInitialize
    Call FillCombo
    'Call FetchAbstracts
    WindowsXPC.InitIDESubClassing
    Frame1.Left = 0
    Frame1.Top = -90
End Sub
Private Sub GridAbstract_DblClick()
    '--------------------------------------------------------------------'
    ' Note:- Get the section ID from the selected grid                   '
    '         and fetch the abstract details and display in Grid         '
    '--------------------------------------------------------------------'
    Dim mID As Long
    Dim mFunctionID As Variant 'Added on 9/03/10
    mID = val(GridAbstract.TextMatrix(GridAbstract.Row, 6))
    mFunctionID = GridAbstract.TextMatrix(GridAbstract.Row, 9) 'val(GridAbstract.TextMatrix(GridAbstract.Row, 9))
    If mID > 0 Then
        If val(GridAbstract.TextMatrix(GridAbstract.Row, 3)) = 0 Then
            Call ShowAbstract(mID, mFunctionID)
        Else
            If gbUserTypeID = UserType.Approver Then
                Call ShowAbstract(mID, mFunctionID)
            End If
        End If
    End If
End Sub
Private Sub Timer_Timer()
    If Timer.Enabled Then
        Timer.Enabled = False
        Me.MousePointer = vbHourglass
        Call FetchAbstracts
        Me.MousePointer = vbDefault
    End If
End Sub


