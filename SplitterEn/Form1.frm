VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "ãÞÓã ÇáãáÝÇÊ - ãÇÑßæ ÓÇãí"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4800
      ScaleHeight     =   825
      ScaleWidth      =   4545
      TabIndex        =   48
      Top             =   840
      Width           =   4575
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   2760
         Top             =   120
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   240
         TabIndex        =   50
         Top             =   -40
         Width           =   1935
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   4440
         X2              =   3840
         Y1              =   360
         Y2              =   120
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         X1              =   3840
         X2              =   3240
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Line Line9 
         BorderWidth     =   3
         X1              =   4440
         X2              =   3240
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Here"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2640
         TabIndex        =   49
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   2025
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minimize"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox PicT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000000FF&
      Caption         =   "Non Order Split"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   9135
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000FF&
      Caption         =   "Order Split"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      MouseIcon       =   "Form1.frx":045C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   9135
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   5160
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1785
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   9105
      TabIndex        =   6
      Top             =   3240
      Width           =   9100
      Begin ComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   10
         Min             =   2
         Max             =   500
         SelStart        =   2
         Value           =   2
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Resize"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7800
         MouseIcon       =   "Form1.frx":05AE
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Min 0.0 KB"
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Max 0.0 KB"
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   33
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Adding Part in 0 KB"
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "is 100% of the file."
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Parts Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         MouseIcon       =   "Form1.frx":0700
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5040
         MouseIcon       =   "Form1.frx":0852
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 MB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         MouseIcon       =   "Form1.frx":09A4
         MousePointer    =   99  'Custom
         TabIndex        =   21
         ToolTipText     =   "Part 1 Size 0 MB"
         Top             =   240
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.Shape Part 
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   180
         Top             =   180
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Top             =   120
         Width           =   8895
      End
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1305
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   9105
      TabIndex        =   5
      Top             =   1440
      Width           =   9100
      Begin VB.PictureBox Hide1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   6360
         ScaleHeight     =   1335
         ScaleWidth      =   2655
         TabIndex        =   43
         Top             =   0
         Width           =   2655
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Click Here To Add Your Own"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            MouseIcon       =   "Form1.frx":0AF6
            MousePointer    =   99  'Custom
            TabIndex        =   47
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2.76MB is DFD Capacity"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "650MB is Cd-Rom Capacity"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1.38MB is Folppy Capacity"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4920
         MaxLength       =   5
         TabIndex        =   19
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Height          =   315
         ItemData        =   "Form1.frx":0C48
         Left            =   6360
         List            =   "Form1.frx":0C5E
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Height          =   340
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "0"
         Top             =   420
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form1.frx":0C7C
         Left            =   120
         List            =   "Form1.frx":0C95
         TabIndex        =   14
         Text            =   "1.38MB"
         ToolTipText     =   "Instant add(write your new value here)"
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         MouseIcon       =   "Form1.frx":0CCE
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7800
         MouseIcon       =   "Form1.frx":0E20
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         MouseIcon       =   "Form1.frx":0F72
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         MouseIcon       =   "Form1.frx":10C4
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Space"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Units Number :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   120
         Width           =   1575
      End
      Begin VB.Line Line8 
         X1              =   6240
         X2              =   6240
         Y1              =   120
         Y2              =   1200
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      MouseIcon       =   "Form1.frx":1216
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8775
      MouseIcon       =   "Form1.frx":1368
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8220
      MouseIcon       =   "Form1.frx":14BA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   8760
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   8220
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "Form1.frx":160C
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":175E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin ComctlLib.ImageList Im 
      Left            =   -120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   327682
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   600
      Top             =   480
      Width           =   6975
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   60
      Picture         =   "Form1.frx":1A68
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   9360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7080
      MouseIcon       =   "Form1.frx":1D72
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size is : 0.0 Bytes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   795
      Width           =   6975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Collect"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      MouseIcon       =   "Form1.frx":1EC4
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      MouseIcon       =   "Form1.frx":2016
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9240
      Y1              =   3240
      Y2              =   5040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3240
      Y2              =   5040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9240
      Y1              =   1440
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Select File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Double Click here to change"
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      MouseIcon       =   "Form1.frx":2168
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillStyle       =   7  'Diagonal Cross
      Height          =   5295
      Left            =   0
      Top             =   360
      Width           =   9375
   End
   Begin VB.Label State 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "By Marco Samy Nasif, Status: Ready  - For Help Click The Help Mark Up"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   5640
      Width           =   9255
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   7080
      Picture         =   "Form1.frx":22BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "[MSN] - File Splitter"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":2629
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9375
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":2C0D
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////
'////////////////////////File Splitter Ver1.00////////////////
'/////////////////////////////////////////////////////////////
'////////////////////////By Marco Samy ---2003////////////////
'//////////////////////// marco_s2@hotmail.com////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
Dim Ox, oY
Private Const MAX_PATH = 260
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const ILD_TRANSPARENT = &H1
Private Type SHFILEINFO 'Structure used by SHGetFileInfo
hIcon As Long
iIcon As Long
dwAttributes As Long
szDisplayName As String * MAX_PATH
szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "Shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private shinfo As SHFILEINFO
Private xPixels As Integer
Private yPixels As Integer
Dim Dont As Boolean
Dim FileSize As Double, UsedSize As Double
Dim ActiveI As Integer
Function FileInsert()
Slider1.Value = 1
Slider1_Click
Label27.Caption = "Min " & SpaceValue(1 / Slider1.Max * FileSize)
Label26.Caption = "Max " & SpaceValue(FileSize)
For I = 1 To Part.UBound
Unload Part(I)
Unload Title(I)
Next I
UsedSize = 0
End Function
Private Sub Combo1_Change()
If Dont = True Then Dont = False: Exit Sub
On Error Resume Next
If Val(OriginalValue(Combo1.Text)) < 1 Then Combo1.Text = 1
If Val(OriginalValue(Combo1.Text)) > Val(FileSize) Then
MsgBox "Unit size is larger than the file size.", vbCritical
Text2.Text = 1
Label25_Click
Else
Dim Osv
Osv = FileSize / OriginalValue(Combo1.Text)
If Int(Osv) > Osv Then Osv = Int(Osv) + 1 Else Osv = Int(Osv)
Text2.Text = Osv
End If
Combo1.Text = SpaceValue(OriginalValue(Combo1.Text))
If InList(Combo1.Text, Combo1) = False Then Combo1.AddItem Combo1.Text
End Sub
Private Sub Combo1_Click()
On Error Resume Next
If Val(OriginalValue(Combo1.Text)) > Val(FileSize) Then
MsgBox "Unit size is larger than the file size.", vbCritical
Text2.Text = 1
Label25_Click
Else
Dim Osv
Osv = FileSize / OriginalValue(Combo1.Text)
If Int(Osv) < Osv Then Osv = Int(Osv) + 1 Else Osv = Int(Osv)
Text2.Text = Osv
End If
End Sub
Private Sub Combo1_LostFocus()
On Error Resume Next
If Val(OriginalValue(Combo1.Text)) > Val(FileSize) Then
MsgBox "Unit size is larger than the file size.", vbCritical
Text2.Text = 1
Label25_Click
Else
Dim Osv
Osv = FileSize / OriginalValue(Combo1.Text)
If Int(Osv) < Osv Then Osv = Int(Osv) + 1 Else Osv = Int(Osv)
Text2.Text = Osv
End If
Combo1.Text = SpaceValue(OriginalValue(Combo1.Text))
If InList(Combo1.Text, Combo1) = False Then Combo1.AddItem Combo1.Text
End Sub
Private Sub Form_Click()
Picture1.Visible = False
End Sub
Private Sub Form_Load()
Image5.Picture = Image4.Picture
Image6.Picture = Image4.Picture
Image7.Picture = Image4.Picture
If App.PrevInstance = True Then MsgBox "Another copy is already running.", vbCritical: End
Combo2.Text = "Byte"
Pb.Appearance = ccFlat
End Sub
Private Sub Image1_Click()
Picture1.Visible = Not Picture1.Visible
On Error Resume Next
Picture1.SetFocus
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Visible = False
Ox = x: oY = y
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then Move Left + x - Ox, Top + y - oY
End Sub
Private Sub Label10_Click()
Picture1.Visible = False
frmAB.Show 1, Me
End Sub
Private Sub Label14_Click()
Picture1.Visible = False
If ActiveI = 0 Then MsgBox "Select Part First", vbCritical: Exit Sub
Dim cUse As Double, Tax As Double, Usage As Double
cUse = FileSize - UsedSize + Val(Title(ActiveI).Tag)
Tax = Title(ActiveI).Tag
Usage = Tax
Form3.ResizeIt Me, cUse, Tax
Usage = Tax - Usage
Dim nWid As Single
Dim FreeWid As Single, FreeStart As Single, GAdd As Single, MyWid As Single, MyLeft As Single
GAdd = Shape3.Left + Screen.TwipsPerPixelX
FreeStart = GAdd + ((UsedSize - (Title(ActiveI).Tag)) / FileSize * Shape3.Width)
FreeWid = Shape3.Width - 60 - FreeStart + GAdd
MyWid = (Tax / (FileSize - UsedSize + Title(ActiveI).Tag) * FreeWid) - 60
If MyWid > 2 Then
nWid = MyWid - Part(ActiveI).Width
Part(ActiveI).Width = MyWid
Title(ActiveI).Width = MyWid
Title(ActiveI).Tag = Tax
Title(ActiveI).Caption = SpaceValue(cUse)
For I = ActiveI + 1 To Part.UBound
Part(I).Left = Part(I).Left + nWid
Title(I).Left = Title(I).Left + nWid
Next I
UsedSize = UsedSize + Usage
If UsedSize > FileSize Then UsedSize = FileSize
Slider1_Change
End If
End Sub
Private Sub Label15_Click()
Label22_Click
Picture1.Visible = False
End Sub
Private Sub Label16_Click()
Picture1.Visible = False
Hide1.Visible = False
End Sub
Private Sub Label17_Click()
Hide1.Visible = True
End Sub
Private Sub Label18_Click()
If Val(Text1.Text) < 0 Then Exit Sub
If InList(Val(Text1.Text) & Combo2.Text, Combo2) = False Then Combo1.AddItem Val(Text1.Text) & Combo2.Text
Combo1.Text = Val(Text1.Text) & Combo2.Text
End Sub
Function InList(Item As Variant, sList As Object) As Boolean
InList = False
For I = 0 To sList.ListCount
If sList.List(I) = Item Then InList = True: GoTo ExitIt
Next I
ExitIt:
End Function
Private Sub Label19_Click()
Picture1.Visible = False
If ActiveI = 0 Then MsgBox "Select Part First", vbCritical: Exit Sub
UsedSize = UsedSize - Title(ActiveI).Tag
For I = ActiveI To Title.UBound - 1
Title(I).Caption = Title(I + 1).Caption
Title(I).ToolTipText = Title(I + 1).ToolTipText
Part(I).Width = Part(I + 1).Width
If Not I = 1 Then Part(I).Left = Part(I - 1).Left + Part(I - 1).Width + 60 + (2 * Screen.TwipsPerPixelX)
Title(I).Width = Title(I + 1).Width
If Not I = 1 Then Title(I).Left = Title(I - 1).Left + Title(I - 1).Width + 60 + (2 * Screen.TwipsPerPixelX)
Title(I).Tag = Title(I + 1).Tag
Next I
Unload Title(Title.UBound)
Unload Part(Part.UBound)
Label21.Caption = "Parts is : " & Part.UBound & " Part."
Slider1_Change
If UsedSize < 0 Then UsedSize = 0
ActiveI = 0
End Sub
Private Sub Label2_Click()
Timer1_Timer
Picture1.Visible = False
On Error GoTo Err1
With Dlg
Begin:
On Error GoTo Err1
.CancelError = True
If Label3.Caption = "Select File" Then .FileName = GetSetting(App.EXEName, "Old", "File") Else .FileName = Label3.Caption
.ShowOpen
If Not Trim$(.FileName) = "" Then
SaveSetting App.EXEName, "Old", "File", .FileName
Label3.Caption = Trim$(.FileName)
Label9.Caption = "Size is : " & SpaceValue(FileLen(Label3.Caption))
FileSize = FileLen(Label3.Caption)
If FileSize = 0 Then
MsgBox "You Have seleted 0 Bytes File, Please Select Another."
GoTo Begin
End If
If Val(FileSize) < 500 Then Slider1.Max = FileSize
GetFileIcon
Combo1_LostFocus
End If
End With
Err1:
Err.Clear
P1.Enabled = Option1.Value
P2.Enabled = Option2.Value
FileInsert
If FileSize = 0 Then
MsgBox "You Have seleted 0 Bytes File, Please Select Another."
GoTo Begin
End If
End Sub
'Get Space Taken in B,KB,MB.GB ...
Function SpaceValue(ByVal sVal As Double) As String
Dim nVal As Double, Level, xStr
nVal = Fix(sVal)
Level = 0
While nVal > 1024
nVal = nVal / 1024
Level = Level + 1
Wend
nVal = Val(Format$(CStr(nVal), "###.##"))
Select Case Level
Case 0: xStr = "Byte"
Case 1: xStr = "KB"
Case 2: xStr = "MB"
Case 3: xStr = "GB"
Case 4: xStr = "TB"
Case 5: xStr = "QB"
End Select
SpaceValue = CStr(nVal) & xStr
End Function
Private Sub Label20_Click()
Picture1.Visible = False
AddPart Slider1.Value / Slider1.Max
End Sub
Function AddPart(ByVal PCNT As Single)
On Error GoTo PointF
Dim FreeWid As Single, FreeStart As Single, GAdd As Single, MyWid As Single, MyLeft As Single
GAdd = Shape3.Left + Screen.TwipsPerPixelX
FreeStart = GAdd + (UsedSize / FileSize * Shape3.Width)
FreeWid = Shape3.Width - 60 - FreeStart + GAdd
MyWid = (PCNT * FreeWid) - 60 - Screen.TwipsPerPixelX
MyLeft = FreeStart + 60
Load Part(Part.UBound + 1)
Load Title(Part.UBound)
Part(Part.UBound).Move MyLeft, Part(0).Top, MyWid
Title(Part.UBound).Move MyLeft, Title(0).Top, MyWid
Part(Part.UBound).Visible = True
Title(Part.UBound).Visible = True
Part(Part.UBound).ZOrder 0
Title(Part.UBound).ZOrder 0
Title(Part.UBound).Caption = SpaceValue(PCNT * (FileSize - UsedSize))
Title(Part.UBound).ToolTipText = "Part " & Part.UBound & ", Size " & SpaceValue(PCNT * (FileSize - UsedSize))
Title(Part.UBound).Tag = (PCNT * (FileSize - UsedSize))
UsedSize = UsedSize + (PCNT * (FileSize - UsedSize))
Label21.Caption = "Parts is : " & Part.UBound & " Part."
Slider1_Change
Exit Function
PointF:
MsgBox "Size Of the part is Small Alot.", vbCritical
On Error Resume Next
If (Part(Part.UBound).Visible = False) And (Part.UBound > 0) Then Unload Part(Part.UBound)
If (Title(Title.UBound).Visible = False) And (Title.UBound > 0) Then Unload Title(Title.UBound)
End Function
Private Sub Label22_Click()
Picture1.Visible = False
Dim Msg As String
If Label3.Caption = "Select File" Then
Msg = "First Select File by Clicking Browse " & vbCrLf & "After , Choose if you want to split (Part) in Normal Parts or not by choosing it's own button , Then Click Help Again."
Else
If Option1.Value = True Then Msg = "Select Your Unit Space, Or Add New One ..." & vbCrLf & "Or Parts Number and Click Set ..." & vbCrLf & "Clcik Begin." Else Msg = "Add Your Parts ..." & vbCrLf & "Edit Them With Delets and Resize ..." & vbCrLf & "Clcik Begin."
End If
MsgBox Msg, vbInformation
End Sub
Function GetFileIcon()
IM.ListImages.Clear
Dim hIcon, himl As Long
himl = SHGetFileInfo(Label3.Caption, 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
PicT.Cls
ImageList_Draw himl, shinfo.iIcon, PicT.hdc, 0, 0, ILD_TRANSPARENT
DestroyIcon shinfo.iIcon
IM.ListImages.Add , , PicT.Image
Image2.Picture = IM.ListImages.Item(1).ExtractIcon
End Function
Private Sub Label25_Click()
Picture1.Visible = False
If Val(Text2.Text) > Val(FileSize) Then MsgBox "Number is larger than file size": Text2.Text = 1
If Val(Text2.Text) < 1 Then MsgBox "Number is too small": Text2.Text = 1
Combo1.Text = SpaceValue(FileSize / Val(Text2.Text))
If Val(Text2.Text) = 1 Then GoTo Skip:
While Val(OriginalValue(Combo1.Text) * Val(Text2.Text)) < Val(FileSize)
Dont = True
Combo1.Text = Val(Combo1.Text) + 0.01 & Right$(Combo1.Text, Len(Combo1.Text) - Len(CStr(Val(Combo1.Text))))
Wend
Skip:
If InList(Combo1.Text, Combo1) = False Then Combo1.AddItem Combo1.Text
End Sub
'get the Real Bytes on a space value
Function OriginalValue(sValue As String) As Double
Dim NumLen
NumLen = Len(CStr(Val(sValue)))
Dim StrLevel As String, mLevel As Integer, Oval As Double
StrLevel = Right$(sValue, Len(sValue) - NumLen)
Select Case Trim$(StrLevel)
Case "Byte": mLevel = 0
Case "KB": mLevel = 1
Case "MB": mLevel = 2
Case "GB": mLevel = 3
Case "TB": mLevel = 4
Case "QB": mLevel = 5
End Select
Oval = Val(sValue)
For I = 1 To mLevel
Oval = Oval * 1024
Next I
OriginalValue = Oval
End Function
Private Sub Label28_Click()
Label10_Click
Picture1.Visible = False
End Sub
Private Sub Label29_Click()
Label4_Click
Picture1.Visible = False
End Sub
Private Sub Label3_DblClick()
Label2_Click
End Sub
Private Sub Label33_Click()
Label16_Click
End Sub
Private Sub Label4_Click()
End
End Sub
Private Sub Label5_Click()
Picture1.Visible = False
WindowState = vbMinimized
End Sub
Private Sub Label6_Click()
Pb.Value = Pb.Min
Pb.Visible = True
Picture1.Visible = False
On Error GoTo Err1
Dim FileSave As String, NamePart As String, ExtPart As String, CurInt As Integer, TrueNum As Double
With Dlg
.CancelError = True
.ShowSave
FileSave = .FileName
End With
On Error Resume Next
NamePart = GetBL(".", FileSave)
If NamePart = "" Then NamePart = FileSave
ExtPart = GetAL(".", FileSave)
If Not ExtPart = "" Then ExtPart = "." & ExtPart
If Option1.Value = True Then
'Split Ordinary
If Val(Text1.Text) = 1 Then
FileCopy Label3.Caption, NamePart & "#1" & ExtPart
Else
TrueNum = FileSize / OriginalValue(Combo1.Text)
If Int(TrueNum) < TrueNum Then TrueNum = Int(TrueNum) + 1
OrderSplit Label3.Caption, NamePart, ExtPart, TrueNum
End If
'/End Split
Else
'Non Order Split
If UsedSize < FileSize Then If MsgBox("Parts dosen't match the end of the file, do you really want to continue paring the file?", vbCritical + vbYesNo) = vbNo Then Exit Sub
Dim sCol As New Collection
ReSetUnits sCol
NonOrderSplit Label3.Caption, NamePart, ExtPart, sCol
'/End Split
End If
Err1:
Pb.Visible = False
End Sub
Function ReSetUnits(mColl As Collection)
Dim Sum As Double, SimSum As Double
For I = 1 To Title.Count
mColl.Add Format(Title(I).Tag)
Next I
ReSum:
For I = 1 To mColl.Count - 1
SimSum = SimSum + CDbl(mColl.Item(I))
Next I
Sum = SimSum + CDbl(mColl.Item(mColl.Count))
If Sum > FileSize Then mColl.Item(mColl.Count) = FileSize - SimSum
If CDbl(mColl.Item(mColl.Count)) = 0 Then mColl.Remove mColl.Count
If CDbl(mColl.Item(mColl.Count)) < 0 Then mColl.Remove mColl.Count: GoTo ReSum
End Function
Function OrderSplit(fFrom As String, sName, sExt, sNums)
On Error GoTo Err1:
Dim fTo As String, Unit As Double, LastUnit As Double, InputFile As String, OutputFile As String
Dim nf, nf2, PatchNum As Double, PatchSize As Double
Dim LastPatch As Double, cUnit As Double, cPatch As Double, Bytes() As Byte
Unit = FileLen(fFrom) / sNums
If Int(Unit) < Unit Then Unit = Int(Unit) + 1
Calc:
LastUnit = FileLen(fFrom) - ((sNums - 1) * Unit)
If LastUnit < 0 Then sNums = sNums - 1: GoTo Calc
Pb.Max = sNums
nf = FreeFile
Open fFrom For Binary As #nf
For I = 1 To sNums
OutputFile = sName & "#" & I & sExt
'-----------------------Calc Patches
If I = sNums Then cUnit = LastUnit Else cUnit = Unit
PatchSize = MaxSize
PatchNum = cUnit / PatchSize
If Int(PatchNum) < PatchNum Then PatchNum = Int(PatchNum) + 1
If Int(PatchSize) < PatchSize Then PatchSize = Int(PatchSize) + 1
Patches:
LastPatch = cUnit - ((PatchNum - 1) * PatchSize) - 1
If LastPatch < 0 Then PatchNum = PatchNum - 1: GoTo Patches
'-----------/Calc Patches
nf2 = FreeFile
Open OutputFile For Binary As #nf2
For Z = 1 To PatchNum
If Z = PatchNum Then cPatch = LastPatch Else cPatch = PatchSize ' Select Current Patch
ReDim Bytes(cPatch)
Get #nf, ((I - 1) * Unit) + ((Z - 1) * PatchSize) + 1, Bytes()
Put #nf2, ((Z - 1) * PatchSize) + 1, Bytes()
Next Z
Close #nf2
Pb.Value = Pb.Value + 1
DoEvents
Next I
Close #nf
Exit Function
Err1:
MsgBox Err.Description, vbCritical
Resume Next
End Function
Function NonOrderSplit(fFrom As String, sName, sExt, Units As Collection)
On Error GoTo Err1:
Dim fTo As String, InputFile As String, OutputFile As String
Dim nf, nf2, PatchNum As Double, PatchSize As Double
Dim LastPatch As Double, cUnit As Double, cPatch As Double, Bytes() As Byte, LastUnits As Double
Pb.Max = Units.Count
nf = FreeFile
Open fFrom For Binary As #nf
For I = 1 To Units.Count
OutputFile = sName & "#" & I & sExt
'-----------------------Calc Patches
cUnit = Units(I)
PatchSize = MaxSize
PatchNum = cUnit / PatchSize
If Int(PatchNum) < PatchNum Then PatchNum = Int(PatchNum) + 1
If Int(PatchSize) < PatchSize Then PatchSize = Int(PatchSize) + 1
Patches:
LastPatch = cUnit - ((PatchNum - 1) * PatchSize) - 1
If LastPatch < 0 Then PatchNum = PatchNum - 1: GoTo Patches
'-----------/Calc Patches
nf2 = FreeFile
Open OutputFile For Binary As #nf2
For Z = 1 To PatchNum
If Z = PatchNum Then cPatch = LastPatch Else cPatch = PatchSize ' Select Current Patch
ReDim Bytes(cPatch)
Get #nf, LastUnits + ((Z - 1) * PatchSize) + 1, Bytes()
Put #nf2, ((Z - 1) * PatchSize) + 1, Bytes()
Next Z
Close #nf2
Pb.Value = Pb.Value + 1
DoEvents
LastUnits = LastUnits + cUnit
Next I
Close #nf
Exit Function
Err1:
MsgBox Err.Description, vbCritical
Resume Next
End Function
Private Sub Label7_Click()
Picture1.Visible = False
Form2.Show 1, Me
End Sub
Private Sub Label8_Click()
Label5_Click
Picture1.Visible = False
End Sub
Private Sub Option1_Click()
If (P1.Enabled = False) And (P2.Enabled = False) Then Exit Sub
P2.BackColor = vbBlack
P1.BackColor = BackColor
P2.Enabled = False
P1.Enabled = True
Hide1.BackColor = P1.BackColor
End Sub
Private Sub Option2_Click()
If (P1.Enabled = False) And (P2.Enabled = False) Then Exit Sub
P1.BackColor = vbBlack
P2.BackColor = BackColor
P2.Enabled = True
P1.Enabled = False
Hide1.BackColor = P1.BackColor
End Sub
Private Sub Picture1_LostFocus()
Picture1.Visible = False
End Sub
Private Sub Slider1_Change()
Label23.Caption = "Adding Part In " & SpaceValue((FileSize - UsedSize) * Slider1.Value / Slider1.Max) & " Byte"
Label24.Caption = "is a " & Int(Slider1.Value / Slider1.Max * 100) & "% Of File Free Space."
End Sub
Private Sub Slider1_Click()
Label23.Caption = "Adding Part In " & SpaceValue((FileSize - UsedSize) * Slider1.Value / Slider1.Max) & " Byte"
Label24.Caption = "is a " & Int(Slider1.Value / Slider1.Max * 100) & "% Of File Free Space."
End Sub
Private Sub State_Click()
Picture1.Visible = False
End Sub
Private Sub Timer1_Timer()
Picture2.Visible = False
Timer1.Enabled = False
End Sub
Private Sub Title_Click(Index As Integer)
For I = 1 To Part.UBound
Part(I).BorderColor = vbYellow
Next I
Part(Index).BorderColor = vbBlue
ActiveI = Index
End Sub
