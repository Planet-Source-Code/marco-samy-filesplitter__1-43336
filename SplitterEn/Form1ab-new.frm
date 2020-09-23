VERSION 5.00
Begin VB.Form frmAB 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1920
   End
   Begin VB.PictureBox Pu 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   1
      Top             =   0
      Width           =   6255
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   2160
      Left            =   1560
      Picture         =   "Form1ab-new.frx":0000
      Top             =   1200
      Width           =   1860
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6240
      MousePointer    =   7  'Size N S
      Picture         =   "Form1ab-new.frx":174D
      Stretch         =   -1  'True
      ToolTipText     =   "ÊÛííÑ ÓÑÚÉ ÇáÊÍÑíß"
      Top             =   1680
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   440
      X2              =   424
      Y1              =   0
      Y2              =   232
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   424
      X2              =   424
      Y1              =   0
      Y2              =   232
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   1800
      Top             =   3120
      Width           =   6255
   End
End
Attribute VB_Name = "frmAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cX, oY
Private Sub Form_Click()
Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Unload Me
End Sub
Private Sub Form_Load()
WindowOnTop Me
P1.Width = Image1.Width
Pu.Width = P1.Width
Dim Str
'//////////////////////////////////////////////////'
''Writting Code Here
'//////////////////////////////////////////////////'
With P1
Str = " "
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
.FontSize = 36
.AutoSize = True
.AutoRedraw = True
P1.Height = Image2.Height + 140
P1.PaintPicture Image2.Picture, ((P1.Width - Image2.Width) / 2), P1.CurrentY
P1.CurrentY = P1.CurrentY + 140
P1.Height = CurrentY + 200
Str = "File Splitter"
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
'-next Statment
Str = "Version 1.00e"
.FontSize = 24
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbGreen
.Height = .CurrentY + 100
P1.Print Str
'-next Statment
Str = "(Free Unit File Splitter)"
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
'-next Statment
Str = "Now You are free!"
.FontSize = 8
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
'-next Statment
Str = "  "
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
'-next Statment
Str = "COPYRIGHT (c) 12/2002"
.FontSize = 26
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
Str = "Desgin & Programming"
.FontSize = 20
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
Str = "Marco Samy Nasif"
.FontSize = 22
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbYellow
.Height = .CurrentY + 100
P1.Print Str
Str = "Mail : marco_s2@hotmail.com"
.FontSize = 16
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = "Call (+20) 012 7242974"
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
Str = "IDEA"
.FontSize = 20
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
Str = "Mohammed Taha Youns"
.FontSize = 22
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbYellow
.Height = .CurrentY + 100
P1.Print Str
Str = "Call (+20) 012 3306936"
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = "Mail: mtmaybasha77@hotmail.com"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
'
Str = "Just Drag Your Files/Folder Onto Program"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
Str = "Select Draged Files You Want To Use"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
Str = "Split Them in your way"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
Str = "and control your data."
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
Str = "This Program is 1 File Needed, Only EXE File"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print
Str = "You Can Copy it onto your desktop"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
Str = "Or to your friends directley[FREEWARE]"
.FontSize = 14
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 24
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = &HFF00FF
.Height = .CurrentY + 100
P1.Print Str
'
Str = "Technosoft Egypt CO."
.FontSize = 12
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbGreen
.Height = .CurrentY + 100
P1.Print Str
Str = "El-Minia"
.FontSize = 20
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 36
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbGreen
.Height = .CurrentY + 100
P1.Print Str
Str = "Any Other brands written inside or outside the program"
.FontSize = 12
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = "in the files related to it are either trade marks or"
.FontSize = 12
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = "registered trade marks of their owners."
.FontSize = 12
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 16
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = "Technosoft is a v. co."
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = "and refers to Marco Samy."
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = "MSN refers to Marco Samy."
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = "for more information send to mail up"
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = "or call me on 0127242974 (+20 outside Egypt.)"
.FontSize = 18
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbBlue
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 16
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
Str = "Thank you for using Technosoft Products"
.FontSize = 12
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbWhite
.Height = .CurrentY + 100
P1.Print Str
Str = " "
.FontSize = 20
.CurrentX = (.Width - .TextWidth(Str)) / 2
.ForeColor = vbRed
.Height = .CurrentY + 100
P1.Print Str
'/End Writing
'///////////////////////////////////////////
'///////////////////////////////////////////
Image1.Picture = .Image
.AutoRedraw = False
.AutoSize = False
.Height = (Height / Screen.TwipsPerPixelY) - Pu.Height
.Cls
End With
Timer1.Interval = 1 + (100 * (Image3.Top / ((Height / Screen.TwipsPerPixelY) - Image3.Height)))
Timer1_Timer
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
oY = y
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
xy = y / Screen.TwipsPerPixelY
xoY = oY / Screen.TwipsPerPixelY
If Val(Image3.Top + xy - xoY) < 0 Then Image3.Top = 0: GoTo Skip
If Val(Image3.Top + xy - xoY + Image3.Height) > Val(Height / Screen.TwipsPerPixelY) Then Image3.Top = (Height / Screen.TwipsPerPixelY) - Image3.Height: GoTo Skip
Image3.Top = Image3.Top + xy - xoY
Skip:
Timer1.Interval = 10 + (200 * (Image3.Top / ((Height / Screen.TwipsPerPixelY) - Image3.Height)))
End If
End Sub
Private Sub P1_Click()
Unload Me
End Sub
Private Sub Pu_Click()
Unload Me
End Sub
Private Sub Timer1_Timer()
cX = cX + 2
If Val(cX + P1.Height - 2) > Val(Image1.Height) Then cX = 0: Sleep (1000)
P1.PaintPicture Image1.Picture, 0, 0, P1.Width, P1.Height, 0, cX, , P1.Height
Pu.PaintPicture Image1.Picture, 0, 0, Pu.Width, Pu.Height, 0, cX, , 1
End Sub
