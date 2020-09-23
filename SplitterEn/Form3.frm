VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   LinkTopic       =   "Form3"
   ScaleHeight     =   3255
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Mid 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      ScaleHeight     =   615
      ScaleWidth      =   7095
      TabIndex        =   3
      Top             =   1480
      Width           =   7095
      Begin VB.PictureBox xR 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   6840
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   0
         Width           =   255
         Begin VB.Image xRi 
            Height          =   615
            Left            =   0
            MousePointer    =   9  'Size W E
            Picture         =   "Form3.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox xL 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   0
         Width           =   255
         Begin VB.Image xLi 
            Height          =   615
            Left            =   0
            MousePointer    =   9  'Size W E
            Picture         =   "Form3.frx":036F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0 KB"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   510
         Left            =   240
         Top             =   60
         Width           =   6615
      End
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " <<"
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
      Left            =   120
      MouseIcon       =   "Form3.frx":06DE
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Set To Min"
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
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
      Left            =   7200
      MouseIcon       =   "Form3.frx":0830
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Set To Max"
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Use From 0 To 0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Size : 0.0 KB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   120
      Picture         =   "Form3.frx":0982
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OK"
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
      Left            =   6120
      MouseIcon       =   "Form3.frx":1DB6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Accept"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   240
      MouseIcon       =   "Form3.frx":1F08
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Select The Space That Will Be Used From Selected File/Part."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   4
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Size As Double
Public nStart As Double
Public nEnd As Double
Dim Ox As Single, AddX As Double
Private Sub Label4_Click()
nStart = 0
nEnd = 0
Unload Me
End Sub
Function Costumize(ByVal sFile As String, sForm As Form, ByRef sFrom As Double, ByRef sTo As Double)
Label3.Visible = True: xLi.Enabled = True
Label3.Caption = sFile
Size = FileLen(sFile) - 1
SetValue1 ((sFrom - 1) / Size)
SetValue2 ((sTo - 1) / Size)
Label1.Caption = Form1.SpaceValue(Size)
xLi_MouseMove 1, 1, Ox, 1
Show 1, sForm
If (nStart = 0) And (nEnd = 0) Then Exit Function
If nStart = nEnd Then If nStart < Size Then nEnd = nEnd + 1 Else nStart = nStart - 1
sFrom = nStart: sTo = nEnd
End Function
Function ResizeIt(sForm As Form, ByRef sMax As Double, ByRef sCur As Double, Optional AddiN As Double)
AddX = AddiN
Size = sMax
SetValue2 (sCur / sMax)
xRi_MouseMove 1, 1, Ox, 1
Label3.Visible = False
Label1.Caption = "Resize Your Part Below."
xLi.Enabled = False
Show 1, sForm
If nEnd = 0 Then Exit Function Else sCur = nEnd
End Function
Private Function GetValue1()
Dim MaxVal As Single
MaxVal = Mid.Width - xL.Width - xR.Width
GetValue1 = xL.Left / MaxVal
End Function
Private Function GetValue2()
Dim MaxVal As Single
MaxVal = Mid.Width - xL.Width - xR.Width
GetValue2 = (xR.Left - xR.Width) / MaxVal
End Function
Private Function SetValue1(sVal As Double)
Dim MaxVal As Single
MaxVal = Mid.Width - xL.Width - xR.Width
xL.Left = sVal * MaxVal
End Function
Private Function SetValue2(sVal As Double)
Dim MaxVal As Single
MaxVal = Mid.Width - xL.Width - xR.Width
xR.Left = (sVal * MaxVal) + xR.Width
End Function
Private Sub Label7_Click()
Unload Me
End Sub
Private Sub Label8_Click()
xR.Move Mid.Width - xR.Width
xRi_MouseMove 1, 1, Ox, 1
End Sub
Private Sub Label9_Click()
xL.Move 0
xLi_MouseMove 1, 1, Ox, 1
End Sub
Private Sub xLi_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Ox = x
End Sub
Private Sub xLi_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err1
If Val(xL.Left + x - Ox) < 0 Then Exit Sub
If Val(xL.Left + x - Ox + xL.Width + 10) >= xR.Left Then Exit Sub
If Button = 1 Then
xL.Move xL.Left + x - Ox
Shape3.Move xL.Left + xL.Width, Shape3.Top, Mid.Width - (Mid.Width - xR.Left) - xL.Left - xL.Width
nStart = CDbl(Format(GetValue1 * Size, "#"))
nEnd = CDbl(Format(GetValue2 * Size, "#"))
SetValues
Label2.Caption = "Use from : " & nStart & " To " & nEnd
Label6.Caption = Form1.SpaceValue(nEnd - nStart + 1)
End If
Exit Sub
Err1:
If nStart = 0 Then nEnd = 0
nStart = 0
Resume Next
End Sub
Private Sub xRi_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Ox = x
End Sub
Private Sub xRi_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If (nStart = 0) And (nEnd = 0) Then nStart = 1
On Error GoTo Err1
If Val(xR.Left + x - Ox + xR.Width) > Mid.Width Then Exit Sub
If Val(xR.Left + x - Ox - 10) <= (xL.Left + xL.Width) Then Exit Sub
If Button = 1 Then
xR.Move xR.Left + x - Ox
Shape3.Move xL.Left + xL.Width, Shape3.Top, Mid.Width - (Mid.Width - xR.Left) - xL.Left - xL.Width
nStart = CDbl(Format(GetValue1 * Size, "#"))
nEnd = CDbl(Format(GetValue2 * Size, "#"))
SetValues
Label2.Caption = "Use from : " & nStart & " To " & nEnd
Label6.Caption = Form1.SpaceValue(nEnd - nStart + 1)
End If
Exit Sub
Err1:
If nStart = 0 Then nEnd = 0
nStart = 0
Resume Next
End Sub
Function SetValues()
nStart = nStart + 1
nEnd = nEnd + 1
If nStart = nEnd Then If nStart < Size Then nEnd = nEnd + 1 Else nStart = nStart - 1
End Function
