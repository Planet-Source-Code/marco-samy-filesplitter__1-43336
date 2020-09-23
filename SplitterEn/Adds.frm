VERSION 5.00
Begin VB.Form Adds 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Item"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "< Add"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Add Selected Drive to the list"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "< Add"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Add Selected File(s)"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change To"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< Add"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Add Selected Folder to the list"
      Top             =   720
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1980
      Hidden          =   -1  'True
      Left            =   120
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4575
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   "Press Ctrl For Multi Select"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   4800
      X2              =   6000
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   6000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   6000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      Caption         =   "Add Selected Folder"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Add Selected File(s)."
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Close The Screen"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "To Show the folder's contents"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Double Click To Add The File"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click To Add The Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   4800
      X2              =   6000
      Y1              =   3975
      Y2              =   3975
   End
End
Attribute VB_Name = "Adds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////Add File(Multi Select) Or Folder-For Magic Copy////////////////////
Private Sub Command1_Click()
Dim IC As New Collection
GetFiles Dir1.Path, IC
For I = 1 To IC.Count
Form2.AddFile IC.Item(I)
Next I
End Sub
Private Sub Command2_Click()
Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub
Private Sub Command3_Click()
For I = 0 To File1.ListCount - 1
If File1.Selected(I) = True Then Form2.AddFile NormPath(Dir1.Path) & File1.List(I)
Next I
End Sub
Private Sub Command4_Click()
Dim IC As New Collection
'GetDirs Left$(Dir1, 2), IC
GetFiles Left$(Dir1.Path, 2), IC
For I = 1 To IC.Count
Form2.AddFile IC.Item(I)
Next I
End Sub
Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_DblClick()
For I = 0 To File1.ListCount - 1
If File1.Selected(I) = True Then Form2.AddFile NormPath(Dir1.Path) & File1.List(I)
Next I
End Sub
Private Sub Form_Load()
WindowOnTop Me
End Sub
Function NormPath(ByVal sPath As String) As String
If Right$(sPath, 1) = "\" Then NormPath = sPath Else NormPath = sPath & "\"
End Function
