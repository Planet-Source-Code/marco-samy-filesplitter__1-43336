VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar Pb 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.PictureBox PicT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin ComctlLib.ListView List1 
      Height          =   3735
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Add Files By Drag Or (Add) Button"
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16711935
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "SS"
         Object.Tag             =   ""
         Text            =   "File Name And Path"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "GG"
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "State"
         Object.Width           =   1763
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3720
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mark"
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Select a space from selected part (add a part only)"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Be Sure that files is placed in right order"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   855
      Left            =   5280
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin ComctlLib.ImageList IM 
      Left            =   4320
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   327682
   End
   Begin VB.Label State 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "By Marco Samy - Status : Ready"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      MouseIcon       =   "Form2.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Start Collecting Files in 1 file"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Left            =   120
      MouseIcon       =   "Form2.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Down"
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Move Selected Down"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Up"
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Move Selected Up"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Insert"
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Insert File ( Before ) Selected Item."
      Top             =   1500
      Width           =   1335
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Remove Selected Item"
      Top             =   540
      Width           =   1335
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":093E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Add New Parts"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim SizeAll As Double
Function AddFile(sEntry, Optional cIdx As Single, Optional cUse As String)
On Error GoTo Err1
If FileLen(sEntry) = 0 Then MsgBox "File Lenght is 0 Byte will not be added.", vbCritical:    Exit Function
If List1.ListItems.Count = 0 Then SizeAll = 0
Dim hIcon, himl As Long
himl = SHGetFileInfo(sEntry, 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
PicT.Cls
ImageList_Draw himl, shinfo.iIcon, PicT.hdc, 0, 0, ILD_TRANSPARENT
DestroyIcon shinfo.iIcon
IM.ListImages.Add , , PicT.Image
Set List1.SmallIcons = IM
If Val(cIdx) = 0 Then
List1.ListItems.Add IM.ListImages.Count, , sEntry, , IM.ListImages.Count
List1.ListItems(List1.ListItems.Count).SubItems(1) = Form1.SpaceValue(FileLen(sEntry))
List1.ListItems(List1.ListItems.Count).SubItems(2) = "Full File"
Else
List1.ListItems.Add cIdx, , sEntry, , IM.ListImages.Count
List1.ListItems(cIdx).SubItems(1) = Form1.SpaceValue(FileLen(sEntry))
List1.ListItems(cIdx).SubItems(2) = cUse
End If
SizeAll = SizeAll + FileLen(sEntry)
State.Caption = "Full Size is: " & Form1.SpaceValue(SizeAll) & " , in " & List1.ListItems.Count & " File ."
Err1:
End Function
Private Sub Form_Load()
State.RightToLeft = True
List1.ColumnHeaders(1).Width = 0.7 * List1.Width
List1.ColumnHeaders(2).Width = List1.Width - List1.ColumnHeaders(1).Width - 680
End Sub
Private Sub Label1_Click()
On Error GoTo Err1
With Dlg
If Label3.Caption = "Select File" Then .FileName = GetSetting(App.EXEName, "Old", "File") Else .FileName = Label3.Caption
.CancelError = True
.ShowOpen
If Not Trim$(.FileName) = "" Then
SaveSetting App.EXEName, "Old", "File", .FileName
AddFile .FileName, List1.SelectedItem.Index
End If
End With
Err1:
End Sub
Private Sub Label19_Click()
On Error Resume Next
State.Caption = "Full Size is: " & Form1.SpaceValue(SizeAll) & " , in " & List1.ListItems.Count & " File ."
RemItem (List1.SelectedItem.Index)
End Sub
Function RemItem(ByVal sIdx As Single)
On Error Resume Next
SizeAll = SizeAll - GetUsage(sIdx)
List1.ListItems.Remove (sIdx)
DoRef
End Function
Private Sub Label2_Click()
On Error GoTo Ext
Dim cIdx As Single, cFile As String, oUse As String
cIdx = List1.SelectedItem.Index
oUse = List1.ListItems(cIdx).SubItems(2)
If cIdx = 1 Then Exit Sub
If Val(cIdx) = 0 Then Exit Sub
cFile = List1.SelectedItem.Text
RemItem (cIdx)
DoRef
AddFile cFile, cIdx - 1, oUse
List1.ListItems.Item(cIdx - 1).Selected = True
Ext:
End Sub
Function GetUsage(sIdx As Single) As Double
Dim uStr As String
uStr = List1.ListItems(sIdx).SubItems(2)
If uStr = "Full File" Then
GetUsage = FileLen(List1.ListItems(sIdx).Text)
Else
Dim nSt, nED
nSt = GetBF("-", uStr, 1)
nED = GetAL("-", uStr)
GetUsage = Val(nED) - Val(nED) + 1
End If
End Function
Private Sub Label20_Click()
Adds.Show 1, Me
End Sub
Function DoRef()
Dim hColl As New Collection
Dim hColl2 As New Collection
For I = 1 To List1.ListItems.Count
hColl.Add List1.ListItems(I).Text
hColl2.Add List1.ListItems(I).SubItems(2)
Next I
Set List1.SmallIcons = Nothing
IM.ListImages.Clear
List1.ListItems.Clear
For I = 1 To hColl.Count
AddFile hColl.Item(I), , hColl2.Item(I)
Next I
End Function
Private Sub Label3_Click()
On Error GoTo Ext
Dim cIdx As Single, cFile As String, oUse As String
cIdx = List1.SelectedItem.Index
oUse = List1.ListItems(cIdx).SubItems(2)
If cIdx = List1.ListItems.Count Then Exit Sub
If Val(cIdx) = 0 Then Exit Sub
cFile = List1.SelectedItem.Text
RemItem (cIdx)
AddFile cFile, cIdx + 1, oUse
List1.ListItems.Item(cIdx + 1).Selected = True
Ext:
End Sub
Private Sub Label4_Click()
Unload Me
End Sub
Function BeginCollect(OutputFile)
On Error GoTo Err1
Pb.Value = Pb.Min
Pb.Visible = True
Pb.Max = List1.ListItems.Count
Dim InputFile As String, Init As Double, Finish As Double, OutMarker As Double, InMarker As Double, Ab() As Byte, cUnit As Double, PatchNum As Double, PatchSize As Double, LastPatch As Double, cPatch As Double, Past As Double, nf, nf2
nf = FreeFile
Open OutputFile For Binary As #nf
nf2 = FreeFile
For I = 1 To List1.ListItems.Count
If List1.ListItems(I).SubItems(2) = "Full File" Then InMarker = 1: OutMarker = FileLen(List1.ListItems(I).Text) Else InMarker = Val(GetBL("-", List1.ListItems(I).SubItems(2))): OutMarker = Val(GetAL("-", List1.ListItems(I).SubItems(2)))
InputFile = List1.ListItems(I).Text
'-----------------------Calc Patches
cUnit = OutMarker - InMarker + 1
PatchSize = MaxSize
PatchNum = cUnit / PatchSize
If Int(PatchNum) < PatchNum Then PatchNum = Int(PatchNum) + 1
If Int(PatchSize) < PatchSize Then PatchSize = Int(PatchSize) + 1
Patches:
LastPatch = cUnit - ((PatchNum - 1) * PatchSize) - 1
If LastPatch <= 0 Then PatchNum = PatchNum - 1: GoTo Patches
'-----------/Calc Patches
Open InputFile For Binary As #nf2
For Z = 1 To PatchNum
If Z = PatchNum Then cPatch = LastPatch Else cPatch = PatchSize ' Select Current Patch
ReDim Ab(cPatch)
Get #nf2, InMarker + ((Z - 1) * PatchSize), Ab()
Put #nf, Past + ((Z - 1) * PatchSize) + 1, Ab()
Next Z
Close #nf2
Past = Past + (OutMarker - InMarker + 1)
Pb.Value = Pb.Value + 1
DoEvents
Next I
Close #nf
Err1:
Pb.Visible = False
End Function
Private Sub Label5_Click()
On Error GoTo Err1
If List1.ListItems.Count = 0 Then Exit Sub
With Dlg
.CancelError = True
.FileName = GetSetting(App.EXEName, "Old", "File")
.ShowSave
BeginCollect (.FileName)
End With
Err1:
End Sub
Private Sub Label7_Click()
On Error GoTo Ext
Dim nSt As Double, nEnd As Double, MyRet As String, oUsage As Double, nUsage As Double, Addins As Double
If List1.SelectedItem.SubItems(2) = "Full File" Then nSt = 1: nEnd = FileLen(List1.SelectedItem.Text) Else nSt = CDbl(GetBF("-", List1.SelectedItem.SubItems(2), 1)): nEnd = CDbl(GetAL("-", List1.SelectedItem.SubItems(2)))
oUsage = nEnd - nSt + 1
Form3.Costumize List1.SelectedItem.Text, Me, nSt, nEnd
nUsage = nEnd - nSt + 1
Addins = nUsage - oUsage
SizeAll = SizeAll + Addins
State.Caption = "Full Size is: " & Form1.SpaceValue(SizeAll) & " , in " & List1.ListItems.Count & " File ."
Select Case nSt
Case Is = nEnd: MyRet = "Full File"
Case 1
If nEnd = FileLen(List1.SelectedItem.Text) Then MyRet = "Full File" Else MyRet = nSt & "-" & nEnd
Case Else: MyRet = nSt & "-" & nEnd
End Select
List1.SelectedItem.SubItems(2) = MyRet
Ext:
End Sub

'The Following Code is disabled by (Rem) Mark
'It's Job is Sorting the list by Clicked Column in any list box, but it isn't needed in our program
Private Sub List1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'List1.SortKey = (ColumnHeader.Index - 1)
'List1.Sorted = True
End Sub

Private Sub List1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(15) = True Then
For I = 1 To Data.Files.Count
AddFile (Data.Files.Item(I))
Next I
End If
End Sub
