Attribute VB_Name = "DirFile"
'//////////////////////////////////////////////////////////////////////////
'/////////////////File And Folder Control Module For Magic Copy
'/////////////////By Marco Samy 2002
'//////////////////////////////////////////////////////////////////////////
Public Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDrives Lib "Kernel32" () As Long
Public Declare Function GetLogicalDriveStrings Lib "Kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Function GetDirs(sPath As String, sColl As Collection)
On Error Resume Next
Dim Dirx As String
Dirx = Dir(NormPath(sPath), vbDirectory + vbHidden + vbReadOnly + vbSystem + vbArchive)
While Not Dirx = ""
If Not Left$(Dirx, 1) = "." Then sColl.Add NormPath(sPath) & Dirx
Dirx = Dir
Wend
End Function
Function GetFiles(sPath As String, sColl As Collection)
On Error Resume Next
Dim Dirx As String
Dirx = Dir(NormPath(sPath), vbNormal + vbHidden + vbReadOnly + vbSystem + vbArchive)
While Not Dirx = ""
If Not Left$(Dirx, 1) = "." Then sColl.Add NormPath(sPath) & Dirx
Dirx = Dir
Wend
End Function
Function NormPath(sPath As String) As String
If Right$(sPath, 1) = "\" Then NormPath = sPath Else NormPath = sPath & "\"
End Function

