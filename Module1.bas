Attribute VB_Name = "Module1"
Option Explicit

 Declare Function GetShortPathName Lib "kernel32" _
   Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
   ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
   As Long


Public Function GetShortFileName(ByVal FullPath As String) _
  As String
Dim lAns As Long
Dim sAns As String
Dim iLen As Integer
   
On Error Resume Next

'this function doesn't work if the file doesn't exist
If Dir(FullPath) = "" Then Exit Function

sAns = Space(255)
lAns = GetShortPathName(FullPath, sAns, 255)
GetShortFileName = Left(sAns, lAns)
    
End Function

