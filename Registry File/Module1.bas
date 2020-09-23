Attribute VB_Name = "Module1"
Option Explicit
 

'Shell32.dll for send e-mail
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'Kernel32.dll writing registry file
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
      "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Function WriteRegFile(ByVal sRegFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
   Dim i As Integer
   On Error GoTo sWriteRegFileError

   i = WritePrivateProfileString(sSection, sItem, sText, sRegFileName)
   WriteRegFile = True

   Exit Function
sWriteRegFileError:
   WriteRegFile = False
End Function


Public Sub FileSave(Text As String, FilePath As String)
'Save string "REGEDIT4" (Why??? I don't now, but this string must be writing in first line)
On Error Resume Next
Dim Directory As String
              Directory$ = FilePath
   
       Open Directory$ For Output As #1
           Print #1, Text
       Close #1
Exit Sub

End Sub


Public Sub DeleteFile(FilePath As String)
'Delete a file
On Error Resume Next
Kill FilePath$
Exit Sub

End Sub
