Attribute VB_Name = "Internet"
'*************************************************
'* Project: Geek Companion                                                      *
'* Programmer: Gnu Kemist GnuKemist@yahoo.com                *
'* Version: 0.0.1 (as of March. 20, 2001)                                   *
'* Known Bugs: None                                                                *
'*************************************************

Option Explicit

Public Const URL = "http://members.bellatlantic.net/~vze26hdt/"
Public Const Email = "gnukemist@yahoo.com"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub GoToWeb()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

Public Sub SendEmail()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, "mailto:" & Email, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

