Attribute VB_Name = "LogMeInTransferModule"
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Sub ClearClipboard()
 On Error Resume Next
 OpenClipboard 0&
 EmptyClipboard
 CloseClipboard
End Sub


