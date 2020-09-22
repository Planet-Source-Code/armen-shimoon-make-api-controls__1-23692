Attribute VB_Name = "Module1"
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const SW_SHOWNORMAL = 1
Public Const WS_CHILD = &H40000000
Public Const WS_BORDER = &H800000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CW_USEDEFAULT = &H80000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Dim a1 As String




Sub MakeButton(hwnd1 As String, x1 As Integer, y1 As Integer, w1 As Integer, h1 As Integer, caption As String)



a1 = FindWindow(vbNullString, hwnd1)
    
    If a1 = "0" Then
        Exit Sub
        MsgBox "Error: " & Err.Description
    Else


b& = CreateWindowEx(0&, "Button", caption, WS_CHILD, x1, y1, w1, h1, a1, 0&, 0&, 0&)


Call ShowWindow(b&, SW_SHOWNORMAL)
End If


End Sub


Sub MakeStatic(hwnd1 As String, x1 As Integer, y1 As Integer, w1 As Integer, h1 As Integer, caption As String)
a1 = FindWindow(vbNullString, hwnd1)
    
    If a1 = "0" Then
        Exit Sub
        MsgBox "Error: " & Err.Description
    Else

c& = CreateWindowEx(0&, "Static", caption, WS_CHILD, x1, y1, w1, h1, a1, 0&, 0&, 0&)


Call ShowWindow(c&, SW_SHOWNORMAL)

End If

End Sub

Sub MakeEdit(hwnd1 As String, x1 As Integer, y1 As Integer, w1 As Integer, h1 As Integer, caption As String)
a1 = FindWindow(vbNullString, hwnd1)
    
    If a1 = "0" Then
        Exit Sub
        MsgBox "Error: " & Err.Description
    Else

a& = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", caption, WS_CHILD, x1, y1, w1, h1, a1, 0&, 0&, O&)

Call ShowWindow(a&, SW_SHOWNORMAL)
End If

End Sub
