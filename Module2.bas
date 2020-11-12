Attribute VB_Name = "Module2"
Option Explicit

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
() '(ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
() '(ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As LongPtr) As Long

Public Sub GetWindows()
    '~~> Pass Full Name or Partial Name. This is not case sensitive
    Debug.Print GetAllWindowHandles(".xlsm")
End Sub

Private Function GetAllWindowHandles(partialName As String)
'returns handle/title of all open windows with partial name match
    Dim hwnd As Long, lngRet As Long
    Dim strText As String

    hwnd = FindWindowEx(0&, 0&, vbNullString, vbNullString)

    While hwnd <> 0
        strText = String$(100, Chr$(0))
        lngRet = GetWindowText(hwnd, strText, 100)

        If InStr(1, strText, partialName, vbTextCompare) > 0 Then
            Debug.Print "The Handle of the window is " & hwnd & " and " & vbNewLine & _
                        "The title of the window is " & Left$(strText, lngRet) & vbNewLine & _
                        "----------------------"
        End If

        '~~> Find next window
        hwnd = FindWindowEx(0&, hwnd, vbNullString, vbNullString)
    Wend
End Function



