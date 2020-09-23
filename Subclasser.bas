Attribute VB_Name = "Subclasser"
Option Explicit

Public Function GetHiWord(ByVal Value As Long) As Integer

    'Return the high word of a long value.
    RtlMoveMemory GetHiWord, ByVal VarPtr(Value) + 2, 2

End Function

Public Function GetLoWord(ByVal Value As Long) As Integer

    'Return the low word of a long value
    RtlMoveMemory GetLoWord, Value, 2

End Function

'Window hook procedure
Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim strBuffer As String, lngNbChars As Long, lngAnswer As Long

   'Gets all the events and messages for our controls and form
    Select Case wMsg
    Case WM_COMMAND
        'Click event of the button, get the text in the textbox and display it
        If lParam = hCmd And GetHiWord(wParam) = BN_CLICKED Then
            lngNbChars = GetWindowTextLength(hText)
            If lngNbChars > 0 Then
                strBuffer = String$(lngNbChars, Chr$(32))
                lngNbChars = GetWindowText(hText, strBuffer, lngNbChars + 1)
                MsgBox strBuffer
            End If
        'Click event of the first radio button
        ElseIf lParam = hRadBut1 And GetHiWord(wParam) = BN_CLICKED Then
            intRadButValue = 1
            'Tell Windows to repaint the textbox
            InvalidateRgn hText, 0&, True
        'Click event of the second radio button
        ElseIf lParam = hRadBut2 And GetHiWord(wParam) = BN_CLICKED Then
            intRadButValue = 2
            'Tell Windows to repaint the textbox
            InvalidateRgn hText, 0&, True
        'Click event of the third radio button
        ElseIf lParam = hRadBut3 And GetHiWord(wParam) = BN_CLICKED Then
            intRadButValue = 3
            'Tell Windows to repaint the textbox
            InvalidateRgn hText, 0&, True
        'Click event of the checkbox
        ElseIf lParam = hChk1 And GetHiWord(wParam) = BN_CLICKED Then
            If SendMessage(hChk1, BM_GETCHECK, 0&, 0&) = BST_CHECKED Then
                EnableWindow hText, True
                CheckMenuItem hFileSubMenu, FILETOGGLEID, MFS_CHECKED
            Else
                EnableWindow hText, False
                CheckMenuItem hFileSubMenu, FILETOGGLEID, MFS_UNCHECKED
            End If
        'Menu
        ElseIf lParam = 0 Then
            If GetHiWord(wParam) = 0 And GetLoWord(wParam) = 100 Then
                SendMessage hChk1, BM_CLICK, 0, 0
            ElseIf GetHiWord(wParam) = 0 And GetLoWord(wParam) = 102 Then
                DestroyWindow hForm
            ElseIf GetHiWord(wParam) = 0 And GetLoWord(wParam) = 2 Then
                MsgBox "All API form" + vbNewLine + "by ReadError :)", vbOKOnly, "All API Form"
            End If
        'Change event of the textbox, enable the button if length is > 0
        ElseIf lParam = hText And GetHiWord(wParam) = EN_CHANGE Then
            lngNbChars = GetWindowTextLength(hText)
            If lngNbChars > 0 Then
                'Check if the button is already enabled to prevent sending
                'unnecessary messages
                If Not blnButtonEnabled Then
                    EnableWindow hCmd, True
                    blnButtonEnabled = True
                    SendMessage hLabel1, WM_SETTEXT, 0, ByVal "You can now press the button"
                End If
            Else
                EnableWindow hCmd, False
                blnButtonEnabled = False
                SendMessage hLabel1, WM_SETTEXT, 0, ByVal "Enter some text and press the button"
            End If
        End If
    Case WM_CTLCOLOREDIT
        If lParam = hText Then
            If hBrush <> 0 Then
                DeleteObject hBrush
            End If
            If intRadButValue = 1 Then
                hBrush = CreateSolidBrush(RGB(255, 255, 255))
                SetBkColor wParam, RGB(255, 255, 255)
                WndProc = hBrush
            ElseIf intRadButValue = 2 Then
                hBrush = CreateSolidBrush(RGB(255, 255, 0))
                SetBkColor wParam, RGB(255, 255, 0)
                WndProc = hBrush
            Else
                hBrush = CreateSolidBrush(RGB(0, 255, 0))
                SetBkColor wParam, RGB(0, 255, 0)
                WndProc = hBrush
            End If
            Exit Function
        End If
    Case WM_DESTROY
        DeleteObject hBrush
        PostQuitMessage 0
    Case WM_QUIT
        DestroyWindow hForm
    End Select
    WndProc = DefWindowProc(hwnd, wMsg, wParam, lParam)

End Function
