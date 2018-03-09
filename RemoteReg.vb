Public Class RemoteReg

    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer

    Private Declare Ansi Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Declare Ansi Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer

    Private Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Integer) As Integer
    Private Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
    Private Declare Function GetSubMenu Lib "user32.dll" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer

    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101

    Private Const WM_SETTEXT As Integer = &HC
    Private Const WM_COMMAND As Integer = &H111


    Public Sub SelectComputer(ByVal strComputer As String)

        Dim dtStartTime As DateTime
        Dim ElapsedTime As TimeSpan

        If (System.Diagnostics.Process.GetProcessesByName("Regedit").Length) = 0 Then
            Process.Start("regedit")
        End If

        Dim ret As Integer

        Dim hWnd As Integer = 0
        Dim hMainMenu As Integer
        Dim hMenu As Integer
        Dim MenuID As Integer

        'Get Regedit window 
        dtStartTime = Now
        Do While hWnd = 0
            hWnd = FindWindow(vbNullString, "Registry Editor")
            If hWnd = 0 Then
                System.Threading.Thread.Sleep(500)
                ElapsedTime = Now.Subtract(dtStartTime)
                If ElapsedTime.TotalSeconds > 5 Then
                    Console.WriteLine("Could not get a handle on the Registry Editor window.")
                    Exit Sub
                End If
            End If
        Loop


        'Get menu strip handle
        hMainMenu = GetMenu(hWnd)

        'Get the first menu
        hMenu = GetSubMenu(hMainMenu, 0)

        'Get the menu item for Connect Network Registry
        'It's the 7th menu item counting separators
        MenuID = GetMenuItemID(hMenu, 6)

        'Use PostMessage to activate the menu item
        'because SendMessage would wait until the dialog closed...
        ret = PostMessage(hWnd, WM_COMMAND, MenuID, Nothing)


        dtStartTime = Now

        Dim dialogWnd As Integer = 0
        Do While dialogWnd = 0
            dialogWnd = FindWindowEx(0, 0, "#32770", "Select Computer")
            If dialogWnd = 0 Then
                System.Threading.Thread.Sleep(500)
                ElapsedTime = Now.Subtract(dtStartTime)
                If ElapsedTime.TotalSeconds > 5 Then
                    Console.WriteLine("Could not get a handle on the Select Computer window.")
                    Exit Sub
                End If
            End If
        Loop

        'Change caption
        'ret = SendMessage(childWnd, WM_SETTEXT, vbNullString, "hello world")

        Dim textboxHandle As Integer
        textboxHandle = FindWindowEx(dialogWnd, vbNullString, "RichEdit50W", vbNullString)
        If textboxHandle = 0 Then
            textboxHandle = FindWindowEx(dialogWnd, vbNullString, "RichEdit20W", vbNullString)
        End If
        ret = SendMessage(textboxHandle, WM_SETTEXT, vbNullString, strComputer)

        Dim buttonHandle As Integer = FindWindowEx(dialogWnd, vbNullString, "Button", "OK")
        SendMessage(buttonHandle, WM_KEYDOWN, Asc(" "), vbNullString)
        SendMessage(buttonHandle, WM_KEYUP, Asc(" "), vbNullString)

        Process.Start("regedit")
    End Sub

End Class
