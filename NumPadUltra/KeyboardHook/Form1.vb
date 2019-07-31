Public Class Form1
    '---Version 1.0
    'Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Integer) As Integer
    Private Declare Sub keybd_event Lib "user32" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    Private TimerT As Single, Lock As Boolean, Mode As Integer, offset As Point, Moving As Boolean
    'Time for winkey to be pressed
    Private TimeWinKey As Date, WinKeyDown As Boolean
    'Green icon for unlock; Red icon for lock; Yellow icon for Numlock
    Private Ico() = {My.Resources.ResourceManager.GetObject("Keybd"), My.Resources.ResourceManager.GetObject("KeybdFbd"), My.Resources.ResourceManager.GetObject("KeybdNum")}

    Private Function KeyBoardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        Dim newPtr As KeyboardHookStruct = CType(Runtime.InteropServices.Marshal.PtrToStructure(lParam, newPtr.GetType()), KeyboardHookStruct)
        If nCode = HC_ACTION Then 'Keyboard event is on
            If newPtr.vkCode = VK_LWin OrElse newPtr.vkCode = VK_RWin Then
                If wParam.ToInt32 = WM_KeyUP OrElse wParam.ToInt32 = WM_SYSKEYUP Then
                    WinKeyDown = False
                    'If long press (>1s) then ignore
                    If (DateTime.Now - TimeWinKey).Seconds < 1 Then
                        If newPtr.vkCode - 91 Then ChangeMode() Else ChangeLock() 'If Left WinKey is up, lock/unlock keyboard
                        SendKeys.Send("^") 'To nullify the action of winkey
                    End If
                ElseIf wParam.ToInt32 = WM_KEYDOWN OrElse wParam.ToInt32 = WM_SYSKEYDOWN Then
                    If Not WinKeyDown Then TimeWinKey = DateTime.Now : WinKeyDown = True
                End If
            Else
                If Lock Then
                    KeyBoardProc = 1 'Disable input
                    If Mode = 0 Then
                        Select Case newPtr.vkCode
                            Case Keys.OemMinus : NumPad("&H6D", wParam)
                            Case Keys.O : NumPad("&H67", wParam)
                            Case Keys.P : NumPad("&H68", wParam)
                            Case Keys.OemOpenBrackets : NumPad("&H69", wParam)
                            Case Keys.OemCloseBrackets : NumPad("&H6B", wParam)
                            Case Keys.OemPipe : NumPad("&H6F", wParam)
                            Case Keys.L : NumPad("&H64", wParam)
                            Case Keys.OemSemicolon : NumPad("&H65", wParam)
                            Case Keys.OemQuotes : NumPad("&H66", wParam)
                            Case Keys.Oemcomma : NumPad("&H61", wParam)
                            Case Keys.OemPeriod : NumPad("&H62", wParam)
                            Case Keys.OemQuestion : NumPad("&H63", wParam)
                            Case Keys.PrintScreen : NumPad("00", wParam)
                            Case Keys.RControlKey : NumPad("&H6E", wParam)
                            Case Keys.RMenu : NumPad("&H60", wParam)
                            Case Keys.Home : NumPad("&H5D", wParam)
                            Case Keys.End : NumPad("&H13", wParam)
                            Case Else : KeyBoardProc = 0
                        End Select
                    End If
                    If KeyBoardProc = 1 Then Exit Function
                End If
            End If
        End If
        Return CallNextHookEx(m_iKeyHandle, nCode, wParam, lParam) 'Call back other responses
    End Function

    Private Sub NumPad(ByVal Key As String, ByVal wParam As IntPtr)
        If wParam.ToInt32 = WM_KEYDOWN OrElse wParam.ToInt32 = WM_SYSKEYDOWN Then
            If Strings.Left(Key, 2) = "&H" Then
                keybd_event(Val(Key), 0, 0, 0)
                keybd_event(Val(Key), 0, &H2, 0)
            Else
                keybd_event(Asc("0"), 0, 0, 0)
                keybd_event(Asc("0"), 0, &H2, 0)
                keybd_event(Asc("0"), 0, 0, 0)
                keybd_event(Asc("0"), 0, &H2, 0)
            End If
        End If
    End Sub

    Private Sub SetHook()
        fix_COCD = New HookCallBack(AddressOf KeyBoardProc)
        m_iKeyHandle = SetWindowsHookEx(WH_KEYBOARD_LL, fix_COCD, GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName), 0)
        If m_iKeyHandle <> 0 Then
            Me.Icon = Ico(2 - Mode)
            If Mode Then Me.Text = "LockAll" Else Me.Text = "NumLock"
            Me.BackgroundImage = Me.Icon.ToBitmap
        Else 'SetHook failure
            ToolTip2.Show("... occurred when trying to lock keyboard.", Label1, 5000)
        End If
    End Sub

    Private Sub UnHook()
        If m_iKeyHandle <> 0 Then UnhookWindowsHookEx(m_iKeyHandle)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Form2.Show()
        Form2.ToolTip1.Show(Form2.ToolTip1.GetToolTip(Form2.Label19), Form2.Label19)
        e.Cancel = True 'Forbid closing directly
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = 64
        Me.Height = 64
        'Saved starting position
        Me.Left = My.Settings.x
        Me.Top = My.Settings.y
        Lock = True
        SetHook()
        '---Version 1.0
        'If BlockInput(True) = 0 Then MsgBox("Please run as administrator.", MsgBoxStyle.Exclamation) : End
    End Sub

    Private Sub Form1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        'Left click: changeLock; Right click: Hide; Middle: Move window to (0,0)
        If Moving Then Exit Sub
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ChangeMode()
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
            ChangeLock()
        ElseIf e.Button = Windows.Forms.MouseButtons.Middle Then
            Form2.Show()
        End If
    End Sub

    Private Sub Form1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDoubleClick
        'Double click to quit
        If e.Button = Windows.Forms.MouseButtons.Left Then
            UnHook()
            End
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            ChangeMode() 'Undo change
            Me.Hide()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Opacity setting
        If Moving Then Me.Opacity = 0.3 : Exit Sub
        If MousePosition.X > Me.Left + 8 AndAlso MousePosition.X < Me.Left + 52 AndAlso MousePosition.Y > Me.Top + 8 AndAlso MousePosition.Y < Me.Top + 52 Then
            If TimerT < 0.5 Then TimerT += 0.1 Else Exit Sub
            If TimerT = 0.5 Then ToolTip1.Show(ToolTip1.GetToolTip(Me), Me, 5000) 'Show tooltip even if the window is deactivated
        Else
            ToolTip1.Hide(Me)
            If TimerT > 0 Then TimerT -= 0.1 Else Exit Sub
        End If
        Me.Opacity = TimerT + 0.5
    End Sub

    Private Sub ChangeLock() 'Lock/unlock keyboard
        Lock = Not Lock
        If Lock Then
            Me.Icon = Ico(2 - Mode)
            If Mode Then Me.Text = "LockAll" Else Me.Text = "NumLock"
        Else
            Me.Icon = Ico(0)
            Me.Text = "Standby"
        End If
        Me.BackgroundImage = Me.Icon.ToBitmap
        TimerT = 0.6 'To avoid showing the tooltip balloon
        Me.Opacity = 1
        Me.Visible = True
    End Sub

    Private Sub ChangeMode() 'Lock/unlock keyboard
        Mode = 1 - Mode
        Lock = False
        ChangeLock()
    End Sub

    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        offset.X = MousePosition.X - Me.Left
        offset.Y = MousePosition.Y - Me.Top
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        'Moving window
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Moving = True
            ToolTip1.Hide(Me) 'No showing tooltip
            Me.Left = MousePosition.X - offset.X
            Me.Top = MousePosition.Y - offset.Y
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        'End of moving window
        If Moving Then
            Moving = False
            TimerT = 0
            'Undo if the change if too small (to avoid misoperation)
            If Math.Abs(My.Settings.x - Me.Left) < 16 AndAlso Math.Abs(My.Settings.y - Me.Top) < 16 Then
                Me.Left = My.Settings.x
                Me.Top = My.Settings.y
                ChangeLock()
            Else
                'Save settings of starting position
                'C:\Users\[username]\AppData\Local\[companyname=Zehao_Sun]\...
                My.Settings.x = Me.Left
                My.Settings.y = Me.Top
                My.Settings.Save()
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            If Math.Abs(MousePosition.X - Me.Left - offset.X) > 16 OrElse Math.Abs(MousePosition.Y - Me.Top - offset.Y) > 16 Then
                Me.Left = 0 : Me.Top = 0
                My.Settings.x = Me.Left
                My.Settings.y = Me.Top
                My.Settings.Save()
            Else
                'ChangeMode()
            End If
        End If
    End Sub
End Class
