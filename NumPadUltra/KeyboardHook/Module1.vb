Module Module1
    'Callbackhandle, Callbackfunction
    Public m_iKeyHandle As Integer = 0
    Public m_clsHookCallback As HookCallBack
    'Event happened (if nCode < 0, then this value must be sent to *CallNextHook*)
    Public Const HC_ACTION As Integer = 0
    'Virtual-key code
    Public Const VK_LWin As Integer = &H5B
    Public Const VK_RWin As Integer = &H5C
    'Hookid
    Public Const WH_KEYBOARD_LL = 13
    'wPawam
    Public Const WM_KEYDOWN As Integer = &H100
    Public Const WM_KeyUP As Integer = &H101
    Public Const WM_SYSKEYDOWN As Integer = &H104
    Public Const WM_SYSKEYUP As Integer = &H105
    'Callbackfunction
    Public fix_COCD As HookCallBack 'Define reference variable to avoid *CallbackOnCollectedDelegate* error
    'Callback
    Public Delegate Function HookCallBack(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    'API
    Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal ModuleName As String) As IntPtr
    Public Declare Function GetLastError Lib "kernel32" Alias "GetLastError" () As Integer
    Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal HookProc As HookCallBack, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    Public Declare Function CallNextHookEx Lib "user32" (ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal idHook As Integer) As Boolean

    Public Structure KeyboardHookStruct
        Public vkCode As Integer 'Specifies a virtual-key code. The code must be a value in the range 1 to 254. 
        Public scanCode As Integer 'Specifies a hardware scan code for the key.
        Public flags As Integer 'Specifies the extended-key flag, event-injected flag, context code, and transition-state flag.
        Public time As Integer 'Specifies the time stamp for this message.
        Public dwExtraInfo As Integer 'Specifies extra information associated with the message.
    End Structure
End Module
