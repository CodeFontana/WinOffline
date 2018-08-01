'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/KeyboardHook
' File Name:    KeyboardHook.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Partial Public Class WinOffline

    Private Shared Function KbHook_KeyDown(ByVal Key As System.Windows.Forms.Keys) As Boolean
        Dim bShift As Boolean = My.Computer.Keyboard.ShiftKeyDown
        Dim bCaps As Boolean = My.Computer.Keyboard.CapsLock
        Select Case CInt(Key)
            Case 48 To 57
                Dim ArrNum As Char() = {")", "!", "@", "#", "$", "%", "^", "&", "*", "("}
                KeyboardHook.CapturedString.Append(If(bShift = True, ArrNum(CInt(Key) - 48), CInt(Key) - 48))
                Console.Write("*")
                Return True
            Case 65 To 90
                If (bShift = False And bCaps = False) Or (bShift = True And bCaps = True) Then
                    KeyboardHook.CapturedString.Append(Key.ToString.ToLower)
                ElseIf (bShift = False And bCaps = True) Or (bShift = True And bCaps = False) Then
                    KeyboardHook.CapturedString.Append(Key.ToString.ToUpper)
                End If
                Console.Write("*")
                Return True
            Case 96 To 105
                KeyboardHook.CapturedString.Append(CInt(Key) - 96)
                Console.Write("*")
                Return True
            Case 186 To 192
                Dim ArrSymbShift As Char() = {":", "+", "<", "_", ">", "?", "~"}
                Dim ArrSymb As Char() = {";", "=", ",", "-", ".", "/", "`"}
                KeyboardHook.CapturedString.Append(If(bShift = True, ArrSymbShift(CInt(Key) - 186), ArrSymb(CInt(Key) - 186)))
                Console.Write("*")
                Return True
            Case 219 To 222
                Dim ArrSymbShift As Char() = {"{", "|", "}", """"}
                Dim ArrSymb As Char() = {"[", "\", "]", "'"}
                KeyboardHook.CapturedString.Append(If(bShift = True, ArrSymbShift(CInt(Key) - 219), ArrSymb(CInt(Key) - 219)))
                Console.Write("*")
                Return True
            Case Else
                Select Case Key
                    Case Keys.Back
                        If KeyboardHook.CapturedString.Length > 0 Then
                            KeyboardHook.CapturedString.Length -= 1
                            Console.Write(ChrW(8) & ChrW(32) & ChrW(8))
                            Return True
                        Else
                            Return True
                        End If
                    Case Keys.Return
                        KeyboardHook.UnSetHook()
                        Return True
                    Case Keys.Space
                        KeyboardHook.CapturedString.Append(" ")
                    Case Keys.Divide
                        KeyboardHook.CapturedString.Append("/")
                    Case Keys.Decimal
                        KeyboardHook.CapturedString.Append(".")
                    Case Keys.Subtract
                        KeyboardHook.CapturedString.Append("-")
                    Case Keys.Add
                        KeyboardHook.CapturedString.Append("+")
                    Case Keys.Multiply
                        KeyboardHook.CapturedString.Append("*")
                    Case Else
                        Return False
                End Select
                Console.Write("*")
                Return True
        End Select
        Return False
    End Function

    Public Class KeyboardHook

        Private Const HC_ACTION As Integer = 0
        Private Const WH_KEYBOARD_LL As Integer = 13
        Private Const WM_KEYDOWN = &H100
        Private Const WM_KEYUP = &H101
        Private Const WM_SYSKEYDOWN = &H104
        Private Const WM_SYSKEYUP = &H105

        Private Structure KBDLLHOOKSTRUCT
            Public vkCode As Integer
            Public scancode As Integer
            Public flags As Integer
            Public time As Integer
            Public dwExtraInfo As Integer
        End Structure

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetWindowsHookEx(
            <[In]()> ByVal idHook As Integer,
            <[In]()> ByVal lpfn As KeyboardProcDelegate,
            <[In]()> ByVal hmod As IntPtr,
            <[In]()> dwThreadId As Integer) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function CallNextHookEx(
            <[In]()> ByVal hHook As Integer,
            <[In]()> ByVal nCode As Integer,
            <[In]()> ByVal wParam As Integer,
            <[In]()> ByVal lParam As KBDLLHOOKSTRUCT) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function UnhookWindowsHookEx(
            <[In]()> ByVal hHook As Integer) As Integer
        End Function

        Private Delegate Function KeyboardProcDelegate(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer

        Public Shared Event KeyDown(ByVal Key As Keys)

        Public Shared KeyboardHooked As Boolean = False
        Public Shared CapturedString As System.Text.StringBuilder

        Private Shared KeyHook As Integer
        Private Shared KeyHookDelegate As KeyboardProcDelegate

        Public Shared Sub SetHook()
            CapturedString = New System.Text.StringBuilder
            KeyHookDelegate = New KeyboardProcDelegate(AddressOf KeyboardProc)
            KeyHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyHookDelegate, Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)), 0)
            KeyboardHooked = True
        End Sub

        Public Shared Sub UnSetHook()
            UnhookWindowsHookEx(KeyHook)
            KeyboardHooked = False
        End Sub

        Private Shared Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer
            If (nCode = HC_ACTION) Then
                Select Case wParam
                    Case WM_KEYDOWN, WM_SYSKEYDOWN
                        If KbHook_KeyDown(CType(lParam.vkCode, Keys)) Then Return IntPtr.op_Explicit(1)
                End Select
            End If
            Return CallNextHookEx(KeyHook, nCode, wParam, lParam)
        End Function

    End Class

End Class