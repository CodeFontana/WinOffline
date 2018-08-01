'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WindowsAPI
' File Name:    WindowsAPI.vb
' Author:       Brian Fontana
'***************************************************************************/

' Imports
Imports Microsoft.Win32.SafeHandles
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Security.Principal

Public Class WindowsAPI

    '******************** WINDOWS API FUNCTION IMPORTS ********************
    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function DuplicateToken(
        <[In]()> ByVal ExistingTokenHandle As SafeTokenHandle,
        <[In]()> ByVal ImpersonationLevel As SECURITY_IMPERSONATION_LEVEL,
        <Out()> ByRef DuplicateTokenHandle As SafeTokenHandle) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function DuplicateTokenEx(
        <[In]()> ByVal hExistingToken As IntPtr,
        <[In]()> ByVal dwDesiredAccess As UInteger,
        <[In]()> ByRef lpTokenAttributes As SECURITY_ATTRIBUTES,
        <[In]()> ByVal ImpersonationLevel As SECURITY_IMPERSONATION_LEVEL,
        <[In]()> ByVal TokenType As TOKEN_TYPE,
        <Out()> ByRef phNewToken As SafeTokenHandle) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetSidSubAuthority(
        <[In]()> ByVal pSid As IntPtr,
        <[In]()> ByVal nSubAuthority As UInt32) As IntPtr
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetTokenInformation(
        <[In]()> ByVal hToken As SafeTokenHandle,
        <[In]()> ByVal tokenInfoClass As TOKEN_INFORMATION_CLASS,
        <Out()> ByVal pTokenInfo As IntPtr,
        <[In]()> ByVal tokenInfoLength As Integer,
        <Out()> ByRef returnLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function OpenProcessToken(
        <[In]()> ByVal hProcess As IntPtr,
        <[In]()> ByVal desiredAccess As UInt32,
        <Out()> ByRef hToken As SafeTokenHandle) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function DeviceIoControl(
        <[In]()> ByVal hDevice As IntPtr,
        <[In]()> ByVal dwIoControlCode As UInt32,
        <[In]()> ByVal InBuffer As IntPtr,
        <[In]()> ByVal nInBufferSize As Integer,
        <[In]()> ByVal OutBuffer As IntPtr,
        <[In]()> ByVal nOutBufferSize As Integer,
        <Out()> ByRef pBytesReturned As Integer,
        <[In]()> ByVal lpOverlapped As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CreateFile(
        <[In]()> ByVal lpFileName As String,
        <[In]()> ByVal dwDesiredAccess As EFileAccess,
        <[In]()> ByVal dwShareMode As EFileShare,
        <[In]()> ByVal lpSecurityAttributes As IntPtr,
        <[In]()> ByVal dwCreationDisposition As ECreationDisposition,
        <[In]()> ByVal dwFlagsAndAttributes As EFileAttributes,
        <[In]()> ByVal hTemplateFile As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function FindWindowEx(
        <[In]()> ByVal parentHandle As IntPtr,
        <[In]()> ByVal childAfter As IntPtr,
        <[In]()> ByVal lclassName As String,
        <[In]()> ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetClientRect(
        <[In]()> ByVal hWnd As System.IntPtr,
        <[Out]()> ByRef lpRECT As RECT) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SendMessage(
        <[In]()> ByVal hWnd As IntPtr,
        <[In]()> ByVal Msg As UInteger,
        <[In]()> ByVal wParam As IntPtr,
        <[In]()> ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function MoveFileEx(
        <[In]()> ByVal lpExistingFileName As String,
        <[In]()> ByVal lpNewFileName As String,
        <[In]()> ByVal dwFlags As MoveFileFlags) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function AttachConsole(
        <[In]()> ByVal dwProcessId As Int32) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function FreeConsole() As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function GetConsoleWindow() As IntPtr
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function AdjustTokenPrivileges(
        <[In]()> ByVal TokenHandle As SafeTokenHandle,
        <[In](), MarshalAs(UnmanagedType.Bool)> ByVal DisableAllPrivileges As Boolean,
        <[In]()> ByRef NewState As TOKEN_PRIVILEGES,
        <[In]()> ByVal BufferLengthInBytes As UInt32,
        <Out()> ByRef PreviousState As TOKEN_PRIVILEGES,
        <Out()> ByRef ReturnLengthInBytes As UInt32) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CreateProcessAsUser(
        <[In]()> ByVal hToken As SafeTokenHandle,
        <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal lpApplicationName As String,
        <[In](), Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal lpCommandLine As String,
        <[In]()> ByRef lpProcessAttributes As SECURITY_ATTRIBUTES,
        <[In]()> ByRef lpThreadAttributes As SECURITY_ATTRIBUTES,
        <[In](), MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandles As Boolean,
        <[In]()> ByVal dwCreationFlags As UInteger,
        <[In]()> ByVal lpEnvironment As IntPtr,
        <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal lpCurrentDirectory As String,
        <[In]()> ByRef lpStartupInfo As STARTUPINFO,
        <Out()> ByRef lpProcessInformation As PROCESS_INFORMATION) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function LookupPrivilegeValue(
        <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal lpSystemName As String,
        <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal lpName As String,
        <Out()> ByRef lpLuid As LUID) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, EntryPoint:="WTSGetActiveConsoleSessionId", SetLastError:=True)>
    Private Shared Function WTSGetActiveConsoleSessionId() As UInt32
    End Function

    <DllImport("Wtsapi32.dll", CharSet:=CharSet.Auto, EntryPoint:="WTSQueryUserToken", SetLastError:=True)>
    Private Shared Function WTSQueryUserToken(
        <[In]()> ByVal SessionId As UInteger,
        <Out()> ByRef phToken As SafeTokenHandle) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, EntryPoint:="SetErrorMode", SetLastError:=True)>
    Public Shared Function SetErrorMode(
        <[In]()> ByVal uMode As ErrorModes) As ErrorModes
    End Function
    '**********************************************************************

    '******************** STRUCTURES ********************
    <StructLayout(LayoutKind.Sequential)>
    Private Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As System.UInt32
        Public dwThreadId As System.UInt32
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SECURITY_ATTRIBUTES
        Public nLength As System.UInt32
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure STARTUPINFO
        Public cb As System.UInt32
        Public lpReserved As String
        Public lpDesktop As String
        Public lpTitle As String
        Public dwX As System.UInt32
        Public dwY As System.UInt32
        Public dwXSize As System.UInt32
        Public dwYSize As System.UInt32
        Public dwXCountChars As System.UInt32
        Public dwYCountChars As System.UInt32
        Public dwFillAttribute As System.UInt32
        Public dwFlags As System.UInt32
        Public wShowWindow As Short
        Public cbReserved2 As Short
        Public lpReserved2 As IntPtr
        Public hStdInput As IntPtr
        Public hStdOutput As IntPtr
        Public hStdError As IntPtr
    End Structure

    Private Structure LUID
        Public LowPart As UInt32
        Public HighPart As Integer
    End Structure

    Private Structure LUID_AND_ATTRIBUTES
        Public Luid As LUID
        Public Attributes As Integer
    End Structure

    Private Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As UInt32
        <MarshalAs(UnmanagedType.ByValArray)> Public Privileges() As LUID_AND_ATTRIBUTES
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SID_AND_ATTRIBUTES
        Public Sid As IntPtr
        Public Attributes As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure TOKEN_ELEVATION
        Public TokenIsElevated As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure TOKEN_MANDATORY_LABEL
        Public Label As SID_AND_ATTRIBUTES
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure REPARSE_DATA_BUFFER
        ' Reparse point tag. Must be a Microsoft reparse point tag.
        Public ReparseTag As UInt32

        ' Size, in bytes, of the data after the Reserved member. This can be calculated by:
        ' (4 * sizeof(ushort)) + SubstituteNameLength + PrintNameLength + (namesAreNullTerminated ? 2 * sizeof(char) : 0)
        Public ReparseDataLength As UShort

        ' Reserved; do not use. 
        Public Reserved As UShort

        ' Offset, in bytes, of the substitute name string in the PathBuffer array.
        Public SubstituteNameOffset As UShort

        ' Length, in bytes, of the substitute name string. If this string is null-terminated, SubstituteNameLength does not include space for the null character.
        Public SubstituteNameLength As UShort

        ' Offset, in bytes, of the print name string in the PathBuffer array.
        Public PrintNameOffset As UShort

        ' Length, in bytes, of the print name string. If this string is null-terminated, PrintNameLength does not include space for the null character. 
        Public PrintNameLength As UShort

        ' A buffer containing the unicode-encoded path string. The path string contains the substitute name string and print name string.
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=&H3FF0)>
        Public PathBuffer As Byte()
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Private _Left As Integer, _Top As Integer, _Right As Integer, _Bottom As Integer

        Public Sub New(ByVal Rectangle As Rectangle)
            Me.New(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Sub
        Public Sub New(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer)
            _Left = Left
            _Top = Top
            _Right = Right
            _Bottom = Bottom
        End Sub

        Public Property X As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Right = _Right - _Left + value
                _Left = value
            End Set
        End Property
        Public Property Y As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = _Bottom - _Top + value
                _Top = value
            End Set
        End Property
        Public Property Left As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Left = value
            End Set
        End Property
        Public Property Top As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Top = value
            End Set
        End Property
        Public Property Right As Integer
            Get
                Return _Right
            End Get
            Set(ByVal value As Integer)
                _Right = value
            End Set
        End Property
        Public Property Bottom As Integer
            Get
                Return _Bottom
            End Get
            Set(ByVal value As Integer)
                _Bottom = value
            End Set
        End Property
        Public Property Height() As Integer
            Get
                Return _Bottom - _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = value + _Top
            End Set
        End Property
        Public Property Width() As Integer
            Get
                Return _Right - _Left
            End Get
            Set(ByVal value As Integer)
                _Right = value + _Left
            End Set
        End Property
        Public Property Location() As Point
            Get
                Return New Point(Left, Top)
            End Get
            Set(ByVal value As Point)
                _Right = _Right - _Left + value.X
                _Bottom = _Bottom - _Top + value.Y
                _Left = value.X
                _Top = value.Y
            End Set
        End Property
        Public Property Size() As Size
            Get
                Return New Size(Width, Height)
            End Get
            Set(ByVal value As Size)
                _Right = value.Width + _Left
                _Bottom = value.Height + _Top
            End Set
        End Property

        Public Shared Widening Operator CType(ByVal Rectangle As RECT) As Rectangle
            Return New Rectangle(Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height)
        End Operator
        Public Shared Widening Operator CType(ByVal Rectangle As Rectangle) As RECT
            Return New RECT(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Operator
        Public Shared Operator =(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Rectangle1.Equals(Rectangle2)
        End Operator
        Public Shared Operator <>(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Not Rectangle1.Equals(Rectangle2)
        End Operator

        Public Overrides Function ToString() As String
            Return "{Left: " & _Left & "; " & "Top: " & _Top & "; Right: " & _Right & "; Bottom: " & _Bottom & "}"
        End Function

        Public Overloads Function Equals(ByVal Rectangle As RECT) As Boolean
            Return Rectangle.Left = _Left AndAlso Rectangle.Top = _Top AndAlso Rectangle.Right = _Right AndAlso Rectangle.Bottom = _Bottom
        End Function
        Public Overloads Overrides Function Equals(ByVal [Object] As Object) As Boolean
            If TypeOf [Object] Is RECT Then
                Return Equals(DirectCast([Object], RECT))
            ElseIf TypeOf [Object] Is Rectangle Then
                Return Equals(New RECT(DirectCast([Object], Rectangle)))
            End If

            Return False
        End Function
    End Structure
    '**********************************************************************

    '******************** CONSTANTS ********************
    Private Const STANDARD_RIGHTS_REQUIRED As UInt32 = &HF0000
    Private Const STANDARD_RIGHTS_READ As UInt32 = &H20000
    Private Const TOKEN_ASSIGN_PRIMARY As UInt32 = 1
    Private Const TOKEN_DUPLICATE As UInt32 = 2
    Private Const TOKEN_IMPERSONATE As UInt32 = 4
    Private Const TOKEN_QUERY As UInt32 = 8
    Private Const TOKEN_QUERY_SOURCE As UInt32 = &H10
    Private Const TOKEN_ADJUST_PRIVILEGES As UInt32 = &H20
    Private Const TOKEN_ADJUST_GROUPS As UInt32 = &H40
    Private Const TOKEN_ADJUST_DEFAULT As UInt32 = &H80
    Private Const TOKEN_ADJUST_SESSIONID As UInt32 = &H100
    Private Const TOKEN_READ As UInt32 = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
    Private Const TOKEN_ALL_ACCESS As UInt32 = (STANDARD_RIGHTS_REQUIRED Or
        TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or
        TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or
        TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT Or TOKEN_ADJUST_SESSIONID)

    Private Const ERROR_INSUFFICIENT_BUFFER As Int32 = 122

    Private Const SE_ASSIGNPRIMARYTOKEN_NAME As String = "SeAssignPrimaryTokenPrivilege"
    Private Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"
    Private Const SE_TCB_NAME As String = "SeTcbPrivilege"
    Private Const SE_PRIVILEGE_ENABLED As UInt32 = &H2

    Public Const SECURITY_MANDATORY_UNTRUSTED_RID As Integer = 0
    Public Const SECURITY_MANDATORY_LOW_RID As Integer = &H1000
    Public Const SECURITY_MANDATORY_MEDIUM_RID As Integer = &H2000
    Public Const SECURITY_MANDATORY_HIGH_RID As Integer = &H3000
    Public Const SECURITY_MANDATORY_SYSTEM_RID As Integer = &H4000

    ' The file or directory is not a reparse point.
    Private Const ERROR_NOT_A_REPARSE_POINT As Integer = 4390
    ' The reparse point attribute cannot be set because it conflicts with an existing attribute.
    Private Const ERROR_REPARSE_ATTRIBUTE_CONFLICT As Integer = 4391
    ' The data present in the reparse point buffer is invalid.
    Private Const ERROR_INVALID_REPARSE_DATA As Integer = 4392
    ' The tag present in the reparse point buffer is invalid.
    Private Const ERROR_REPARSE_TAG_INVALID As Integer = 4393
    ' There is a mismatch between the tag specified in the request and the tag present in the reparse point.
    Private Const ERROR_REPARSE_TAG_MISMATCH As Integer = 4394
    ' Command to set the reparse point data block.
    Private Const FSCTL_SET_REPARSE_POINT As Integer = &H900A4
    ' Command to get the reparse point data block.
    Private Const FSCTL_GET_REPARSE_POINT As Integer = &H900A8
    ' Command to delete the reparse point data base.
    Private Const FSCTL_DELETE_REPARSE_POINT As Integer = &H900AC
    ' Reparse point tag used to identify mount points and junction points.
    Private Const IO_REPARSE_TAG_MOUNT_POINT As UInt32 = &HA0000003UI
    ' This prefix indicates to NTFS that the path is to be treated as a non-interpreted path in the virtual file system.
    Private Const NonInterpretedPathPrefix As String = "\??\"
    '**********************************************************************

    '******************** ENUMS ********************
    Private Enum TOKEN_INFORMATION_CLASS
        TokenUser = 1
        TokenGroups
        TokenPrivileges
        TokenOwner
        TokenPrimaryGroup
        TokenDefaultDacl
        TokenSource
        TokenType
        TokenImpersonationLevel
        TokenStatistics
        TokenRestrictedSids
        TokenSessionId
        TokenGroupsAndPrivileges
        TokenSessionReference
        TokenSandBoxInert
        TokenAuditPolicy
        TokenOrigin
        TokenElevationType
        TokenLinkedToken
        TokenElevation
        TokenHasRestrictions
        TokenAccessInformation
        TokenVirtualizationAllowed
        TokenVirtualizationEnabled
        TokenIntegrityLevel
        TokenUIAccess
        TokenMandatoryPolicy
        TokenLogonSid
        MaxTokenInfoClass
    End Enum

    Private Enum TOKEN_ELEVATION_TYPE
        TokenElevationTypeDefault = 1
        TokenElevationTypeFull
        TokenElevationTypeLimited
    End Enum

    Private Enum SECURITY_IMPERSONATION_LEVEL
        SecurityAnonymous = 0
        SecurityIdentification
        SecurityImpersonation
        SecurityDelegation
    End Enum

    Private Enum TOKEN_TYPE
        TokenPrimary = 1
        TokenImpersonation = 2
    End Enum

    Public Enum ErrorModes As UInteger
        SYSTEM_DEFAULT = &H0
        SEM_FAILCRITICALERRORS = &H1
        SEM_NOALIGNMENTFAULTEXCEPT = &H4
        SEM_NOGPFAULTERRORBOX = &H2
        SEM_NOOPENFILEERRORBOX = &H8000
    End Enum

    Private Enum EFileAccess
        GenericRead = &H80000000
        GenericWrite = &H40000000
        GenericExecute = &H20000000
        GenericAll = &H10000000
    End Enum

    Private Enum EFileShare
        None = &H0
        Read = &H1
        Write = &H2
        Delete = &H4
    End Enum

    Private Enum ECreationDisposition
        [New] = 1
        CreateAlways = 2
        OpenExisting = 3
        OpenAlways = 4
        TruncateExisting = 5
    End Enum

    Private Enum EFileAttributes
        [Readonly] = &H1
        Hidden = &H2
        System = &H4
        Directory = &H10
        Archive = &H20
        Device = &H40
        Normal = &H80
        Temporary = &H100
        SparseFile = &H200
        ReparsePoint = &H400
        Compressed = &H800
        Offline = &H1000
        NotContentIndexed = &H2000
        Encrypted = &H4000
        Write_Through = &H80000000
        Overlapped = &H40000000
        NoBuffering = &H20000000
        RandomAccess = &H10000000
        SequentialScan = &H8000000
        DeleteOnClose = &H4000000
        BackupSemantics = &H2000000
        PosixSemantics = &H1000000
        OpenReparsePoint = &H200000
        OpenNoRecall = &H100000
        FirstPipeInstance = &H80000
    End Enum

    Public Enum MoveFileFlags
        None = 0
        ReplaceExisting = 1
        CopyAllowed = 2
        DelayUntilReboot = 4
        WriteThrough = 8
        CreateHardlink = 16
        FailIfNotTrackable = 32
    End Enum

    Private Enum CreateProcessFlags
        DEBUG_PROCESS = &H1
        DEBUG_ONLY_THIS_PROCESS = &H2
        CREATE_SUSPENDED = &H4
        DETACHED_PROCESS = &H8
        CREATE_NEW_CONSOLE = &H10
        NORMAL_PRIORITY_CLASS = &H20
        IDLE_PRIORITY_CLASS = &H40
        HIGH_PRIORITY_CLASS = &H80
        REALTIME_PRIORITY_CLASS = &H100
        CREATE_NEW_PROCESS_GROUP = &H200
        CREATE_UNICODE_ENVIRONMENT = &H400
        CREATE_SEPARATE_WOW_VDM = &H800
        CREATE_SHARED_WOW_VDM = &H1000
        CREATE_FORCEDOS = &H2000
        BELOW_NORMAL_PRIORITY_CLASS = &H4000
        ABOVE_NORMAL_PRIORITY_CLASS = &H8000
        INHERIT_PARENT_AFFINITY = &H10000
        INHERIT_CALLER_PRIORITY = &H20000
        CREATE_PROTECTED_PROCESS = &H40000
        EXTENDED_STARTUPINFO_PRESENT = &H80000
        PROCESS_MODE_BACKGROUND_BEGIN = &H100000
        PROCESS_MODE_BACKGROUND_END = &H200000
        CREATE_BREAKAWAY_FROM_JOB = &H1000000
        CREATE_PRESERVE_CODE_AUTHZ_LEVEL = &H2000000
        CREATE_DEFAULT_ERROR_MODE = &H4000000
        CREATE_NO_WINDOW = &H8000000
        PROFILE_USER = &H10000000
        PROFILE_KERNEL = &H20000000
        PROFILE_SERVER = &H40000000
        CREATE_IGNORE_SYSTEM_DEFAULT = &H80000000
    End Enum
    '**********************************************************************

    '**********************************************************************
    Private Class SafeTokenHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid

        Private Sub New()
            MyBase.New(True)
        End Sub

        Friend Sub New(ByVal handle As IntPtr)
            MyBase.New(True)
            MyBase.SetHandle(handle)
        End Sub

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Friend Shared Function CloseHandle(ByVal handle As IntPtr) As Boolean
        End Function

        Protected Overrides Function ReleaseHandle() As Boolean
            Return SafeTokenHandle.CloseHandle(MyBase.handle)
        End Function

    End Class
    '**********************************************************************

    Public Shared Function IncreasePrivileges() As Boolean

        ' Local variables
        Dim hToken As SafeTokenHandle = Nothing
        Dim luid As LUID
        Dim NewState As TOKEN_PRIVILEGES
        NewState.PrivilegeCount = 1
        ReDim NewState.Privileges(0)

        ' Get current process token
        If OpenProcessToken(Diagnostics.Process.GetCurrentProcess.Handle, TokenAccessLevels.MaximumAllowed, hToken) = False Then

            ' Write debug
            WriteEvent("Error: Windows API OpenProcessToken function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Lookup SeIncreaseQuotaPrivilege
        If Not LookupPrivilegeValue(Nothing, SE_INCREASE_QUOTA_NAME, luid) Then

            ' Write debug
            WriteEvent("Error: Windows API LookupPrivilegeValue function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Enable SeIncreaseQuotaPrivilege
        NewState.Privileges(0).Luid = luid
        NewState.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED

        ' Adjust the token privileges
        If Not AdjustTokenPrivileges(hToken, False, NewState, Marshal.SizeOf(NewState), Nothing, Nothing) Then

            ' Write debug
            WriteEvent("Error: Windows API AdjustTokenPrivileges function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Lookup SeAssignPrimaryTokenPrivilege
        If Not LookupPrivilegeValue(Nothing, SE_ASSIGNPRIMARYTOKEN_NAME, luid) Then

            ' Write debug
            WriteEvent("Error: Windows API LookupPrivilegeValue function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Enable SeAssignPrimaryTokenPrivilege
        NewState.Privileges(0).Luid = luid
        NewState.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED

        ' Adjust the token privileges
        If Not AdjustTokenPrivileges(hToken, False, NewState, Marshal.SizeOf(NewState), Nothing, Nothing) Then

            ' Write debug
            WriteEvent("Error: Windows API AdjustTokenPrivileges function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Lookup SeTcbPrivilege
        If Not LookupPrivilegeValue(Nothing, SE_TCB_NAME, luid) Then

            ' Write debug
            WriteEvent("Error: Windows API LookupPrivilegeValue function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Enable SeTcbPrivilege
        NewState.Privileges(0).Luid = luid
        NewState.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED

        ' Adjust the token privileges
        If Not AdjustTokenPrivileges(hToken, False, NewState, Marshal.SizeOf(NewState), Nothing, Nothing) Then

            ' Write debug
            WriteEvent("Error: Windows API AdjustTokenPrivileges function returns an error." + Environment.NewLine +
                       "Windows API error code: " + Marshal.GetLastWin32Error.ToString, EventLogEntryType.Error)

            ' Return
            Return False

        End If

        ' Return
        Return True

    End Function

    Public Shared Sub LaunchProcess(ByVal LaunchType As String, ByVal CmdLine As String, ByVal args As String())

        ' ************
        ' LaunchType: System  --> Launch process as service account
        ' LaunchType: All     --> Launch process for all logged on users
        ' LaunchType: <User>  --> Launch process for a specific user
        ' LaunchType: Console --> Launch process for the console user
        ' ************

        ' Local variables
        Dim Arguments As String = ""
        Dim ExplorerProcesses As Process()
        Dim hToken As SafeTokenHandle = Nothing
        Dim principle As WindowsIdentity
        Dim principles As New ArrayList
        Dim phNewToken As SafeTokenHandle = Nothing
        Dim si As STARTUPINFO
        Dim pi As PROCESS_INFORMATION
        Dim FoundUserMatch As Boolean = False
        Dim DuplicatePrinciple As Boolean = False

        ' Process arguments
        For Each arg As String In args

            ' Build argument string
            Arguments += " " + arg

        Next

        ' Suppress Windows errors
        SetErrorMode(ErrorModes.SEM_FAILCRITICALERRORS Or ErrorModes.SEM_NOALIGNMENTFAULTEXCEPT)

        ' Increase Privileges
        If IncreasePrivileges() = False Then

            ' Write debug
            WriteEvent("Warning: Failed to increase current process privileges.", EventLogEntryType.Warning)

        End If

        ' Write debug
        WriteEvent("Launch Service Request: " + Environment.NewLine +
                   "Launch Type:     " + LaunchType + Environment.NewLine +
                   "Application:     " + CmdLine + Environment.NewLine +
                   "Arguments:       " + Arguments,
                   EventLogEntryType.Information)

        ' Check launch type
        If LaunchType.ToLower.Equals("system") Then

            ' ******************************
            ' * Launch process as service account.
            ' ******************************

            ' More local variables
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(CmdLine, Arguments)
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Write debug
            WriteEvent("Created new system process: " + Environment.NewLine +
                       "User:     " + System.Security.Principal.WindowsIdentity.GetCurrent.Name + Environment.NewLine +
                       "Process:  " + CmdLine + Arguments + Environment.NewLine +
                       "PID:      " + RunningProcess.Id.ToString, EventLogEntryType.Information)

        ElseIf LaunchType.ToLower.Equals("all") Then

            ' ******************************
            ' * Launch process for all logged on users.
            ' ******************************

            ' Get all explorer.exe IDs
            ExplorerProcesses = Process.GetProcessesByName("explorer")

            ' Verify explorers were found
            If ExplorerProcesses.Length = 0 Then

                ' Write debug
                WriteEvent("Error: No explorer.exe processes found." + Environment.NewLine + Environment.NewLine +
                    "Unable to complete launch request: " + Environment.NewLine +
                    "Launch Type:     " + LaunchType + Environment.NewLine +
                    "Application:     " + CmdLine + Environment.NewLine +
                    "Arguments:       " + Arguments,
                   EventLogEntryType.Error)

                ' Return
                Exit Sub

            End If

            ' Iterate each explorer.exe process
            For Each hProcess As Process In ExplorerProcesses

                ' Get the user token handle
                If OpenProcessToken(hProcess.Handle, TokenAccessLevels.MaximumAllowed, hToken) = False Then

                    ' Write debug
                    WriteEvent("Error: Windows API OpenProcessToken function returns an error." + Environment.NewLine +
                               "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                               "Unable to complete launch request: " + Environment.NewLine +
                               "Launch Type:     " + LaunchType + Environment.NewLine +
                               "Application:     " + CmdLine + Environment.NewLine +
                               "Arguments:       " + Arguments, EventLogEntryType.Error)

                    ' Iterate the next process
                    Continue For

                End If

                ' Get the windows identity
                principle = New WindowsIdentity(hToken.DangerousGetHandle)

                ' Iterate the principles list
                For Each login As WindowsIdentity In principles

                    ' Check for a match
                    If principle.Name.Equals(login.Name) Then

                        ' Set flag
                        DuplicatePrinciple = True

                        ' Match found -- break the loop
                        Exit For

                    End If

                Next

                ' Check if user is duplicate
                If DuplicatePrinciple Then

                    ' Reset flag
                    DuplicatePrinciple = False

                    ' Iterate the next process -- already launched for this user
                    Continue For

                Else

                    ' Add to principles list
                    principles.Add(principle)

                End If

                ' Get a primary token
                If Not DuplicateTokenEx(hToken.DangerousGetHandle,
                    TokenAccessLevels.MaximumAllowed,
                    Nothing,
                    SECURITY_IMPERSONATION_LEVEL.SecurityImpersonation,
                    TOKEN_TYPE.TokenPrimary,
                    phNewToken) Then

                    ' Write debug
                    WriteEvent("Error: Windows API DuplicateTokenEx function returns an error." + Environment.NewLine +
                               "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                               "Unable to complete launch request: " + Environment.NewLine +
                               "Launch Type:     " + LaunchType + Environment.NewLine +
                               "Application:     " + CmdLine + Environment.NewLine +
                               "Arguments:       " + Arguments, EventLogEntryType.Error)

                    ' Iterate the next process
                    Continue For

                End If

                ' Initialize process and startup info
                pi = New PROCESS_INFORMATION
                si = New STARTUPINFO
                si.cb = Marshal.SizeOf(si)
                si.lpDesktop = Nothing

                ' Launch the process in the client's logon session
                If Not CreateProcessAsUser(phNewToken,
                    Nothing,
                    CmdLine + Arguments,
                    Nothing,
                    Nothing,
                    False,
                    Nothing,
                    Nothing,
                    Nothing,
                    si,
                    pi) Then

                    ' Write debug
                    WriteEvent("Error: Windows API CreateProcessAsUser function returns an error." + Environment.NewLine +
                               "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                               "Unable to complete launch request: " + Environment.NewLine +
                               "Launch Type:     " + LaunchType + Environment.NewLine +
                               "Application:     " + CmdLine + Environment.NewLine +
                               "Arguments:       " + Arguments, EventLogEntryType.Error)

                Else

                    ' Write debug
                    WriteEvent("Created new user process: " + Environment.NewLine +
                               "User:     " + principle.Name + Environment.NewLine +
                               "Process:  " + CmdLine + Arguments + Environment.NewLine +
                               "PID:      " + pi.dwProcessId.ToString, EventLogEntryType.Information)

                End If

                ' Free resources
                hToken.Close()
                hToken = Nothing
                phNewToken.Close()
                phNewToken = Nothing
                principle = Nothing
                pi = Nothing
                si = Nothing

            Next

        ElseIf LaunchType.ToLower.Equals("console") Then

            ' ******************************
            ' * Launch process for the console user.
            ' ******************************

            ' Get the console user token handle
            If WTSQueryUserToken(WTSGetActiveConsoleSessionId, hToken) = False Then

                ' Write debug
                WriteEvent("Error: Windows API WTSQueryUserToken function returns an error." + Environment.NewLine +
                           "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                           "Unable to complete launch request: " + Environment.NewLine +
                           "Launch Type:     " + LaunchType + Environment.NewLine +
                           "Application:     " + CmdLine + Environment.NewLine +
                           "Arguments:       " + Arguments, EventLogEntryType.Error)

            End If

            ' Get the windows identity
            principle = New WindowsIdentity(hToken.DangerousGetHandle)

            ' Get a primary token
            If Not DuplicateTokenEx(hToken.DangerousGetHandle,
                TokenAccessLevels.MaximumAllowed,
                Nothing,
                SECURITY_IMPERSONATION_LEVEL.SecurityImpersonation,
                TOKEN_TYPE.TokenPrimary,
                phNewToken) Then

                ' Write debug
                WriteEvent("Error: Windows API DuplicateTokenEx function returns an error." + Environment.NewLine +
                           "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                           "Unable to complete launch request: " + Environment.NewLine +
                           "Launch Type:     " + LaunchType + Environment.NewLine +
                           "Application:     " + CmdLine + Environment.NewLine +
                           "Arguments:       " + Arguments, EventLogEntryType.Error)

                ' Return
                Exit Sub

            End If

            ' Initialize process and startup info
            pi = New PROCESS_INFORMATION
            si = New STARTUPINFO
            si.cb = Marshal.SizeOf(si)
            si.lpDesktop = Nothing

            ' Launch the process in the client's logon session
            If Not CreateProcessAsUser(phNewToken,
                Nothing,
                CmdLine + Arguments,
                Nothing,
                Nothing,
                False,
                CreateProcessFlags.CREATE_UNICODE_ENVIRONMENT,
                Nothing,
                Nothing,
                si,
                pi) Then

                ' Write debug
                WriteEvent("Error: Windows API CreateProcessAsUser function returns an error." + Environment.NewLine +
                           "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                           "Unable to complete launch request: " + Environment.NewLine +
                           "Launch Type:     " + LaunchType + Environment.NewLine +
                           "Application:     " + CmdLine + Environment.NewLine +
                           "Arguments:       " + Arguments, EventLogEntryType.Error)

            Else

                ' Write debug
                WriteEvent("Created new user process: " + Environment.NewLine +
                           "User:     " + principle.Name + Environment.NewLine +
                           "Process:  " + CmdLine + Arguments + Environment.NewLine +
                           "PID:      " + pi.dwProcessId.ToString, EventLogEntryType.Information)

            End If

            ' Free resources
            hToken.Close()
            hToken = Nothing
            phNewToken.Close()
            phNewToken = Nothing
            principle = Nothing
            pi = Nothing
            si = Nothing

        Else

            ' ******************************
            ' * Launch process for a specific user.
            ' ******************************

            ' Get all explorer.exe IDs
            ExplorerProcesses = Process.GetProcessesByName("explorer")

            ' Verify explorers were found
            If ExplorerProcesses.Length = 0 Then

                ' Write debug
                WriteEvent("Error: No explorer.exe processes found." + Environment.NewLine + Environment.NewLine +
                    "Unable to complete launch request: " + Environment.NewLine +
                    "Launch Type:     " + LaunchType + Environment.NewLine +
                    "Application:     " + CmdLine + Environment.NewLine +
                    "Arguments:       " + Arguments,
                   EventLogEntryType.Error)

                ' Return
                Exit Sub

            End If

            ' Iterate each explorer.exe process
            For Each hProcess As Process In ExplorerProcesses

                ' Get the user token handle
                If OpenProcessToken(hProcess.Handle, TokenAccessLevels.MaximumAllowed, hToken) = False Then

                    ' Write debug
                    WriteEvent("Error: Windows API OpenProcessToken function returns an error." + Environment.NewLine +
                               "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                               "Unable to complete launch request: " + Environment.NewLine +
                               "Launch Type:     " + LaunchType + Environment.NewLine +
                               "Application:     " + CmdLine + Environment.NewLine +
                               "Arguments:       " + Arguments, EventLogEntryType.Error)

                    ' Return
                    Continue For

                End If

                ' Get the windows identity
                principle = New WindowsIdentity(hToken.DangerousGetHandle)

                ' Check if this identity matches specified launch user
                If principle.Name.ToLower.Equals(LaunchType.ToLower) Then

                    ' Set the flag
                    FoundUserMatch = True

                    ' Get a primary token
                    If Not DuplicateTokenEx(hToken.DangerousGetHandle,
                        TokenAccessLevels.MaximumAllowed,
                        Nothing,
                        SECURITY_IMPERSONATION_LEVEL.SecurityImpersonation,
                        TOKEN_TYPE.TokenPrimary,
                        phNewToken) Then

                        ' Write debug
                        WriteEvent("Error: Windows API DuplicateTokenEx function returns an error." + Environment.NewLine +
                                   "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                                   "Unable to complete launch request: " + Environment.NewLine +
                                   "Launch Type:     " + LaunchType + Environment.NewLine +
                                   "Application:     " + CmdLine + Environment.NewLine +
                                   "Arguments:       " + Arguments, EventLogEntryType.Error)

                        ' Return
                        Exit Sub

                    End If

                    ' Initialize process and startup info
                    pi = New PROCESS_INFORMATION
                    si = New STARTUPINFO
                    si.cb = Marshal.SizeOf(si)
                    si.lpDesktop = Nothing

                    ' Launch the process in the client's logon session
                    If Not CreateProcessAsUser(phNewToken,
                        Nothing,
                        CmdLine + Arguments,
                        Nothing,
                        Nothing,
                        False,
                        CreateProcessFlags.CREATE_UNICODE_ENVIRONMENT,
                        Nothing,
                        Nothing,
                        si,
                        pi) Then

                        ' Write debug
                        WriteEvent("Error: Windows API CreateProcessAsUser function returns an error." + Environment.NewLine +
                                   "Windows API error code: " + Marshal.GetLastWin32Error.ToString + Environment.NewLine + Environment.NewLine +
                                   "Unable to complete launch request: " + Environment.NewLine +
                                   "Launch Type:     " + LaunchType + Environment.NewLine +
                                   "Application:     " + CmdLine + Environment.NewLine +
                                   "Arguments:       " + Arguments, EventLogEntryType.Error)

                    Else

                        ' Write debug
                        WriteEvent("Created new user process: " + Environment.NewLine +
                                   "User:     " + principle.Name + Environment.NewLine +
                                   "Process:  " + CmdLine + Arguments + Environment.NewLine +
                                   "PID:      " + pi.dwProcessId.ToString, EventLogEntryType.Information)

                    End If

                    ' Free resources
                    hToken.Close()
                    hToken = Nothing
                    phNewToken.Close()
                    phNewToken = Nothing
                    principle = Nothing
                    pi = Nothing
                    si = Nothing

                    ' Return
                    Exit For

                End If

            Next

            ' Check if a user was matched
            If FoundUserMatch = False Then

                ' Write debug
                WriteEvent("Error: Unable to create specified user process." + Environment.NewLine +
                           "Specified launch user is not logged on: " + LaunchType + Environment.NewLine + Environment.NewLine +
                           "Unable to complete launch request: " + Environment.NewLine +
                           "Launch Type:     " + LaunchType + Environment.NewLine +
                           "Application:     " + CmdLine + Environment.NewLine +
                           "Arguments:       " + Arguments, EventLogEntryType.Error)

            End If

        End If

    End Sub

    Public Shared Sub WriteEvent(EventMessage As String, EntryType As EventLogEntryType)

        ' Check if event source exists
        If Not EventLog.SourceExists("Launch Service") Then

            ' Create the event source
            EventLog.CreateEventSource("Launch Service", "System")

        End If

        ' Write the message
        EventLog.WriteEntry("Launch Service", EventMessage, EntryType)

    End Sub

    Public Shared Function IsProcessElevated() As Boolean

        ' Local variables
        Dim IsElevated As Boolean = False
        Dim hToken As SafeTokenHandle = Nothing
        Dim cbTokenElevation As Integer = 0
        Dim pTokenElevation As IntPtr = IntPtr.Zero

        ' Encapsulate API calls
        Try

            ' Open the access token of the current process with TOKEN_QUERY
            If Not OpenProcessToken(Process.GetCurrentProcess.Handle, TOKEN_QUERY, hToken) Then

                ' Unable to open current process token
                Throw New Win32Exception

            End If

            ' Allocate a buffer for the elevation information
            cbTokenElevation = Marshal.SizeOf(GetType(TOKEN_ELEVATION))
            pTokenElevation = Marshal.AllocHGlobal(cbTokenElevation)

            ' Check buffer
            If (pTokenElevation = IntPtr.Zero) Then

                ' Unable to allocate buffer
                Throw New Win32Exception

            End If

            ' Retrieve token elevation information
            If Not GetTokenInformation(hToken, TOKEN_INFORMATION_CLASS.TokenElevation, pTokenElevation, cbTokenElevation, cbTokenElevation) Then

                ' Unable to get token information
                Throw New Win32Exception

            End If

            ' Marshal the TOKEN_ELEVATION struct from native to .NET
            Dim elevation As TOKEN_ELEVATION = Marshal.PtrToStructure(pTokenElevation, GetType(TOKEN_ELEVATION))

            ' TOKEN_ELEVATION.TokenIsElevated is a non-zero value if the token has elevated privileges; otherwise, a zero value
            IsElevated = (elevation.TokenIsElevated <> 0)

        Finally

            ' Dispose of access token
            If (Not hToken Is Nothing) Then
                hToken.Close()
                hToken = Nothing
            End If

            ' Free pointer memory
            If (pTokenElevation <> IntPtr.Zero) Then
                Marshal.FreeHGlobal(pTokenElevation)
                pTokenElevation = IntPtr.Zero
                cbTokenElevation = 0
            End If

        End Try

        ' Return
        Return IsElevated

    End Function

    Public Shared Function GetProcessIntegrityLevel() As Integer

        ' Decalre and initialize variables
        Dim IntegrityLevel As Integer = -1
        Dim hToken As SafeTokenHandle = Nothing
        Dim cbTokenIL As Integer = 0
        Dim pTokenIL As IntPtr = IntPtr.Zero

        Try
            ' Open the access token of the current process with TOKEN_QUERY
            If (Not OpenProcessToken(Process.GetCurrentProcess.Handle, TOKEN_QUERY, hToken)) Then

                ' Unable to open process token
                Throw New Win32Exception

            End If

            ' Get buffer size (cbTokenIL)
            If (Not GetTokenInformation(hToken, TOKEN_INFORMATION_CLASS.TokenIntegrityLevel, IntPtr.Zero, 0, cbTokenIL)) Then

                ' Capture last error
                Dim err As Integer = Marshal.GetLastWin32Error

                ' Check error cause was for buffer
                If (err <> ERROR_INSUFFICIENT_BUFFER) Then

                    ' A different error occurred
                    Throw New Win32Exception(err)

                End If

            End If

            ' Allocate the buffer size.
            pTokenIL = Marshal.AllocHGlobal(cbTokenIL)

            ' Check buffer
            If (pTokenIL = IntPtr.Zero) Then

                ' Unable to allocate buffer
                Throw New Win32Exception

            End If

            ' Get token integrity level with proper buffer size
            If (Not GetTokenInformation(hToken, TOKEN_INFORMATION_CLASS.TokenIntegrityLevel, pTokenIL, cbTokenIL, cbTokenIL)) Then

                ' Unable to get token information
                Throw New Win32Exception

            End If

            ' Marshal the TOKEN_MANDATORY_LABEL struct from native to .NET object
            Dim tokenIL As TOKEN_MANDATORY_LABEL = Marshal.PtrToStructure(pTokenIL, GetType(TOKEN_MANDATORY_LABEL))

            ' Integrity Level SIDs are in the form of S-1-16-0xXXXX. (e.g. 
            ' S-1-16-0x1000 stands for low integrity level SID). There is one and 
            ' only one subauthority.
            Dim pIL As IntPtr = GetSidSubAuthority(tokenIL.Label.Sid, 0)
            IntegrityLevel = Marshal.ReadInt32(pIL)

        Finally

            ' Centralized cleanup for all allocated resources.
            If (Not hToken Is Nothing) Then
                hToken.Close()
                hToken = Nothing
            End If

            If (pTokenIL <> IntPtr.Zero) Then
                Marshal.FreeHGlobal(pTokenIL)
                pTokenIL = IntPtr.Zero
                cbTokenIL = 0
            End If

        End Try

        ' Return
        Return IntegrityLevel

    End Function

    Private Shared Function OpenReparsePoint(ByVal reparsePoint As String, ByVal accessMode As EFileAccess) As SafeFileHandle

        ' Local variables
        Dim reparsePointHandle As SafeFileHandle = Nothing

        ' Open handle to reparse point
        reparsePointHandle = New SafeFileHandle(CreateFile(reparsePoint,
                                                           accessMode,
                                                           EFileShare.Read Or EFileShare.Write Or EFileShare.Delete,
                                                           IntPtr.Zero,
                                                           ECreationDisposition.OpenExisting,
                                                           EFileAttributes.BackupSemantics Or EFileAttributes.OpenReparsePoint,
                                                           IntPtr.Zero), True)

        ' Check last error
        If Marshal.GetLastWin32Error <> 0 Then

            ' Throw an exception
            Throw New Win32Exception("Unable to open reparse point.")

        End If

        ' Return the handle
        Return reparsePointHandle

    End Function

    Public Shared Sub RemoveJunction(ByVal JunctionPoint As String)

        ' Local variables
        Dim FileHandle As SafeFileHandle
        Dim ReparseDataBuffer As REPARSE_DATA_BUFFER = New REPARSE_DATA_BUFFER()
        Dim InBufferSize As Integer
        Dim InBuffer As IntPtr
        Dim BytesReturned As Integer
        Dim Result As Boolean

        ' Verify the directory/junction exists
        If Not System.IO.Directory.Exists(JunctionPoint) Then

            ' Return
            Return

        End If

        ' Open the junction point
        FileHandle = OpenReparsePoint(JunctionPoint, EFileAccess.GenericWrite)

        ' Setup reparse structure
        ReparseDataBuffer.ReparseTag = IO_REPARSE_TAG_MOUNT_POINT
        ReparseDataBuffer.ReparseDataLength = 0
        ReDim ReparseDataBuffer.PathBuffer(&H3FF0)

        ' Calculate buffer size and allocate
        InBufferSize = Marshal.SizeOf(ReparseDataBuffer)
        InBuffer = Marshal.AllocHGlobal(InBufferSize)

        ' Try to delete reparse point
        Try

            ' Create the pointer
            Marshal.StructureToPtr(ReparseDataBuffer, InBuffer, False)

            ' Delete the reparse point
            Result = DeviceIoControl(FileHandle.DangerousGetHandle(),
                                     FSCTL_DELETE_REPARSE_POINT,
                                     InBuffer,
                                     8,
                                     IntPtr.Zero,
                                     0,
                                     BytesReturned,
                                     IntPtr.Zero)

            ' Check result
            If Not Result Then

                ' Throw an exception
                Throw New Win32Exception("Unable to delete reparse point.")

            End If

        Finally

            ' Release the handle!!!!!!!!!!
            FileHandle.Close()

            ' No matter what, free the allocated buffer
            Marshal.FreeHGlobal(InBuffer)

        End Try

    End Sub

    Public Shared Sub RefreshNotificationArea()

        ' Local variables
        Dim notificationAreaHandle = GetNotificationAreaHandle()

        ' Exit on invalid handle
        If notificationAreaHandle = IntPtr.Zero Then
            Return
        End If

        ' Refresh window (via mouse-over messages)
        RefreshWindow(notificationAreaHandle)

    End Sub

    Private Shared Sub RefreshWindow(windowHandle As IntPtr)

        ' Local constants and variables
        Const wmMousemove As UInteger = &H200
        Dim rect As RECT

        ' Retrieve coordinates of window area
        GetClientRect(windowHandle, rect)

        ' Send mouse move message to window area
        For x As Integer = 0 To rect.Right - 1 Step 5
            For y As Integer = 0 To rect.Bottom - 1 Step 5
                SendMessage(windowHandle, wmMousemove, 0, (y << 16) + x)
            Next
        Next

    End Sub

    Private Shared Function GetNotificationAreaHandle() As IntPtr

        ' Local constants
        Const notificationAreaTitle As String = "Notification Area"
        Const notificationAreaTitleInWindows7 As String = "User Promoted Notification Area"

        ' Local variables
        Dim systemTrayContainerHandle = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Shell_TrayWnd", String.Empty)
        Dim systemTrayHandle = FindWindowEx(systemTrayContainerHandle, IntPtr.Zero, "TrayNotifyWnd", String.Empty)
        Dim sysPagerHandle = FindWindowEx(systemTrayHandle, IntPtr.Zero, "SysPager", String.Empty)
        Dim notificationAreaHandle = FindWindowEx(sysPagerHandle, IntPtr.Zero, "ToolbarWindow32", notificationAreaTitleInWindows7)

        ' Check for empty handle
        ' Previous to Windows 7, the "User Promoted Notification Area" was named the "Notification Area"
        If notificationAreaHandle = IntPtr.Zero Then
            notificationAreaHandle = FindWindowEx(sysPagerHandle, IntPtr.Zero, "ToolbarWindow32", notificationAreaTitle)
        End If

        ' Return
        Return notificationAreaHandle

    End Function

    Public Shared Sub DetachConsole()

        ' Local variables
        Dim cw As IntPtr = GetConsoleWindow()

        ' Free console
        FreeConsole()

        ' Send message to console
        SendMessage(cw, &H102, 13, IntPtr.Zero)

    End Sub

End Class