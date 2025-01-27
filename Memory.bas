Attribute VB_Name = "Memory"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Global Const KEYEVENTF_KEYUP = &H2
    Global Const KEYEVENTF_EXTENDEDKEY = &H1
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As KeyCodeConstants) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As ProcessEntry32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As ProcessEntry32) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Const WM_LBUTTONDOWN = &H201
Public Const MOUSEEVENTF_LEFTDOWN = &H2      ' left button down
Public Const MOUSEEVENTF_LEFTUP = &H4        ' left button up

Global pt As POINTAPI


      Public Type ProcessEntry32
         dwSize As Long
         cntUsage As Long
         th32ProcessID As Long
         th32DefaultHeapID As Long
         th32ModuleID As Long
         cntThreads As Long
         th32ParentProcessID As Long
         pcPriClassBase As Long
         dwFlags As Long
         szExeFile As String * 260
      End Type

      Public Type OSVERSIONINFO
         dwOSVersionInfoSize As Long
         dwMajorVersion As Long
         dwMinorVersion As Long
         dwBuildNumber As Long
         dwPlatformId As Long
         szCSDVersion As String * 128
      End Type
      
      
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0
Public Const LB_SETTABSTOPS = &H192
Public Const WM_KEYDOWN = &H100
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION
Const MEM_COMMIT = &H1000
Const MEM_RESERVE = &H2000
Const MEM_DECOMMIT = &H4000
Const MEM_RELEASE = &H8000
Const MEM_FREE = &H10000
Const MEM_PRIVATE = &H20000
Const MEM_MAPPED = &H40000
Const MEM_TOP_DOWN = &H100000

'==========Memory access constants===========
Public Const PAGE_NOACCESS = &H1&
Public Const PAGE_READONLY = &H2&
Public Const PAGE_READWRITE = &H4&
Public Const PAGE_WRITECOPY = &H8&
Public Const PAGE_EXECUTE = &H10&
Public Const PAGE_EXECUTE_READ = &H20&
Public Const PAGE_EXECUTE_READWRITE = &H40&
Public Const PAGE_EXECUTE_WRITECOPY = &H80&
Public Const PAGE_GUARD = &H100&
Public Const PAGE_NOCACHE = &H200&


    Global pid As Long, hProcess As Long, hWin As Long
    Global sBuffer As String
    Global sngbuffer As Single
    Global byteBuffer As Byte
    Global bytesBuffer() As Byte
    Global byteInj() As Byte
    Global longBuffer As Long
    Global intBuffer As Integer

    Dim Test As Long
    Dim offset As Long
    
    
Public Function GetString(address As Long, Length As Long)
On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
sBuffer = String(Length, 0)
Test = ReadProcessMemory(hProcess, ByVal address, ByVal sBuffer, ByVal Length, 0)
GetString = sBuffer
CloseHandle hProcess
End Function

Public Function GetID(name)
'hWin = FindWindow(vbNullString, name) 'InstanceToWnd(pid) 'get handle of launched window - only to repaint it after changes
hWin = FindWindow(name, vbNullString)
Call GetWindowThreadProcessId(hWin, pid)
End Function
Public Function SetString(address As Long, value, Optional Length As Long)
On Error Resume Next

Length = Len(value)
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = VirtualProtectEx(hProcess, ByVal address, ByVal Length, PAGE_READWRITE, 0)
sBuffer = value
Test = WriteProcessMemory(hProcess, ByVal address, ByVal sBuffer, ByVal Length, 0)
CloseHandle hProcess
SetString = Test
End Function
Public Function SetHex(address As Long, value As String, Optional Length As Long)
On Error Resume Next
value = Replace(value, " ", "")
value = Replace(value, "-", "")
Length = Len(value) / 2
HexToBytes (value)

hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = VirtualProtectEx(hProcess, ByVal address, ByVal Length, PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address, bytesBuffer(1), ByVal Length, 0)
CloseHandle hProcess
SetHex = Test
End Function
Public Function GetFloat(address As Long) As Single
On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = ReadProcessMemory(hProcess, ByVal address, sngbuffer, ByVal 4, 0)
CloseHandle hProcess
GetFloat = sngbuffer
End Function
Public Function SetFloat(address As Long, value As Single)
On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = VirtualProtectEx(hProcess, ByVal address, ByVal 4, PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address, value, ByVal 4, 0)
CloseHandle hProcess
SetFloat = Test
End Function

Public Function GetByte(address As Long) As Byte

On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = ReadProcessMemory(hProcess, ByVal address, byteBuffer, ByVal 1, 0)
CloseHandle hProcess
GetByte = byteBuffer
End Function
Public Function GetLong(address As Long) As Long

On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = ReadProcessMemory(hProcess, ByVal address, longBuffer, ByVal 4, 0)
CloseHandle hProcess
GetLong = longBuffer
End Function
Public Function GetInteger(address As Long) As Integer

On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = ReadProcessMemory(hProcess, ByVal address, intBuffer, ByVal 2, 0)
CloseHandle hProcess
GetInteger = intBuffer
End Function
Public Function SetLong(address As Long, value As Long)

On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = VirtualProtectEx(hProcess, ByVal address, ByVal 4, PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address, value, ByVal 4, 0)
CloseHandle hProcess

End Function
Public Function SetInteger(address As Long, value As Integer)

On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = VirtualProtectEx(hProcess, ByVal address, ByVal 2, PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address, value, ByVal 2, 0)
CloseHandle hProcess

End Function
Public Function SetByte(address As Long, value As Byte)

On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
Test = VirtualProtectEx(hProcess, ByVal address, ByVal 1, PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address, value, ByVal 1, 0)
CloseHandle hProcess

End Function
Public Function SetBytes(address As Long, Optional Start As Long)
On Error Resume Next
hProcess = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
If Start > 0 Then
Test = VirtualProtectEx(hProcess, ByVal address + Start, (UBound(byteInj) + 1) - Start, PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address + Start, byteInj(Start), ByVal (UBound(byteInj) + 1) - Start, 0)
Else
Test = VirtualProtectEx(hProcess, ByVal address, (UBound(byteInj) + 1), PAGE_READWRITE, 0)
Test = WriteProcessMemory(hProcess, ByVal address, byteInj(0), ByVal (UBound(byteInj) + 1), 0)

End If
CloseHandle hProcess
SetBytes = Test
End Function





Public Function Pause(Seconds As Single)
On Error Resume Next
Dim Start As Single
    Start = GetTickCount
    Do While GetTickCount < Start + Seconds
         DoEvents
    Loop
End Function
Public Function HexToBytes(hexstring As String) As String
Dim yLoop As Long
Dim myResult As String
ReDim bytesBuffer(1 To (Len(hexstring) / 2))
For yLoop = 1 To Len(hexstring) Step 2
bytesBuffer(((yLoop - 1) / 2) + 1) = CByte("&H" & Mid$(hexstring, yLoop, 2))
Next yLoop
End Function


Public Function getVersion() As Long
         Dim osinfo As OSVERSIONINFO
         Dim retvalue As Integer
         osinfo.dwOSVersionInfoSize = 148
         osinfo.szCSDVersion = Space$(128)
         retvalue = GetVersionExA(osinfo)
         getVersion = osinfo.dwPlatformId
End Function
      Function StrZToStr(s As String) As String
         StrZToStr = Left$(s, Len(s) - 1)
      End Function

Public Function BringWindowTop(inHwnd As Long)
SetWindowPos inHwnd, 0, 0, 0, 0, 0, &H80 Or 2 Or 1
'DoEvents
SetWindowPos inHwnd, 0, 0, 0, 0, 0, &H40 Or 2 Or 1
End Function
