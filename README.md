<div align="center">

## RegSvr API


</div>

### Description

The purpose of this function is to register/unregister a DLL/OCX with NO INTERFACE. That's right, no message boxes (what a concept).

It is in function form with an enumerated return value along with a PrintXXX function to convert a return value to string (my personal touch).

This code was originally written by Herman Liu, but my 120 minutes of editing and consolodating is worth the ink.
 
### More Info
 
Filespec (string): The complete filename of the .OCX/.DLL

RVsU (Boolean): Register or unregister. Use yer head.

Assumptions are that you can rectify mistakes you make with this code :)

An enumerated value- 0 means success.

Use the PrintDLLRegService() function to get a string version.

ONLY TESTED ON WIN98


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Micah Epps](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/micah-epps.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/micah-epps-regsvr-api__1-13630/archive/master.zip)

### API Declarations

See below...


### Source Code

```
'''By Herman Liu, EDITED by Micah Epps: MTEXX@zebra.net
Option Explicit
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpdata As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
'Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long,  ByVal dwExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Enum DLLRegServiceResults
  regSuccess = 0
  regFailLoadLib
  regFailCreateThread
  regThreadTimeout
End Enum
Public Function PrintDLLRegServiceResults(ByVal Value As DLLRegServiceResults) As String
  Dim Temp As String '''typing the above sux
  Select Case Value
  Case regSuccess: Temp = "success"
  Case regFailLoadLib: Temp = "failed to load library"
  Case regFailCreateThread: Temp = "failed to create thread"
  Case regThreadTimeout: Temp = "thread timed out"
  Case Else: Temp = "UNKNOWN"
  End Select
  PrintDLLRegServiceResults = Temp
End Function
Public Function DLLRegisterService(ByVal Filespec As String, ByVal RegVsUnreg As Boolean) As DLLRegServiceResults
  '''DOS filenames (8.3 / no spaces) are NOT necesary! :)
  Dim hLib As Long         ' Store handle of the control library
  Dim lpDLLEntryPoint As Long   ' Store the address of function called
  Dim lpThreadID As Long      ' Pointer that receives the thread identifier
  Dim lpExitCode As Long      ' Exit code of GetExitCodeThread
  Dim mResult As Long
  Dim hThread
  Const RegProcName = "DllRegisterServer"
  Const UnregProcName = "DllUnregisterServer"
  '''Load the control DLL, i. e. map the specified DLL file into the address space of the calling process
  hLib = LoadLibrary(Filespec)
  If hLib = 0 Then
    DLLRegisterService = regFailLoadLib
    Exit Function
  End If
  '''Find and store the DLL entry point, i.e. obtain the address of the &#8220;DllRegisterServer&#8221; or "DllUnregisterServer" function (to register or deregister the server&#8217;s components in the registry)
  lpDLLEntryPoint = GetProcAddress(hLib, IIf(RegVsUnreg, RegProcName, UnregProcName))
  If lpDLLEntryPoint = vbNull Then
    FreeLibrary hLib
    DLLRegisterService = regFailLoadLib
    Exit Function
  End If
  '''Create a thread to execute within the virtual address space of the calling process
  hThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
  If hThread = 0 Then
    FreeLibrary hLib
    DLLRegisterService = regFailCreateThread
    Exit Function
  End If
  '''Use WaitForSingleObject to check the return state (i) when the specified object is in the signaled state or (ii) when the time-out interval elapses. This function can be used to test Process and Thread.
  mResult = WaitForSingleObject(hThread, 10000)
  If mResult <> 0 Then
    FreeLibrary hLib
    lpExitCode = GetExitCodeThread(hThread, lpExitCode)
    ExitThread lpExitCode
    DLLRegisterService = regThreadTimeout
    Exit Function
  End If
  '''We don't call the dangerous TerminateThread(); after the last handle to an object is closed, the object is removed from the system.
  CloseHandle hThread
  FreeLibrary hLib
  DLLRegisterService = regSuccess
End Function
```

