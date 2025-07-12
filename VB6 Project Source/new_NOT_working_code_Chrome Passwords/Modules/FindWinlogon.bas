Attribute VB_Name = "FindWinlogon"
Option Explicit

Public Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Long = 260

Public Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * MAX_PATH
End Type


Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long



Public mbImpersonating As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const PROCESS_VM_READ As Long = (&H10)
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const MAXIMUM_ALLOWED As Long = &H2000000

Public Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Public Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, ByRef TokenHandle As Long) As Long
Public Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'Windows Directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'--------------------------------
' Find "winlogon.exe" Process ID |
'--------------------------------
Public Sub FindWinlogonEXE()

Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim Success As Long
Dim Buffer As String
Dim Ret As Long
Dim Ruta As String
Dim Handle_Proceso As Long


Form1.ListWinlogon.ListItems.Clear
 
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

If hSnapShot = -1 Then Exit Sub
uProcess.dwSize = Len(uProcess)
Success = ProcessFirst(hSnapShot, uProcess)

If Success = 1 Then
    Do
        Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + PROCESS_VM_READ, 0, uProcess.th32ProcessID)
        'Buffer = Space(255)
        'ret = GetModuleFileNameExA(Handle_Proceso, 0, Buffer, 255)
        'path = Left(Buffer, ret)
                       
        With Form1.ListWinlogon.ListItems.Add()
            .text = uProcess.szExeFile
            .SubItems(1) = uProcess.th32ProcessID
            If LCase(.text) = "winlogon.exe" Then
                .SubItems(2) = "<--- Found !"
                Form1.txtWinLogon.text = uProcess.th32ProcessID
            End If
        End With
    
    
    Loop While ProcessNext(hSnapShot, uProcess)
End If

Call CloseHandle(hSnapShot)

End Sub

'-----------------
' ImpersonateSelf |
'-----------------
Public Function ImpersonateSelf() As Boolean

Dim hToken As Long
Dim hNewSelf As Long

If mbImpersonating Then StopImpersonation

'Get "winlogon.exe" to impersonate (OpenProcess)
Dim getProcessHandle As Long
getProcessHandle = OpenProcess(PROCESS_ALL_ACCESS + PROCESS_QUERY_INFORMATION + PROCESS_VM_READ, 0, Form1.txtWinLogon.text)
If (getProcessHandle = 0) Then
   'MsgBox "OpenProcess failed with error code " & Err.LastDllError
   Else
   'MsgBox "OpenProcess success", , "OpenProcess"
End If


'Get "winlogon.exe" to impersonate (OpenProcessToken)
Dim TokOpen As Long
TokOpen = OpenProcessToken(getProcessHandle, MAXIMUM_ALLOWED, hToken)
If (TokOpen = 0) Then
   MsgBox "OpenProcessToken failed with error code " & Err.LastDllError
   Else
   MsgBox "OpenProcessToken SUCCESS!", , "OpenProcessToken"
End If

'Get "winlogon.exe" to impersonate (ImpersonateLoggedOnUser)
hNewSelf = ImpersonateLoggedOnUser(hToken)
If (hNewSelf = 0) Then
    MsgBox "ImpersonateLoggedOnUser failed with error code " & Err.LastDllError
    Else
    MsgBox "ImpersonateLoggedOnUser SUCCESS!", , "ImpersonateUser"
    mbImpersonating = True
    ImpersonateSelf = True
    hToken = CloseHandle(hToken)
End If

End Function

'Stop the Impersonation of the processs on EXIT (Unload)
Public Sub StopImpersonation()
    RevertToSelf
    mbImpersonating = False
End Sub

Public Function WinDrive() As String
    Dim Buffer As String * 512, Length As Integer
    Length = GetWindowsDirectory(Buffer, Len(Buffer))
    WinDrive = Left$(Buffer, Length)
    WinDrive = Mid(WinDir, 1, 3)
End Function

Public Function WinDir() As String
    Dim Buffer As String * 512, Length As Integer
    Length = GetWindowsDirectory(Buffer, Len(Buffer))
    WinDir = Left$(Buffer, Length)
End Function


