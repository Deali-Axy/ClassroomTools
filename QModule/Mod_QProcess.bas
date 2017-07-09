Attribute VB_Name = "QProcess"
Option Explicit
'【代码协会 VB通用模块库】CodeInstitute VB Common Modules Library
'【模块名】 进程管理模块
'【作者】CI Deali-Axy

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function NtSuspendProcess Lib "NTDLL.DLL" (ByVal hProc As Long) As Long
Private Declare Function NtResumeProcess Lib "NTDLL.DLL" (ByVal hProc As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function DbgUiDebugActiveProcess Lib "ntdll" (ByVal ProcessHandle As Long) As Long
Private Declare Function DbgUiStopDebugging Lib "ntdll" (ByVal ProcessHandle As Long) As Long
Private Type PROCESSENTRY32
    dwsize As Long
    cntusage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Private hProcess As Long

Public Function SuspendProcess(ProcessName As String) As Long
    Dim pid As Long
    pid = GetPID(ProcessName)
    SuspendProcess = pid
    If IsNumeric(pid) Then
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CLng(pid))
        If hProcess <> 0 Then
            NtSuspendProcess hProcess
        End If
    End If
    CloseHandle hProcess
End Function

Public Function ResumeProcess(ProcessName As String)
    Dim pid As Long
    pid = GetPID(ProcessName)
    If IsNumeric(pid) Then
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CLng(pid))
        If hProcess <> 0 Then
            NtResumeProcess hProcess
        End If
    End If
    CloseHandle hProcess
End Function

Public Function TerminateProcessEx(ProcessName As String)
    Dim pid As Long
    pid = GetPID(ProcessName)
    If IsNumeric(pid) Then
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CLng(pid))
        If hProcess <> 0 Then
            TerminateProcess hProcess, 0
        End If
    End If
End Function

Public Function GetPID(ProcessName As String) As Long
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(&H2, 0&)
    If lSnapShot <> -1 Then
        tPE.dwsize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If LCase$(ProcessName) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                Dim lProcess As Long
                Dim lExitCode As Long
                GetPID = tPE.th32ProcessID
                CloseHandle lProcess
            End If
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        CloseHandle (lSnapShot)
    End If
End Function
