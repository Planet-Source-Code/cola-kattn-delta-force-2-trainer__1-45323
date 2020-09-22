VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'API and functions to read\write bytes to memory.
'This class module only include functions to read\write BYTE values (0 to 255)
'because there is no higher ammo value in df2.
'This code was written by me (Cola-Kattn) for my needs in the trainer.

'However, please feel free to copy and learn from this code!


'API declarations for the trainer itself and memory manipulation
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'API declarations for the MoveForm sub
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim MemWriteFail As Integer
Dim MemReadFail As Integer

Public Sub MoveForm(Form As Object, ButtonDown As Integer)
    'Move the form
   Dim lngReturnValue As Long
   If ButtonDown = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Form.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
  
End Sub

Public Function Test(WindowText As String) As Integer

    ' I use this function to test if the trainer is able to find the game window,
    ' if not, the fucntion will return a null and an error message should show.
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, WindowText)
    If hwnd = 0 Then Test = 0 Else Test = 1
    
End Function

Public Function WriteByte(WindowText As String, adress As Long, value As Byte)

    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    'Attempt to find the process window ("Delta Force 2,  V1.06.15")
    hwnd = FindWindow(vbNullString, WindowText)
    
    'If the window is not found, the function will return a null.
    'This is not really necessary, because I test to see if the window can be found before the timers
    'are enabled.
    If hwnd = 0 Then
        ResetAll
        Exit Function
    End If
    
    'Get the process ID
    GetWindowThreadProcessId hwnd, ProcessID
    
    'Attempt to get a process handle, and then check if one is found.
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
        MsgBox "Could not find the process ID", vbCritical, "Error"
        Exit Function
    End If
    
    'Write the new value to memory.
    WriteProcessMemory ProcessHandle, adress, value, 2, 0&
    
    'Close the process handle
    CloseHandle ProcessHandle
    
End Function

Public Function ReadByte(WindowText As String, adress As Long, value As Byte)

    'This code is the same as above except that on the end, we read a byte instead of writing.
    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    
    hwnd = FindWindow(vbNullString, WindowText)
    
    If hwnd = 0 Then
        ResetAll
        Exit Function
    End If
    
    GetWindowThreadProcessId hwnd, ProcessID
    
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    
    If ProcessHandle = 0 Then
        MsgBox "Could not find the process ID", vbCritical, "Error"
        Exit Function
    End If
    
    'Read from memory.
    ReadProcessMem ProcessHandle, adress, value, 2, 0&
    
    CloseHandle ProcessHandle
    
End Function

Private Sub ResetAll()
    
    'This will reset all timers and command buttons the default.
    'Only time this would be necessary is when DF2 is shutdown when some of the timers are enabled.
    
    With Form1
        .tmrPrimary.Enabled = False
        .cmdPrimary.Caption = "Enable"
        .tmrSidearm.Enabled = False
        .cmdSidearm.Caption = "Enable"
        .tmrSecondary.Enabled = False
        .cmdSecondary.Caption = "Enable"
        .tmrGrenades.Enabled = False
        .cmdGrenades.Caption = "Enable"
        .tmrHealth.Enabled = False
        .cmdHealth.Caption = "Enable"
    End With

End Sub
