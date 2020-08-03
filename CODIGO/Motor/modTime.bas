Attribute VB_Name = "modTime"
Option Explicit

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Ticks_Per_Second As Currency
Public Start_Time As Currency

Public Function Hi_Res_Timer_Initialize() As Boolean

    If QueryPerformanceFrequency(Ticks_Per_Second) = 0 Then
        Hi_Res_Timer_Initialize = False
    Else
        QueryPerformanceCounter Start_Time
        Hi_Res_Timer_Initialize = True
    End If

End Function

Public Function Get_Elapsed_Time() As Single
    
    Dim Last_Time As Currency
    Dim Current_Time As Currency
    
    QueryPerformanceCounter Current_Time
    Get_Elapsed_Time = (Current_Time - Last_Time) / Ticks_Per_Second
    QueryPerformanceCounter Last_Time
    
End Function

Public Function Format_Time(ByVal Time As Single) As String

    Dim Milliseconds As Long
    Dim Seconds As Long
    Dim Minutes As Long
    
    Milliseconds = Time * 1000
    Seconds = Time
    Minutes = Seconds \ 60
    Seconds = Seconds - Minutes * 60
    
    Format_Time = Right$("00" & Minutes, 2) & "m " & Right$("00" & Seconds, 2) & "s " & Right$("000" & Milliseconds, 3) & "ms"

End Function

Public Sub Lock_Framerate(Target_FPS As Long)
 
    Static Last_Time As Currency
    Dim Current_Time As Currency
    Dim FPS As Single
   
    Do
        QueryPerformanceCounter Current_Time
        FPS = Ticks_Per_Second / (Current_Time - Last_Time)
    Loop While (FPS > Target_FPS)
   
    QueryPerformanceCounter Last_Time
 
End Sub

