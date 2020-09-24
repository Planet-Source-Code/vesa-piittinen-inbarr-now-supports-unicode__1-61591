Attribute VB_Name = "TimeIt"
Option Explicit

Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Dim cStartTime As Currency
Dim cPerfFreq As Currency

Public Function Start() As Long
  If QueryPerformanceFrequency(cPerfFreq) = False Then
    Debug.Print "High-perf counter not supported"
  End If
  QueryPerformanceCounter cStartTime
End Function

Public Function Finish() As Long
  Dim cCurrentTime As Currency
  
  QueryPerformanceCounter cCurrentTime
  Finish = 1000 * (cCurrentTime - cStartTime) / cPerfFreq
End Function
