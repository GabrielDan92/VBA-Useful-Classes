'THIS IS NOT MY CODE. This is a version of the QueryPerformanceCounter Windows API wrote using VBA. 
'More details here: https://docs.microsoft.com/ro-ro/windows/win32/sysinfo/acquiring-high-resolution-time-stamps?redirectedfrom=MSDN
'And here: https://stackoverflow.com/questions/198409/how-do-you-test-running-time-of-vba-code

'The objective of this class is to create an accurate VBA timer and count the total runtime.
'The class can be called by creating a new instance of the class: Dim timer (or whatever class name you want) As new clsTimer.
'The counter will be initialized using className.StartCounter. className.TimeElapsed will be used for retrieving the results.


Option Explicit

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Double

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Function LI2Double(LI As LARGE_INTEGER) As Double
Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Private Sub Class_Initialize()
Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
End Sub

Public Sub StartCounter()
    QueryPerformanceCounter m_CounterStart
End Sub

Property Get TimeElapsed() As Double
Dim crStart As Double
Dim crStop As Double
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / m_crFrequency
End Property
