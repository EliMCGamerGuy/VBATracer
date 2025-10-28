Attribute VB_Name = "AdvancedProfiler"
' AdvancedProfiler.bas
'
' CODE ATTRIBUTION:
'
' This specific module is cut and paste from OpenAI's ChatGPT GPT-5.
' It is wizardry, and I hope it never breaks.
' *(crosses fingers)*
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private gFreq As Currency
Private gFreqInit As Boolean

Public Sub Prof_Init()
    Dim ok As Long
    ok = QueryPerformanceFrequency(gFreq)
    gFreqInit = (ok <> 0)
End Sub

Public Function PerfNow() As Double
    ' returns seconds as Double
    Dim val As Currency
    If Not gFreqInit Then Prof_Init
    QueryPerformanceCounter val
    If gFreqInit Then
        PerfNow = (val / gFreq) ' Currency/Currency -> Double-ish
    Else
        PerfNow = Timer
    End If
End Function

