Attribute VB_Name = "APIs"
Option Explicit
' API's Declarations

Public Declare Function mciGetDeviceIDFromElementID Lib "winmm.dll" Alias "mciGetDeviceIDFromElementIDA" (ByVal dwElementID As Long, ByVal lpstrType As String) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciGetYieldProc Lib "winmm" (ByVal mciId As Long, pdwYieldData As Long) As Long
Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As Any) As Long
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long

