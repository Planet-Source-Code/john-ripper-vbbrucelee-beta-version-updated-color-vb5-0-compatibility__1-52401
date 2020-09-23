Attribute VB_Name = "ModAPIs"
Option Explicit

'###################### CURSOR APIS #################################
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'###################### COPY MEMORY/VARPTRARRAY API #################
Public pictBuff()  As Byte
Public saBuff      As SAFEARRAY2D
Public bmpBuff     As BitMap

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Public Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

'Use this API Declaration if you dont have VB6 Runtime and only have VB5:
'Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Type BitMap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'####################### BLITS APIS ########################
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

'####################### TIME/CPERFORMANCE API #############

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'####################### SOUND API #############
Public Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Public Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Global Const SND_ASYNC = &H1
Global Const SND_FILENAME = &H20000

'###################### COLLISION IntersectRECT API
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
