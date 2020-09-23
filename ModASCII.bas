Attribute VB_Name = "ModASCII"
Option Explicit

Public Type tAscii
    CodeASC As Integer
    PosX As Integer
End Type

Public XaxisHor As Long
Public YaxisHor As Long
Public XaxisVer As Long
Public YaxisVer As Long

Public MyAscii(45) As tAscii

Public Sub InitializeAscii()
            
    MyAscii(36).CodeASC = Asc("1")
    MyAscii(36).PosX = 395
    
    MyAscii(37).CodeASC = Asc("2")
    MyAscii(37).PosX = 402
    
    
    MyAscii(38).CodeASC = Asc("3")
    MyAscii(38).PosX = 409
    
    
    MyAscii(39).CodeASC = Asc("4")
    MyAscii(39).PosX = 416
    
    
    MyAscii(40).CodeASC = Asc("5")
    MyAscii(40).PosX = 423
    
    
    MyAscii(41).CodeASC = Asc("6")
    MyAscii(41).PosX = 430
    
    
    MyAscii(42).CodeASC = Asc("7")
    MyAscii(42).PosX = 437
    
    
    MyAscii(43).CodeASC = Asc("8")
    MyAscii(43).PosX = 444
    
    
    MyAscii(44).CodeASC = Asc("9")
    MyAscii(44).PosX = 451
    
    
    MyAscii(45).CodeASC = Asc("0")
    MyAscii(45).PosX = 388
        
End Sub

Public Sub WriteMyAscii(hDCOrg As Long, hDCDest As Long, DestWidth As Long, DisplayText As String, Xdest As Long, Ydest As Long, ByRef FinishScroll As Boolean, Optional Leading As Integer = 0, Optional Scrolling As Boolean = False)
Dim i As Integer
Dim j As Integer
Dim CounterX As Long
Dim OnlyChar As String
Dim tempAsc As tAscii
Dim TempUcase As String
    
    If Len(DisplayText) = 0 Then
        Exit Sub
    End If

    TempUcase = UCase$(DisplayText)
    CounterX = Xdest
        
    For i = 1 To Len(TempUcase)
        OnlyChar = Mid(TempUcase, i, 1)
        For j = 1 To UBound(MyAscii)
            If Asc(OnlyChar) = MyAscii(j).CodeASC Then
                tempAsc = MyAscii(j)
                Exit For
            End If
        Next j
    
        BitBlt hDCDest, CounterX, Ydest, 6, 7, hDCOrg, tempAsc.PosX, 104, SRCCOPY
        CounterX = CounterX + 6 + Leading
    Next i
    
End Sub

Public Sub CleanASCII()
    SDC = frmMain.PicCleanASCII.hDC
    DDC = frmMain.PicWorkASCII.hDC
    BitBlt DDC, 0, 0, frmMain.PicCleanASCII.ScaleWidth, frmMain.PicCleanASCII.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub

Public Sub FlipASCII()
    SDC = frmMain.PicWorkASCII.hDC
    DDC = frmMain.PicMainASCII.hDC
    BitBlt DDC, 0, 0, frmMain.PicCleanASCII.ScaleWidth, frmMain.PicCleanASCII.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub

