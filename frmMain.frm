VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicCleanASCII 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   10
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   15
      Top             =   180
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Timer tmrLogo 
      Enabled         =   0   'False
      Left            =   600
      Top             =   3060
   End
   Begin VB.Timer tmrGameOver 
      Enabled         =   0   'False
      Left            =   1020
      Top             =   3060
   End
   Begin VB.PictureBox PicWorkASCII 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   16
      Top             =   180
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox PicMainASCII 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   14
      Top             =   180
      Width           =   4800
   End
   Begin VB.Timer tmrBaca 
      Enabled         =   0   'False
      Left            =   2700
      Top             =   3060
   End
   Begin VB.Timer tmrPuente 
      Enabled         =   0   'False
      Left            =   2280
      Top             =   3060
   End
   Begin VB.Timer tmrCataratas 
      Enabled         =   0   'False
      Left            =   180
      Top             =   3060
   End
   Begin VB.CommandButton Command4 
      Caption         =   "B"
      Height          =   315
      Left            =   2340
      TabIndex        =   13
      Top             =   2700
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Timer tmrFuegos 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1860
      Top             =   3060
   End
   Begin VB.Timer tmrLasers 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   3060
   End
   Begin VB.Timer tmrRayos 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4740
      Top             =   3420
   End
   Begin VB.Timer tmrMuere 
      Enabled         =   0   'False
      Left            =   5700
      Top             =   3180
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Muerte"
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   2700
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3"
      Height          =   315
      Index           =   2
      Left            =   3480
      TabIndex        =   10
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   9
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   315
      Index           =   0
      Left            =   2880
      TabIndex        =   8
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5700
      Top             =   2700
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5700
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Máscara"
      Height          =   315
      Left            =   3900
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   960
      ScaleHeight     =   122
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   433
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   0
      ScaleHeight     =   176
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   480
      Width           =   4800
   End
   Begin VB.PictureBox PicBuff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   0
      ScaleHeight     =   176
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox PicClean 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   0
      ScaleHeight     =   176
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox PicWork 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   0
      ScaleHeight     =   176
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image ImgGameOver 
      Height          =   105
      Left            =   1200
      Top             =   350
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   3480
      Picture         =   "frmMain.frx":0CCB
      Top             =   3420
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2100
      TabIndex        =   12
      Top             =   2700
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "xxxx"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "xxxx"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ret As Boolean

Dim BrucePlayer As cBruceLee

Dim Dx As New DirectX7
Dim objDx As DirectX7
Dim objDraw7 As DirectDraw7

Dim perfStart As Currency
Dim perfEnd As Currency
Dim perfFreq As Currency
Dim Elapsed As Double
Dim ElapsedPataSalta As Double
Dim checkSpace As Boolean
Dim checkUp As Boolean
Dim checkPuñetazo As Boolean
Dim FullScreen As Boolean
Dim StatMuere As Long
Dim DelayRayos As Boolean
Dim DelayLasers As Boolean
Dim DelayFuego As Boolean
Dim DelayCataratas As Boolean
Dim velocidadCataratas As Integer
Dim aVeloCata As Integer
Dim LQuePuente As Integer
Dim LQueBaca As Integer
Dim EnPantallaPrincipal As Boolean
Dim EstoyJugando As Boolean

Dim Ds As DirectSound
Dim DsBuffer(1) As DirectSoundBuffer
Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

Sub EndIt()
    If FullScreen = True Then
        objDraw7.RestoreDisplayMode
        objDraw7.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
        ShowCursor 1
    End If
    If CanPlayWave = True Then
        DsBuffer(0).Stop
        DsBuffer(1).Stop
        DsBuffer(0).SetCurrentPosition 0
        DsBuffer(1).SetCurrentPosition 0
    End If
    
    Set BrucePlayer = Nothing
    
    End
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Máscara" Then
        Command1.Caption = "Mapa"
        PicClean.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapMask)
    Else
        Command1.Caption = "Máscara"
        PicClean.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapPic)
    End If
    PicMain.SetFocus
End Sub



'Private Sub Command2_Click(Index As Integer)
'    BrucePlayer.Habitacion = Val(Command2(Index).Caption)
'    PicClean.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapPic)
'    PicBuff.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapMask)
'
'    GetObjectAPI frmMain.PicBuff.Picture, Len(bmpBuff), bmpBuff
'
'
'    With saBuff
'        .cbElements = 1
'        .cDims = 2
'        .Bounds(0).lLbound = 0
'        .Bounds(0).cElements = bmpBuff.bmHeight
'        .Bounds(1).lLbound = 0
'        .Bounds(1).cElements = bmpBuff.bmWidthBytes
'        .pvData = bmpBuff.bmBits
'    End With
'
'   CopyMemory ByVal VarPtrArray(pictBuff), VarPtr(saBuff), 4
'
'    PicMain.SetFocus
'End Sub
'
'Private Sub Command3_Click()
'    tmrMuere.Enabled = True
'    tmrMuere.Interval = 250
'    BrucePlayer.Muere = True
'    PicMain.SetFocus
'End Sub

'Private Sub Command4_Click()
'    If Command4.Caption = "B" Then
'        Habitaciones(4).Cataratas(1).Direccion = HaciaAbajo
'        Command4.Caption = "S"
'    Else
'        Command4.Caption = "B"
'        Habitaciones(4).Cataratas(1).Direccion = HaciaArriba
'    End If
'End Sub

'Private Sub Form_DblClick()
'EndIt
'End Sub



Private Sub Form_Load()
Dim aux As Long
    
    On Error GoTo ErrorHandler
    Unload Form1
    aux = CanPlaySound
    If aux = AUDIO_NONE Then
        CanPlayMidi = False
        CanPlayWave = False
    End If

    If aux = AUDIO_WAVE Then
        CanPlayMidi = True
        CanPlayWave = False
    End If

    If aux = AUDIO_MIDI Then
        CanPlayMidi = False
        CanPlayWave = True
    End If

    If aux = AUDIO_BOTH Then
        CanPlayMidi = True
        CanPlayWave = True
    End If
    
    FullScreen = True 'False

    InitializeAscii
    
    Show
     
    If FullScreen = True Then
        Set objDx = New DirectX7
        Set objDraw7 = objDx.DirectDrawCreate("")
        
        objDraw7.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
        objDraw7.SetDisplayMode 320, 240, 16, 0, DDSDM_STANDARDVGAMODE
        ShowCursor 0
    End If
    
    If CanPlayWave = True Then
        Initialise
    End If
    
    Preparativos

Exit Sub
ErrorHandler:
    EndIt
End Sub

Public Sub Preparativos()
    tmrLogo.Interval = 1
    tmrLogo.Enabled = True
End Sub

Public Sub Init()
    Image1.Visible = True
    EstoyJugando = True
    tmrBaca.Enabled = False
    
    On Local Error GoTo errOut
    EnPantallaPrincipal = False
    Me.BackColor = 0
    PicMainASCII.Visible = True
    
    GetObjectAPI frmMain.PicBuff.Picture, Len(bmpBuff), bmpBuff


    With saBuff
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bmpBuff.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bmpBuff.bmWidthBytes
        .pvData = bmpBuff.bmBits
    End With
    
    CopyMemory ByVal VarPtrArray(pictBuff), VarPtr(saBuff), 4

    QueryPerformanceFrequency perfFreq
    
    Do
        
        If EnPantallaPrincipal = True Then
            Exit Do
        End If
        
        
        Borra
        
'        Label2.Caption = pictBuff(BrucePlayer.PosX, bmpBuff.bmHeight - BrucePlayer.PosY) & "," & pictBuff(BrucePlayer.PosX, bmpBuff.bmHeight - BrucePlayer.PosY - 1)
'        Label3.Caption = BrucePlayer.PosX & "," & BrucePlayer.PosY
        
        BrucePlayer.ChequeaLamparas
        BrucePlayer.ChequeaFuego
        BrucePlayer.ChequeaLasers
        BrucePlayer.ChequeaRayos
        BrucePlayer.ChequeaCogerVidas
        
        WriteMyAscii Picture1.hDC, PicWorkASCII.hDC, PicWorkASCII.ScaleWidth, Format(BrucePlayer.Puntos, "000000"), 60, 1, False, 1
        If BrucePlayer.Puntos > TopScore Then
            WriteMyAscii Picture1.hDC, PicWorkASCII.hDC, PicWorkASCII.ScaleWidth, Format(BrucePlayer.Puntos, "000000"), 150, 1, False, 1
        Else
            WriteMyAscii Picture1.hDC, PicWorkASCII.hDC, PicWorkASCII.ScaleWidth, Format(TopScore, "000000"), 150, 1, False, 1
        End If
        
        If BrucePlayer.Vidas >= 0 Then
            WriteMyAscii Picture1.hDC, PicWorkASCII.hDC, PicWorkASCII.ScaleWidth, Format(BrucePlayer.Vidas, "00"), 260, 1, False, 1
        Else
            WriteMyAscii Picture1.hDC, PicWorkASCII.hDC, PicWorkASCII.ScaleWidth, "00", 260, 1, False, 1
            tmrGameOver.Enabled = True
            tmrGameOver.Interval = 50
        End If
        
        If BrucePlayer.Muere = False Then
            BrucePlayer.ChequeaPinchos
        End If

        
        OperatoriaEspecialMuros BrucePlayer.Habitacion
        OperatoriaEspecialRayos BrucePlayer.Habitacion
        OperatoriaEspecialLasers BrucePlayer.Habitacion
        
        PintaLamparas
        PintarFuegos
        PintaCataratas
        PintaPuente LQuePuente
        PintaBaca LQueBaca
        
        If BrucePlayer.Muere = True Then
            If tmrMuere.Enabled = False Then
                tmrMuere.Enabled = True
                tmrMuere.Interval = 150
            End If
        End If
        
        If BrucePlayer.Muere = False Then
            moveBruceLee
        End If
        
        If BrucePlayer.Habitacion = 15 Then
            If HaCogidoVida = False And VisitasVidas > 0 Then
                PintaVidas
            End If
        End If
        
        If BrucePlayer.CheckCambioHabitacion = True Then
            PantallaEspecial
            BrucePlayer.CayendoDiagonal = False
            If BrucePlayer.Habitacion = 13 Then
                If HaCogidoVida = True Then
                    HaCogidoVida = False
                    VisitasVidas = VisitasVidas - 1
                End If
            End If

            If (BrucePlayer.Habitacion = 1 Or BrucePlayer.Habitacion = 3) And _
                (Habitaciones(4).Lamparas(4).Estado = False And _
                Habitaciones(4).Lamparas(5).Estado = False) Then
                    tmrBaca.Interval = 1
                tmrBaca.Enabled = True
            Else
                tmrBaca.Enabled = False
            End If
                
            If BrucePlayer.Habitacion = 10 Then
                tmrPuente.Interval = 1
                tmrPuente.Enabled = True
            Else
                tmrPuente.Enabled = False
            End If
            
            
            If Habitaciones(BrucePlayer.Habitacion).NumeroRayos > 0 Then
                tmrRayos.Interval = 50
                tmrRayos.Enabled = True
            Else
                tmrRayos.Enabled = False
            End If
            
            If Habitaciones(BrucePlayer.Habitacion).NumeroLasers > 0 Then
                tmrLasers.Interval = 25
                tmrLasers.Enabled = True
            Else
                tmrLasers.Enabled = False
            End If
            
            If Habitaciones(BrucePlayer.Habitacion).NumeroFuegos > 0 Then
                tmrFuegos.Interval = 250
                tmrFuegos.Enabled = True
            Else
                tmrFuegos.Enabled = False
            End If
            
            If Habitaciones(BrucePlayer.Habitacion).NumeroCataratas > 0 Then
                tmrCataratas.Interval = 200
                tmrCataratas.Enabled = True
            Else
                tmrCataratas.Enabled = False
            End If
                
            PicClean.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapPic)
            PicBuff.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapMask)
            GetObjectAPI frmMain.PicBuff.Picture, Len(bmpBuff), bmpBuff

            With saBuff
                .cbElements = 1
                .cDims = 2
                .Bounds(0).lLbound = 0
                .Bounds(0).cElements = bmpBuff.bmHeight
                .Bounds(1).lLbound = 0
                .Bounds(1).cElements = bmpBuff.bmWidthBytes
                .pvData = bmpBuff.bmBits
            End With
             
            CopyMemory ByVal VarPtrArray(pictBuff), VarPtr(saBuff), 4
        End If
        
        If BrucePlayer.Muere = True Then
            StatMuere = StatMuere + 1
            If StatMuere Mod 2 = 0 Then
                If BrucePlayer.GameOver = False Then
                    BrucePlayer.Render
                End If
            End If
        Else
            BrucePlayer.Render
        End If
      
        Flip
        
        DoEvents
        Sleep 1
    
    Loop

    Exit Sub

errOut:
    EndIt
End Sub

Private Sub PantallaEspecial()
    If BrucePlayer.Habitacion >= 16 Then
        Exit Sub
    End If
    If Habitaciones(13).Lamparas(1).Estado = False And _
        Habitaciones(13).Lamparas(2).Estado = False And _
        Habitaciones(14).Lamparas(1).Estado = False And _
        Habitaciones(14).Lamparas(2).Estado = False And _
        Habitaciones(15).Lamparas(1).Estado = False And _
        Habitaciones(15).Lamparas(2).Estado = False Then
            BrucePlayer.PosX = 282
            BrucePlayer.PosY = 24
            BrucePlayer.Habitacion = 16
        End If
    
End Sub

Private Sub PintaVidas()
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC
    BitBlt DDC, 288, 128, 16, 32, SDC, 451, 3, SRCAND
    BitBlt DDC, 288, 128, 16, 32, SDC, 434, 3, SRCPAINT
End Sub

Private Sub PintarFuegos()
  Dim avanzaFrames As Boolean
  Dim i As Integer
  Dim X As Integer
  Dim Y As Integer
    
    If Habitaciones(BrucePlayer.Habitacion).NumeroFuegos = 0 Then
        Exit Sub
    End If
    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC
    
    If DelayFuego = True Then
        DelayFuego = False
        tmrFuegos.Enabled = False
        tmrFuegos.Enabled = True
        avanzaFrames = True
    Else
        avanzaFrames = False
    End If
    
    For i = 1 To Habitaciones(BrucePlayer.Habitacion).NumeroFuegos
        X = Habitaciones(BrucePlayer.Habitacion).Fuegos(i).PosX
        Y = Habitaciones(BrucePlayer.Habitacion).Fuegos(i).PosY
        If Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Activo = False Then
        Else
            Select Case Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Tipo
                Case 2
                    Select Case Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame
                        Case 1, 8
                            Y = Y - 4
                            BitBlt DDC, X, Y, 8, 4, SDC, 219, 69, SRCAND
                            BitBlt DDC, X, Y, 8, 4, SDC, 183, 69, SRCPAINT

                        Case 2, 7
                            Y = Y - 4
                            BitBlt DDC, X, Y, 8, 4, SDC, 228, 69, SRCAND
                            BitBlt DDC, X, Y, 8, 4, SDC, 192, 69, SRCPAINT
                        Case 3, 6
                            Y = Y - 10
                            BitBlt DDC, X, Y, 8, 10, SDC, 237, 63, SRCAND
                            BitBlt DDC, X, Y, 8, 10, SDC, 201, 63, SRCPAINT
                        
                        Case 4, 5
                            Y = Y - 20
                            BitBlt DDC, X, Y, 8, 20, SDC, 246, 53, SRCAND
                            BitBlt DDC, X, Y, 8, 20, SDC, 210, 53, SRCPAINT
                    
                    End Select
                
                Case 1
                    Select Case Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame
                        Case 1, 6
                            Y = Y - 6
                            BitBlt DDC, X, Y, 8, 6, SDC, 289, 87, SRCAND
                            BitBlt DDC, X, Y, 8, 6, SDC, 289, 74, SRCPAINT
                            
                        Case 2, 5
                            Y = Y - 10
                            BitBlt DDC, X, Y, 8, 10, SDC, 299, 83, SRCAND
                            BitBlt DDC, X, Y, 8, 10, SDC, 299, 70, SRCPAINT
                        
                        Case 3, 4
                            X = X - 4
                            Y = Y - 12
                            BitBlt DDC, X, Y, 16, 12, SDC, 308, 81, SRCAND
                            BitBlt DDC, X, Y, 16, 12, SDC, 308, 68, SRCPAINT
                        
                    End Select
            End Select
            
            If avanzaFrames = True Then
                Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Espera = Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Espera + 1
                If Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Espera >= 3 Then
                    Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame = Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame + 1
                End If
                Select Case Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Tipo
                    Case 1
                        If Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame > 6 Then
                            Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame = 0
                            Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Activo = False
                            Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Espera = 0
                        End If
                    Case 2
                        If Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame > 8 Then
                            Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Frame = 0
                            Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Activo = False
                            Habitaciones(BrucePlayer.Habitacion).Fuegos(i).Espera = 0
                        End If
                
                End Select
            End If
        End If
    Next i
End Sub

Private Sub PintaCataratas()
  Dim i As Integer
  Dim lHab As Integer
  Dim X As Integer
  Dim Y As Integer
  Dim j As Integer
  Dim alturaSPRcatarata As Integer
  Dim anchuraSPRcatarata As Integer
  Dim posXSPRcatarata As Integer
  Dim posYSPRcatarata As Integer
  Dim RealVelo As Integer
  Dim nRepeticiones As Integer
  Dim TempValue As Integer
  Dim yTemp As Integer
    
    lHab = BrucePlayer.Habitacion
    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC
    
    If Habitaciones(lHab).NumeroCataratas = 0 Then
        Exit Sub
    End If
    
    If DelayCataratas = True Then
        DelayCataratas = False
        tmrCataratas.Enabled = False
        tmrCataratas.Enabled = True
        RealVelo = velocidadCataratas
        'Label5.Caption = RealVelo
    Else
        RealVelo = 0
    End If
    
    For i = 1 To Habitaciones(lHab).NumeroCataratas
        Select Case Habitaciones(lHab).Cataratas(i).Tipo
            Case 1
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 24
                posXSPRcatarata = 289
                posYSPRcatarata = 94
            Case 2
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 32
                posXSPRcatarata = 289
                posYSPRcatarata = 103
            Case 3
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 48
                posXSPRcatarata = 289
                posYSPRcatarata = 112
            Case 4
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 40
                posXSPRcatarata = 289
                posYSPRcatarata = 112
            Case 5
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 56
                posXSPRcatarata = 338
                posYSPRcatarata = 112
            Case 6
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 32
                posXSPRcatarata = 322
                posYSPRcatarata = 103
            Case 7
                alturaSPRcatarata = 8
                anchuraSPRcatarata = 32
                posXSPRcatarata = 355
                posYSPRcatarata = 103
        
        End Select
    
        nRepeticiones = (Habitaciones(lHab).Cataratas(i).PosY2 - Habitaciones(lHab).Cataratas(i).PosY) \ alturaSPRcatarata
        
        If Habitaciones(lHab).Cataratas(i).Direccion = HaciaArriba Then
            Habitaciones(lHab).Cataratas(i).TempY = Habitaciones(lHab).Cataratas(i).TempY - RealVelo
            If Abs(Habitaciones(lHab).Cataratas(i).TempY) > alturaSPRcatarata Then
                Habitaciones(lHab).Cataratas(i).TempY = 0
            End If
        Else
            Habitaciones(lHab).Cataratas(i).TempY = Habitaciones(lHab).Cataratas(i).TempY + RealVelo
            If Habitaciones(lHab).Cataratas(i).TempY >= 0 Then
                Habitaciones(lHab).Cataratas(i).TempY = alturaSPRcatarata * -1
            End If
        
        End If
        TempValue = Habitaciones(lHab).Cataratas(i).TempY
        X = Habitaciones(lHab).Cataratas(i).PosX
        Y = Habitaciones(lHab).Cataratas(i).PosY
        
        For j = 0 To nRepeticiones + 1
                If j = 0 Then
                    BitBlt DDC, X, Y, anchuraSPRcatarata, alturaSPRcatarata + TempValue, SDC, posXSPRcatarata, (posYSPRcatarata + ((TempValue) * (-1))), SRCAND
                    yTemp = Y + alturaSPRcatarata + TempValue

                Else
                    If yTemp + alturaSPRcatarata > Habitaciones(lHab).Cataratas(i).PosY2 Then
                        BitBlt DDC, X, yTemp, anchuraSPRcatarata, Habitaciones(lHab).Cataratas(i).PosY2 - yTemp, SDC, posXSPRcatarata, posYSPRcatarata, SRCAND
                        Exit For
                    Else
                        BitBlt DDC, X, yTemp, anchuraSPRcatarata, alturaSPRcatarata, SDC, posXSPRcatarata, posYSPRcatarata, SRCAND
                        yTemp = yTemp + alturaSPRcatarata
                    End If
                End If
        Next j
    Next i
End Sub

Private Sub PintaLamparas()
  Dim i As Integer
  Dim lHab As Integer
    
    lHab = BrucePlayer.Habitacion
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC

    For i = 1 To Habitaciones(lHab).NumeroLamparas
    If Habitaciones(lHab).Lamparas(i).Estado = True Then
        Select Case Habitaciones(lHab).Lamparas(i).Tipo
            Case 0
                BitBlt DDC, Habitaciones(lHab).Lamparas(i).PosX, Habitaciones(lHab).Lamparas(i).PosY, 8, 10, SDC, 242, 2, SRCAND
                BitBlt DDC, Habitaciones(lHab).Lamparas(i).PosX, Habitaciones(lHab).Lamparas(i).PosY, 8, 10, SDC, 233, 2, SRCPAINT
            Case 1
                BitBlt DDC, Habitaciones(lHab).Lamparas(i).PosX, Habitaciones(lHab).Lamparas(i).PosY, 8, 10, SDC, 242, 2, SRCAND
                BitBlt DDC, Habitaciones(lHab).Lamparas(i).PosX, Habitaciones(lHab).Lamparas(i).PosY, 8, 10, SDC, 251, 2, SRCPAINT
            Case 2
                BitBlt DDC, Habitaciones(lHab).Lamparas(i).PosX, Habitaciones(lHab).Lamparas(i).PosY, 6, 12, SDC, 267, 2, SRCAND
                BitBlt DDC, Habitaciones(lHab).Lamparas(i).PosX, Habitaciones(lHab).Lamparas(i).PosY, 6, 12, SDC, 260, 2, SRCPAINT
        End Select
    End If
Next i
End Sub

'Private Sub PicBuff_Click()
'EndIt
'End Sub

Private Sub Borra()
    CleanASCII
    SDC = PicClean.hDC
    DDC = PicWork.hDC
    BitBlt DDC, 0, 0, PicClean.ScaleWidth, PicClean.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub
Private Sub Flip()
    FlipASCII
    SDC = PicWork.hDC
    DDC = PicMain.hDC
    BitBlt DDC, 0, 0, PicWork.ScaleWidth, PicWork.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub

Private Sub Render()
    
    SDC = Picture1.hDC
    DDC = PicWork.hDC
    BitBlt DDC, 0, 0, 16, 23, SDC, 1, 25, SRCAND
    BitBlt DDC, 0, 0, 16, 23, SDC, 1, 1, SRCPAINT

End Sub
Private Sub OperatoriaEspecialLasers(queHabitacion As Integer)
  Dim Velo As Integer
  Dim i As Integer
  Dim X As Integer
  Dim Y As Integer
    
    If Habitaciones(queHabitacion).NumeroLasers = 0 Then
        Exit Sub
    End If

    If DelayLasers = True Then
        DelayLasers = False
        tmrLasers.Enabled = False
        tmrLasers.Enabled = True
        Velo = 5
    Else
        Velo = 0
    End If
    
    If BrucePlayer.Muere = True Then
        Velo = 0
    End If
    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC

    For i = 1 To Habitaciones(queHabitacion).NumeroLasers
        If Habitaciones(queHabitacion).Laser(i).Stop = True Then
            Velo = 0
        End If
        
        If Habitaciones(queHabitacion).Laser(i).Direccion = deIzquierdaADerecha Then
            Habitaciones(queHabitacion).Laser(i).tempX = Habitaciones(queHabitacion).Laser(i).tempX + Velo
            If Habitaciones(queHabitacion).Laser(i).tempX > (Habitaciones(queHabitacion).Laser(i).PosX1 + Habitaciones(queHabitacion).Laser(i).Distancia - 8) Then
                Habitaciones(queHabitacion).Laser(i).tempX = Habitaciones(queHabitacion).Laser(i).PosX1
            End If
        Else
            Habitaciones(queHabitacion).Laser(i).tempX = Habitaciones(queHabitacion).Laser(i).tempX - Velo
            If Habitaciones(queHabitacion).Laser(i).tempX < (Habitaciones(queHabitacion).Laser(i).PosX1) Then
                Habitaciones(queHabitacion).Laser(i).tempX = Habitaciones(queHabitacion).Laser(i).PosX1 + Habitaciones(queHabitacion).Laser(i).Distancia - 8
            End If
        End If
    
        X = Habitaciones(queHabitacion).Laser(i).tempX
        Y = Habitaciones(queHabitacion).Laser(i).PosY1
    
        BitBlt DDC, X, Y, 8, 2, SDC, 255, 53, SRCCOPY
    Next i
End Sub

Private Sub OperatoriaEspecialRayos(queHabitacion As Integer)
  Dim i As Integer
  Dim X As Integer
  Dim Y As Integer
  Dim Velo As Integer
  Dim VeloReal As Integer
    
    If Habitaciones(queHabitacion).NumeroRayos = 0 Then
        Exit Sub
    End If
    
    If DelayRayos = True Then
        Velo = 1
        DelayRayos = False
        tmrRayos.Enabled = False
        tmrRayos.Enabled = True
    Else
        Velo = 0
    End If
        
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC

    For i = 1 To Habitaciones(queHabitacion).NumeroRayos
        If Velo = 1 Then
            VeloReal = Habitaciones(queHabitacion).Rayos(i).Velo
        Else
            VeloReal = 0
        End If
        
        If Habitaciones(queHabitacion).Rayos(i).Stop = True Then
            VeloReal = 0
        End If
        
        Habitaciones(queHabitacion).Rayos(i).tempX = Habitaciones(queHabitacion).Rayos(i).tempX + VeloReal
        X = Habitaciones(queHabitacion).Rayos(i).tempX
        Y = Habitaciones(queHabitacion).Rayos(i).PosY1
        
        If X > Habitaciones(queHabitacion).Rayos(i).InvisX2 Then
            Habitaciones(queHabitacion).Rayos(i).tempX = Habitaciones(queHabitacion).Rayos(i).InvisX1
        End If
        
        If X >= Habitaciones(queHabitacion).Rayos(i).PosX1 And X <= Habitaciones(queHabitacion).Rayos(i).PosX2 Then
            BitBlt DDC, X, Y, 8, 10, SDC, 265, 60, SRCAND
            BitBlt DDC, X, Y, 8, 10, SDC, 265, 49, SRCPAINT
            Habitaciones(queHabitacion).Rayos(i).Visible = True
        Else
            Habitaciones(queHabitacion).Rayos(i).Visible = False
        End If
    Next i
    
    Select Case queHabitacion
        Case 4
            BitBlt DDC, 0, 55, 8, 10, SDC, 343, 68, SRCCOPY
            BitBlt DDC, 32, 54, 16, 11, SDC, 326, 68, SRCCOPY
            BitBlt DDC, 192, 114, 24, 12, SDC, 326, 80, SRCCOPY
            BitBlt DDC, 240, 114, 10, 14, SDC, 351, 78, SRCCOPY
        Case 5
            BitBlt DDC, 0, 148, 8, 10, SDC, 362, 78, SRCCOPY
            BitBlt DDC, 40, 148, 8, 10, SDC, 362, 78, SRCCOPY
            BitBlt DDC, 80, 84, 8, 10, SDC, 362, 78, SRCCOPY
            BitBlt DDC, 120, 84, 8, 10, SDC, 362, 78, SRCCOPY
            BitBlt DDC, 192, 84, 8, 10, SDC, 362, 78, SRCCOPY
            BitBlt DDC, 232, 84, 8, 10, SDC, 362, 78, SRCCOPY
        Case 7
            BitBlt DDC, 40, 52, 8, 10, SDC, 352, 67, SRCCOPY
            BitBlt DDC, 104, 52, 8, 10, SDC, 352, 67, SRCCOPY
        Case 8
            BitBlt DDC, 120, 68, 8, 10, SDC, 362, 89, SRCCOPY
            BitBlt DDC, 176, 68, 8, 10, SDC, 352, 67, SRCCOPY
    End Select

End Sub
Private Sub OperatoriaEspecialMuros(queHabitacion As Integer)
On Error GoTo ErrorHandler

    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC
    
    Select Case queHabitacion
        Case 2
            If TotalLamparasRecogidas = 22 Then '22 Then
                ModificarMascaraJIT 149, 170, 170, 173
                BitBlt DDC, 134, 164, 16, 6, SDC, 274, 9, SRCAND
                BitBlt DDC, 134, 164, 16, 6, SDC, 274, 2, SRCPAINT
            Else
                BitBlt DDC, 152, 168, 16, 6, SDC, 274, 9, SRCAND
                BitBlt DDC, 152, 168, 16, 6, SDC, 274, 2, SRCPAINT
            
            End If
        
        Case 3
            If Habitaciones(4).Lamparas(4).Estado = False And Habitaciones(4).Lamparas(5).Estado = False Then
                ModificarMascaraJIT 318, 128, 319, 170
            Else
                BitBlt DDC, 312, 128, 8, 42, SDC, 407, 49, SRCAND
            End If
        
        Case 4
            If Habitaciones(4).Lamparas(7).Estado = False And Habitaciones(4).Lamparas(8).Estado = False Then
                ModificarMascaraJIT 0, 124, 7, 166
            Else
                BitBlt DDC, 0, 124, 8, 42, SDC, 407, 49, SRCAND
            End If
            
            If Habitaciones(4).Lamparas(9).Estado = False And Habitaciones(4).Lamparas(10).Estado = False Then
                ModificarMascaraJIT 200, 128, 207, 166
            Else
                BitBlt DDC, 200, 128, 8, 38, SDC, 416, 49, SRCAND
            End If
            If Habitaciones(4).Lamparas(4).Estado = False And _
                Habitaciones(4).Lamparas(5).Estado = False Then
                    ModificarMascaraJIT 272, 6, 295, 16
              Else
                BitBlt DDC, 273, 0, 22, 16, SDC, 373, 79, SRCAND
                BitBlt DDC, 273, 0, 22, 16, SDC, 373, 62, SRCPAINT
              End If
        
        Case 5
            If Habitaciones(5).Lamparas(1).Estado = False Then
                ModificarMascaraJIT 88, 110, 119, 112
            Else
                BitBlt DDC, 88, 110, 32, 16, SDC, 307, 49, SRCAND
                BitBlt DDC, 88, 110, 32, 16, SDC, 274, 49, SRCPAINT
            End If
            If Habitaciones(5).Lamparas(2).Estado = False Then
                ModificarMascaraJIT 200, 110, 231, 112
            Else
                BitBlt DDC, 200, 110, 32, 12, SDC, 373, 49, SRCAND
                BitBlt DDC, 200, 110, 32, 12, SDC, 340, 49, SRCPAINT
            End If
        
        Case 6
            If Habitaciones(6).Lamparas(1).Estado = False And _
                Habitaciones(6).Lamparas(2).Estado = False And _
                Habitaciones(6).Lamparas(3).Estado = False And _
                Habitaciones(6).Lamparas(4).Estado = False And _
                Habitaciones(6).Lamparas(5).Estado = False Then
                ModificarMascaraJIT 272, 1, 279, 42
            Else
                BitBlt DDC, 272, 0, 8, 42, SDC, 407, 49, SRCAND
            End If
    
            If Habitaciones(6).Lamparas(6).Estado = False And _
                Habitaciones(6).Lamparas(7).Estado = False And _
                Habitaciones(6).Lamparas(8).Estado = False And _
                Habitaciones(6).Lamparas(9).Estado = False And _
                Habitaciones(6).Lamparas(10).Estado = False And _
                Habitaciones(6).Lamparas(11).Estado = False Then
                ModificarMascaraJIT 312, 48, 319, 90
            Else
                BitBlt DDC, 312, 48, 8, 42, SDC, 407, 49, SRCAND
            End If
            
            If Habitaciones(6).Lamparas(12).Estado = False And _
               Habitaciones(6).Lamparas(13).Estado = False Then
                Habitaciones(6).Laser(1).Direccion = deDerechaAIzquierda
            End If
        
        Case 7
            If Habitaciones(6).Lamparas(12).Estado = False Then
                ModificarMascaraJIT 112, 80, 113, 118
            Else
                BitBlt DDC, 112, 80, 8, 38, SDC, 416, 49, SRCAND
            End If
    
            If Habitaciones(7).Lamparas(1).Estado = False Then
                ModificarMascaraJIT 0, 112, 1, 166
            Else
                BitBlt DDC, 0, 112, 8, 54, SDC, 425, 49, SRCAND
            End If
            
            If Habitaciones(7).Lamparas(2).Estado = False And Habitaciones(7).Lamparas(3).Estado = False Then
                ModificarMascaraJIT 200, 80, 201, 118
            Else
                BitBlt DDC, 200, 80, 8, 38, SDC, 416, 49, SRCAND
            End If
    
            If Habitaciones(7).Lamparas(4).Estado = False Then
                ModificarMascaraJIT 318, 112, 319, 166
            Else
                BitBlt DDC, 312, 112, 8, 54, SDC, 425, 49, SRCAND
            End If
            If Habitaciones(8).Lamparas(3).Estado = False And _
                Habitaciones(8).Lamparas(4).Estado = False Then
                ModificarMascaraJIT 318, 1, 319, 44
            Else
                BitBlt DDC, 312, 0, 8, 44, SDC, 425, 49, SRCAND
            End If
            
        Case 8
            If Habitaciones(8).Lamparas(3).Estado = False And _
                Habitaciones(8).Lamparas(4).Estado = False Then
                Habitaciones(8).Laser(1).Direccion = deDerechaAIzquierda
            End If

            If Habitaciones(8).Lamparas(2).Estado = False Then
                ModificarMascaraJIT 184, 78, 185, 118
            Else
                BitBlt DDC, 184, 78, 8, 44, SDC, 425, 49, SRCAND
            End If
            
            If Habitaciones(8).Lamparas(1).Estado = False And _
                Habitaciones(8).Lamparas(5).Estado = False Then
                ModificarMascaraJIT 280, 6, 311, 16
            Else
                BitBlt DDC, 280, 0, 30, 16, SDC, 373, 79, SRCAND
                BitBlt DDC, 280, 0, 30, 16, SDC, 373, 62, SRCPAINT
            End If
        
        Case 9
            If Habitaciones(9).Lamparas(1).Estado = False And _
                Habitaciones(9).Lamparas(2).Estado = False And _
                Habitaciones(9).Lamparas(3).Estado = False And _
                Habitaciones(9).Lamparas(4).Estado = False And _
                Habitaciones(9).Lamparas(5).Estado = False And _
                Habitaciones(9).Lamparas(6).Estado = False And _
                Habitaciones(9).Lamparas(7).Estado = False And _
                Habitaciones(9).Lamparas(8).Estado = False Then
                ModificarMascaraJIT 0, 44, 1, 88
            Else
                BitBlt DDC, 0, 46, 8, 44, SDC, 425, 49, SRCAND
            End If
        
        Case 10
            If Habitaciones(10).Lamparas(1).Estado = False And _
                Habitaciones(10).Lamparas(2).Estado = False Then
                ModificarMascaraJIT 244, 170, 267, 173
            Else
                BitBlt DDC, 248, 168, 16, 6, SDC, 274, 9, SRCAND
                BitBlt DDC, 248, 168, 16, 6, SDC, 274, 2, SRCPAINT
            End If
        
        Case 11
            If Habitaciones(11).Lamparas(1).Estado = False And _
                Habitaciones(11).Lamparas(2).Estado = False Then
                ModificarMascaraJIT 318, 16, 319, 67
            Else
                BitBlt DDC, 312, 16, 8, 49, SDC, 434, 49, SRCAND
            End If
        
        Case 13
            If Habitaciones(13).Lamparas(3).Estado = False Then
                ModificarMascaraJIT 309, 128, 310, 162
            Else
                BitBlt DDC, 312, 128, 8, 36, SDC, 434, 49, SRCAND
            End If
            
            If Habitaciones(13).Lamparas(1).Estado = False And _
                Habitaciones(13).Lamparas(2).Estado = False Then
            
                ModificarMascaraJIT 309, 80, 310, 115
            Else
                BitBlt DDC, 312, 80, 8, 36, SDC, 434, 49, SRCAND
            End If
        
        Case 14
            If Habitaciones(14).Lamparas(1).Estado = False And _
                Habitaciones(14).Lamparas(2).Estado = False Then
            
                ModificarMascaraJIT 309, 17, 310, 66
            Else
                BitBlt DDC, 312, 17, 8, 49, SDC, 434, 49, SRCAND
            End If
        
        Case 15
            If Habitaciones(15).Lamparas(1).Estado = False And _
                Habitaciones(15).Lamparas(2).Estado = False Then
            
                ModificarMascaraJIT 9, 80, 10, 114
            Else
                BitBlt DDC, 0, 80, 8, 36, SDC, 434, 49, SRCAND
            End If
        
        Case 17
            If Habitaciones(17).Lamparas(1).Estado = False And _
                Habitaciones(17).Lamparas(2).Estado = False Then
                ModificarMascaraJIT 142, 16, 143, 42
            Else
                BitBlt DDC, 136, 16, 8, 26, SDC, 444, 49, SRCAND
            End If
            If Habitaciones(17).Lamparas(5).Estado = False Then
                ModificarMascaraJIT 184, 16, 185, 42
            Else
                BitBlt DDC, 184, 16, 8, 26, SDC, 453, 49, SRCAND
            End If
            
            If Habitaciones(17).Lamparas(6).Estado = False Then
                ModificarMascaraJIT 309, 80, 310, 124
            Else
                BitBlt DDC, 312, 80, 8, 44, SDC, 462, 49, SRCAND
            End If
        
        Case 18
            If Habitaciones(18).Lamparas(1).Estado = False Then
                BitBlt DDC, 296, 112, 16, 16, SDC, 471, 98, SRCAND
                BitBlt DDC, 296, 112, 16, 16, SDC, 471, 65, SRCPAINT
                ModificarMascaraJIT 302, 112, 305, 128, 1
            End If
            
            If Habitaciones(18).Lamparas(2).Estado = False Then
                BitBlt DDC, 192, 112, 24, 32, SDC, 471, 82, SRCAND
                BitBlt DDC, 192, 112, 24, 32, SDC, 471, 49, SRCPAINT
                ModificarMascaraJIT 192, 112, 216, 144, 1
            End If
    End Select
Exit Sub
ErrorHandler:
    Debug.Print Err.Number & ":" & Err.Description
    'Resume
End Sub

Private Sub ModificarMascaraJIT(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional Valor As Integer = 4)
Dim i As Integer
Dim j As Integer
    For i = X1 To X2
        For j = bmpBuff.bmHeight - Y2 To bmpBuff.bmHeight - Y1
            pictBuff(i, j) = Valor
        Next j
    Next i
End Sub

Private Sub PicMain_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
    If EstoyJugando = True Then
        If BrucePlayer.Muere = True Then
            If BrucePlayer.GameOver = False Then
                Exit Sub
            End If
        End If
    End If
    
    Select Case KeyCode
        Case 65 'A (arriba)
        
            If BrucePlayer.Accion = SaltoDiagonal Then
                If checkUp = True Then
                Else
                    checkUp = False
                    pUp = True
                
                End If
            Else
                pUp = True
            End If
            
        Case 90 'Z (abajo)
            pDown = True
        Case 188 ', (izq)
            pLeft = True
        Case 190 '. (derecha)
            pRight = True
        Case 32 'barra espaciadora
            'Debug.Print checkSpace
            If checkSpace = True Then
            Else
                checkSpace = False
                pSpaceBar = True
            End If
    End Select
End Sub

Private Sub PicMain_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If EstoyJugando = True Then
        If BrucePlayer.Muere = True Then
            If BrucePlayer.GameOver = False Then
                Exit Sub
            End If
        End If
    End If
    
    Select Case KeyCode
        Case 13 'Ent
            If EnPantallaPrincipal = True Then
                BrucePlayer.Vidas = 4
                Init
                
            End If
        Case 27 'ESC
            If EnPantallaPrincipal = True Then
                EndIt
            Else
                If EstoyJugando = True Then
                    If BrucePlayer.Puntos > TopScore Then
                        TopScore = BrucePlayer.Puntos
                    End If
                End If
                PantallaPrincipal
            End If
        Case 65 'A (arriba)
            pUp = False
        Case 90 'Z (abajo)
            pDown = False
        Case 188 ', (izq)
            pLeft = False
        Case 190 '. (derecha)
            pRight = False
        Case 32 'barra espaciadora
            pSpaceBar = False
    
    End Select
End Sub

Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Solo para debugger
    BrucePlayer.PosX = X
    BrucePlayer.PosY = Y
End Sub


Private Sub moveBruceLee()
  Static test As Single
  Static autoYcata As Single
  Static playRun As Integer
  Dim retval As Long
  Dim sPan As Single
  Dim lPan As Long


    Timer
    test = test + Elapsed

    If test < 10 Then
        Exit Sub
    Else
        test = 0
    End If
    
    playRun = playRun + 1
    
    If BrucePlayer.Accion = Correr Then
        If playRun > 10 Then
            If CanPlayWave = True Then
                Select Case BrucePlayer.PosX
                    Case 0 To 32
                        lPan = -1000
                    Case 33 To 65
                        lPan = -800
                    Case 65 To 97
                        lPan = -600
                    Case 98 To 130
                        lPan = -400
                    Case 131 To 163
                        lPan = -200
                    Case 164 To 196
                        lPan = 200
                    Case 197 To 229
                        lPan = 400
                    Case 230 To 262
                        lPan = 600
                    Case 263 To 295
                        lPan = 800
                    Case 296 To 320
                        lPan = 1000
                End Select
                    
                'retval = PlaySound(App.Path & "\sonidos\run.wav", 0, SND_FILENAME Or SND_ASYNC)
                DsBuffer(0).SetPan lPan
                DsBuffer(0).Play DSBPLAY_DEFAULT
            End If
            playRun = 0
        End If
    End If

   
   '####################################################
    If BrucePlayer.Habitacion = 18 Then
        If BrucePlayer.PosX < 226 And BrucePlayer.PosY = 85 Then
            pUp = False
        End If
        If BrucePlayer.PosX > 272 And BrucePlayer.PosY <= 22 Then
            pRight = False
        End If
            
    End If
    
    '####################################################
    If BrucePlayer.EstaSobreCatarata = True Then
        Select Case velocidadCataratas
            Case 1
                autoYcata = autoYcata + 0.1
            Case 2
                autoYcata = autoYcata + 0.2
            Case 3
                autoYcata = autoYcata + 0.5
        End Select
        If autoYcata > 1 Then
            autoYcata = 0
        End If
        If Habitaciones(BrucePlayer.Habitacion).Cataratas(1).Direccion = HaciaAbajo Then
            BrucePlayer.PosY = BrucePlayer.PosY + autoYcata
            If BrucePlayer.PosY > 174 Then
                BrucePlayer.PosY = 174
            End If
        Else
            BrucePlayer.PosY = BrucePlayer.PosY - autoYcata
            If BrucePlayer.PosY < 22 Then
                BrucePlayer.PosY = 22
            End If
        End If
        If autoYcata > 1 Then
            autoYcata = 0
        End If
        If pDown = True And BrucePlayer.Accion = Quieto Then
            BrucePlayer.Accion = Escalar
            Exit Sub
        End If
    End If
    
    '####################################################
    If BrucePlayer.Accion = Escalar Then
        If pUp = True Then
            If BrucePlayer.PosY <= 22 Then
                Exit Sub
            End If
            BrucePlayer.PosY = BrucePlayer.PosY - 1
            If BrucePlayer.EstaEscalando = False Then
                BrucePlayer.Accion = Caer
            End If
        
        ElseIf pDown = True Then
            If BrucePlayer.PosY > 174 Then
                Exit Sub
            End If
            BrucePlayer.PosY = BrucePlayer.PosY + 1
            If BrucePlayer.EstaEscalando = False Then
                BrucePlayer.Accion = Caer
            End If
        ElseIf pLeft = True Then
            BrucePlayer.PosX = BrucePlayer.PosX - 1
            If BrucePlayer.EstaEscalando = False Then
                BrucePlayer.Accion = Caer
            End If
        ElseIf pRight = True Then
            BrucePlayer.PosX = BrucePlayer.PosX + 1
            If BrucePlayer.EstaEscalando = False Then
                BrucePlayer.Accion = Caer
            End If
        End If
        Exit Sub
    End If
    
    '####################################################
    If BrucePlayer.Accion <> Patada And BrucePlayer.Accion <> Saltar And BrucePlayer.Accion <> SaltoDiagonal And BrucePlayer.Accion <> Escalar Then
        If BrucePlayer.EstaCayendo = True Then
            BrucePlayer.Accion = Caer
            BrucePlayer.PosY = BrucePlayer.PosY + 1
            Exit Sub
        End If
    End If
    
    '####################################################
    If BrucePlayer.Accion = Correr And pUp = True Then
        BrucePlayer.Accion = SaltoDiagonal
        BrucePlayer.OrigenSaltoDiagonalY = BrucePlayer.PosY
        Exit Sub
    End If
    
    '####################################################
    If BrucePlayer.Accion = Quieto And pSpaceBar = True Then
        BrucePlayer.Accion = Puñetazo
        checkPuñetazo = True
        Timer2.Enabled = True
        Exit Sub
    End If
    
    '####################################################
    If BrucePlayer.Accion = Puñetazo And checkPuñetazo = False Then
        BrucePlayer.Accion = Quieto
        checkSpace = True
        pSpaceBar = False
        Timer1.Enabled = True
    End If
    
    '####################################################
    If BrucePlayer.Accion = Correr And pSpaceBar = True Then
        BrucePlayer.Accion = Patada
        BrucePlayer.OrigenPatada = BrucePlayer.PosX
        Exit Sub
    End If
    
    '####################################################
    If BrucePlayer.Accion = Quieto And pUp = True Then
        BrucePlayer.Accion = Saltar
        BrucePlayer.OrigenSalto = BrucePlayer.PosY
        Exit Sub
    End If
    
    '####################################################
    If BrucePlayer.Accion = SaltoDiagonal Then
        If BrucePlayer.Direccion = Derecha Then
            If BrucePlayer.CayendoDiagonal = False Then
                BrucePlayer.PosX = BrucePlayer.PosX + 1
                If Abs(BrucePlayer.PosY - BrucePlayer.OrigenSaltoDiagonalY) <= 13 Then
                    BrucePlayer.PosY = BrucePlayer.PosY - 1
                Else
                    BrucePlayer.CayendoDiagonal = True
                    BrucePlayer.OrigenSaltoDiagonalX = BrucePlayer.PosX
                End If
            Else
                If BrucePlayer.PosX - BrucePlayer.OrigenSaltoDiagonalX < 5 Then
                    BrucePlayer.PosX = BrucePlayer.PosX + 1
                Else
                    BrucePlayer.PosX = BrucePlayer.PosX + 1
                    BrucePlayer.PosY = BrucePlayer.PosY + 1
                End If
            End If
            If BrucePlayer.EstaEscalando = True Then
                BrucePlayer.CayendoDiagonal = False
                checkUp = True
                pUp = False
                BrucePlayer.Accion = Escalar
                Exit Sub
            End If
            If BrucePlayer.ImposibleMoverDerecha = True Then
                BrucePlayer.Accion = Quieto
                BrucePlayer.SubeUnaLinea
                BrucePlayer.CayendoDiagonal = False
                checkUp = True
                pUp = False
                Timer1.Enabled = True
            End If
        Else
            If BrucePlayer.CayendoDiagonal = False Then
                BrucePlayer.PosX = BrucePlayer.PosX - 1
                If Abs(BrucePlayer.PosY - BrucePlayer.OrigenSaltoDiagonalY) <= 13 Then
                    BrucePlayer.PosY = BrucePlayer.PosY - 1
                Else
                    BrucePlayer.CayendoDiagonal = True
                    BrucePlayer.OrigenSaltoDiagonalX = BrucePlayer.PosX
                End If
            Else
                If BrucePlayer.OrigenSaltoDiagonalX - BrucePlayer.PosX < 5 Then
                    BrucePlayer.PosX = BrucePlayer.PosX - 1
                    BrucePlayer.OrigenSaltoDiagonalY = BrucePlayer.PosY
                Else
                    BrucePlayer.PosX = BrucePlayer.PosX - 1
                    BrucePlayer.PosY = BrucePlayer.PosY + 1
                End If
            End If
            If BrucePlayer.EstaEscalando = True Then
                BrucePlayer.Accion = Escalar
                BrucePlayer.CayendoDiagonal = False
                checkUp = True
                pUp = False
                
                Exit Sub
            End If
            If BrucePlayer.ImposibleMoverIzquierda = True Then
                BrucePlayer.PosY = BrucePlayer.PosY - 1
                BrucePlayer.Accion = Quieto
                BrucePlayer.SubeUnaLinea
                BrucePlayer.CayendoDiagonal = False
                checkUp = True
                pUp = False
                Timer1.Enabled = True
            End If
        End If
        Exit Sub
    End If
    
    
    '####################################################
    If BrucePlayer.Accion = Patada Then
        If BrucePlayer.Direccion = Derecha Then
            BrucePlayer.PosX = BrucePlayer.PosX + 1
            If BrucePlayer.ImposibleMoverDerecha = True Then
                BrucePlayer.Accion = Quieto
                checkSpace = True
                pSpaceBar = False
                Timer1.Enabled = True
            End If
        Else
            BrucePlayer.PosX = BrucePlayer.PosX - 1
            If BrucePlayer.ImposibleMoverIzquierda = True Then
                BrucePlayer.Accion = Quieto
                checkSpace = True
                pSpaceBar = False
                Timer1.Enabled = True
            End If
        End If
    
        If Abs(BrucePlayer.PosX - BrucePlayer.OrigenPatada) >= 40 Then
            If BrucePlayer.EstaCayendo = True Then
                BrucePlayer.Accion = Caer
                checkSpace = True
                Timer1.Enabled = True
                pSpaceBar = False
            Else
                BrucePlayer.Accion = Quieto
                checkSpace = True
                pSpaceBar = False
                Timer1.Enabled = True
            End If
        End If
        Exit Sub
    End If
    
    '####################################################
    If BrucePlayer.Accion = Saltar Then
        BrucePlayer.PosY = BrucePlayer.PosY - 1
        If Abs(BrucePlayer.PosY - BrucePlayer.OrigenSalto) >= 15 Then
            BrucePlayer.Accion = Caer
        End If
        If BrucePlayer.EstaEscalando = True Then
            BrucePlayer.Accion = Escalar
        End If
        Exit Sub
    End If
    
    '####################################################
    If pDown = True And (BrucePlayer.Accion = Agacharse Or BrucePlayer.Accion = Quieto Or BrucePlayer.Accion = Correr) Then
        BrucePlayer.Accion = Agacharse
        Exit Sub
    End If
    
    '####################################################
    If pDown = False And BrucePlayer.Accion = Agacharse Then
        BrucePlayer.Accion = Quieto
    End If
    
    '####################################################
    If pLeft = True Then
        If BrucePlayer.Accion <> Correr Then
            BrucePlayer.Accion = Correr
        End If
        If BrucePlayer.ImposibleMoverIzquierda = True Then
            BrucePlayer.Accion = Quieto
        Else
            BrucePlayer.PosX = BrucePlayer.PosX - 1
            BrucePlayer.Direccion = Izquieda
        End If
    End If
    
    '####################################################
    If pRight = True Then
        If BrucePlayer.Accion <> Correr Then
            BrucePlayer.Accion = Correr
        End If
        If BrucePlayer.ImposibleMoverDerecha = True Then
            BrucePlayer.Accion = Quieto
        Else
            BrucePlayer.PosX = BrucePlayer.PosX + 1
            BrucePlayer.Direccion = Derecha
        End If
    End If

    '####################################################
    If pRight = False And pLeft = False And pDown = False And BrucePlayer.Accion <> Puñetazo Then
        BrucePlayer.Accion = Quieto
    End If

End Sub
Private Sub Timer()
    QueryPerformanceCounter perfStart
    Elapsed = (perfStart - perfEnd) / perfFreq * 1000
    QueryPerformanceCounter perfEnd
End Sub

Private Sub Timer1_Timer()
    'Timer para evitar las repeticiones del salto y patada
    Timer1.Enabled = False
    checkSpace = False
    checkUp = False
End Sub

Private Sub Timer2_Timer()
    'timer para evitar las repeticiones en el puñetazo
    Timer2.Enabled = False
    checkPuñetazo = False
    checkSpace = False
End Sub

Private Sub tmrBaca_Timer()
  Static QueBaca As Integer
  Dim retval As Long
    
    'timer que hace sonar el cow.wav pantalla 1 y 3
    
    If tmrBaca.Interval < 1500 Then
        tmrBaca.Interval = 1500
    End If
    
    QueBaca = QueBaca + 1
    If QueBaca = 1 Then
        If CanPlayWave = True Then
            'retval = PlaySound(App.Path & "\sonidos\cow.wav", 0, SND_FILENAME Or SND_ASYNC)
            DsBuffer(1).Play DSBPLAY_DEFAULT
        End If
    Else
        QueBaca = 0
    End If
    
    LQueBaca = QueBaca
End Sub

Private Sub tmrCataratas_Timer()
  'timer para el retardo de la animacion de las cataratas
  Static contadorVelocidad As Integer
  Static subeobaja As Integer
  Dim i As Integer

    DelayCataratas = True
    contadorVelocidad = contadorVelocidad + 1
    
    If contadorVelocidad >= 25 Then
        contadorVelocidad = 0
        velocidadCataratas = velocidadCataratas + aVeloCata
        If velocidadCataratas >= 3 Then
            aVeloCata = -1
        ElseIf velocidadCataratas <= 1 Then
            aVeloCata = 1
            subeobaja = subeobaja + 1
            For i = 1 To Habitaciones(BrucePlayer.Habitacion).NumeroCataratas
                If subeobaja = 1 Then
                    Habitaciones(BrucePlayer.Habitacion).Cataratas(i).Direccion = HaciaAbajo
                Else
                    Habitaciones(BrucePlayer.Habitacion).Cataratas(i).Direccion = HaciaArriba
                End If
            Next i
            If subeobaja = 1 Then
                subeobaja = -1
            End If
        End If
    End If
End Sub

Private Sub tmrFuegos_Timer()
    DelayFuego = True
End Sub
Private Sub tmrLasers_Timer()
    DelayLasers = True
End Sub
Private Sub tmrRayos_Timer()
    DelayRayos = True
End Sub

Private Sub tmrGameOver_Timer()
  Static contador As Integer
    contador = contador + 1
    If tmrGameOver.Interval <> 500 Then
        tmrGameOver.Interval = 500
    End If
    If contador = 1 Then
        If ColorMode = False Then
            ImgGameOver.Picture = LoadPicture(App.Path & "\Pants\gameover.gif")
        Else
            ImgGameOver.Picture = LoadPicture(App.Path & "\Pants\gameoverC.gif")
        End If
    Else
        If ColorMode = False Then
            ImgGameOver.Picture = LoadPicture(App.Path & "\Pants\gameover2.gif")
        Else
            ImgGameOver.Picture = LoadPicture(App.Path & "\Pants\gameover2C.gif")
        End If
        contador = 0
    End If
End Sub

Private Sub tmrLogo_Timer()
  Static contador As Integer
    If tmrLogo.Interval <> 3000 Then
        tmrLogo.Interval = 3000
    End If
    contador = contador + 1
    PicMainASCII.Visible = False
    If contador = 1 Then
        If ColorMode = False Then
            PicMain.Picture = LoadPicture(App.Path & "\pants\logo.gif")
        Else
            PicMain.Picture = LoadPicture(App.Path & "\pants\logoc.gif")
        End If
    ElseIf contador = 2 Then
        If ColorMode = False Then
            PicMain.Picture = LoadPicture(App.Path & "\pants\bruce1.gif")
            Me.BackColor = &H7B00
        Else
            PicMain.Picture = LoadPicture(App.Path & "\pants\bruce1c.gif")
            Me.BackColor = RGB(99, 99, 99)
        End If
    Else
        PantallaPrincipal
    End If
End Sub

Private Sub PantallaPrincipal()
On Error Resume Next
    tmrLogo.Enabled = False
    Image1.Visible = False
    ImgGameOver.Visible = False
    
    TotalLamparasRecogidas = 0
    Set BrucePlayer = Nothing
    Set BrucePlayer = New cBruceLee
    
    BrucePlayer.PosX = 61
    BrucePlayer.PosY = 71
    BrucePlayer.Accion = Quieto
    BrucePlayer.Habitacion = 1
    
    'IniciarHabitaciones    'Uses Split Funcion (only for VB6.0)
    IniciarHabitacionesVB5  '<----Now Compatible with VB4.0 and VB5.0 & VB6.0
    
    VisitasVidas = 5
    aVeloCata = 1
    
    velocidadCataratas = aVeloCata
    
    If ColorMode = False Then
        PicClean.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapPic)
        PicBuff.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapMask)
        Picture1.Picture = LoadPicture(App.Path & "\pants\sprites.gif")
        Me.BackColor = &H7B00
        PicMain.Picture = LoadPicture(App.Path & "\pants\bruce2.gif")
    Else
        PicClean.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapPic)
        PicBuff.Picture = LoadPicture(App.Path & Habitaciones(BrucePlayer.Habitacion).MapMask)
        Picture1.Picture = LoadPicture(App.Path & "\pants\spritesc.gif")
        Me.BackColor = RGB(99, 99, 99)
        PicMain.Picture = LoadPicture(App.Path & "\pants\bruce2c.gif")
    End If
    EnPantallaPrincipal = True
    PicMainASCII.Visible = False
    

End Sub

Private Sub tmrMuere_Timer()
  Static contador As Integer
  
    contador = contador + 1
    If contador >= 8 Then
        pUp = False
        pDown = False
        pRight = False
        pLeft = False
        pSpaceBar = False
        contador = 0
        StatMuere = 0
        tmrMuere.Enabled = False
        If BrucePlayer.GameOver = False Then
            BrucePlayer.ReposicionaMuerteBruce
        End If
    End If
End Sub

Private Sub tmrPuente_Timer()
  Static QuePuente As Integer
    
    If tmrPuente.Interval < 1000 Then
        tmrPuente.Interval = 1000
    End If
    QuePuente = QuePuente + 1
    If QuePuente = 1 Then
    Else
        QuePuente = 0
    End If
    LQuePuente = QuePuente
End Sub

Private Sub PintaBaca(QueBaca As Integer)
    If tmrBaca.Enabled = False Then
        Exit Sub
    End If
    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC
    Select Case BrucePlayer.Habitacion
        Case 1
            If QueBaca = 1 Then
                BitBlt DDC, 58, 120, 10, 6, SDC, 183, 57, SRCAND
                BitBlt DDC, 58, 120, 10, 6, SDC, 183, 50, SRCPAINT
            End If
        Case 3
            If QueBaca = 1 Then
                BitBlt DDC, 266, 120, 10, 6, SDC, 183, 57, SRCAND
                BitBlt DDC, 266, 120, 10, 6, SDC, 183, 50, SRCPAINT
            End If
    End Select
End Sub

Private Sub PintaPuente(QuePuente As Integer)
    
    If BrucePlayer.Habitacion <> 10 Then
        Exit Sub
    End If
    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC
    
    If QuePuente = 1 Then
        BitBlt DDC, 72, 74, 56, 4, SDC, 233, 31, SRCAND
        BitBlt DDC, 72, 74, 56, 4, SDC, 233, 26, SRCPAINT
        ModificarMascaraJIT 72, 75, 128, 76, 0
        ModificarMascaraJIT 168, 75, 224, 76
    Else
        QuePuente = 0
        BitBlt DDC, 168, 74, 56, 4, SDC, 233, 31, SRCAND
        BitBlt DDC, 168, 74, 56, 4, SDC, 233, 26, SRCPAINT
        ModificarMascaraJIT 72, 75, 128, 76
        ModificarMascaraJIT 168, 75, 224, 76, 0
    End If
End Sub
Sub Initialise()
On Error GoTo ErrorHandler
    Set Ds = Dx.DirectSoundCreate("")
    Ds.SetCooperativeLevel frmMain.hWnd, DSSCL_NORMAL


    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DsWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
    DsWave.nChannels = 2    '1= Mono, 2 = Stereo
    DsWave.lSamplesPerSec = 22050
    DsWave.nBitsPerSample = 16 '16 =16bit, 8=8bit
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign

    Set DsBuffer(0) = Ds.CreateSoundBufferFromFile(App.Path & "\sonidos\run.Wav", DsDesc, DsWave)
    Set DsBuffer(1) = Ds.CreateSoundBufferFromFile(App.Path & "\sonidos\cow.Wav", DsDesc, DsWave)

Exit Sub
ErrorHandler:
    MsgBox "Unable to Continue, Error creating Directsound object."
End Sub
