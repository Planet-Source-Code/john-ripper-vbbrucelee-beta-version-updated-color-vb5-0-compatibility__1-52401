VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBruceLee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Dim perfStart As Currency 'these are for timing the animation
Dim perfEnd As Currency
Dim Elapsed As Double

Dim perfFreq As Currency

Enum eAcciones
    Quieto = 0
    Caer
    Correr
    Agacharse
    Saltar
    SaltoDiagonal
    Puñetazo
    Patada
    Escalar
    DañoDePie
    DañoTumbado
End Enum

Dim mAccion As eAcciones

Enum eDireccion
    Derecha = 0
    Izquieda
End Enum

Dim mDireccion As eDireccion

Private Type tSrprites
    TablaPosX As Integer    'Posicion X en el Picture de los sprites
    TablaPosY As Integer    'Posicion Y en el Picture de los sprites
    Ancho As Integer        'Alto del sprite en el Picture de los sprites
    Alto As Integer         'Ancho del sprite en el Picture de los sprites
    TablaMaskPosX As Integer
    TablaMaskPosY As Integer
End Type

Private Sprite As tSrprites

Dim mAnimacion As Integer
Dim FrameAnterior As Integer
Dim mX As Integer
Dim mY As Integer
Dim mHabitacion As Integer
Dim mOrigenSalto As Integer
Dim mOrigenPatada As Integer
Dim mOrigenSaltoDiagonalX As Integer
Dim mOrigenSaltoDiagonalY As Integer
Dim mCayendoDiagonal As Boolean
Dim mMuere As Boolean
Dim mPuntos As Long
Dim mVidas As Integer
Dim mGameOver As Boolean

Private Sub Actualizar_tSprites()
    With Sprite
        Select Case mDireccion
            Case eDireccion.Derecha
                Select Case mAccion
                    Case eAcciones.Agacharse
                        .Alto = 7
                        .Ancho = 28
                        .TablaPosX = 233
                        .TablaPosY = 17
                        .TablaMaskPosX = 233
                        .TablaMaskPosY = 41
                    Case eAcciones.Caer
                        .Alto = 23
                        .Ancho = 16
                        .TablaPosX = 102
                        .TablaPosY = 1
                        .TablaMaskPosX = 102
                        .TablaMaskPosY = 25
                    Case eAcciones.Correr
                        Select Case mAnimacion
                            Case 1
                                .Alto = 21
                                .Ancho = 14
                                .TablaPosX = 377
                                .TablaPosY = 3
                                .TablaMaskPosX = 377
                                .TablaMaskPosY = 27
                            Case 2
                                .Alto = 21
                                .Ancho = 14
                                .TablaPosX = 392
                                .TablaPosY = 3
                                .TablaMaskPosX = 392
                                .TablaMaskPosY = 27
                        End Select
                    Case eAcciones.DañoDePie
                        .Alto = 21
                        .Ancho = 15
                        .TablaPosX = 291
                        .TablaPosY = 3
                        .TablaMaskPosX = 291
                        .TablaMaskPosY = 27

                    Case eAcciones.DañoTumbado
                        .Alto = 13
                        .Ancho = 26
                        .TablaPosX = 323
                        .TablaPosY = 11
                        .TablaMaskPosX = 323
                        .TablaMaskPosY = 35
                    
                    Case eAcciones.Escalar
                        .Alto = 22
                        .Ancho = 16
                        .TablaPosX = 199
                        .TablaPosY = 2
                        .TablaMaskPosX = 199
                        .TablaMaskPosY = 26
                    
                    Case eAcciones.Patada
                        .Alto = 17
                        .Ancho = 24
                        .TablaPosX = 119
                        .TablaPosY = 7
                        .TablaMaskPosX = 119
                        .TablaMaskPosY = 31

                    Case eAcciones.Puñetazo
                        .Alto = 20
                        .Ancho = 24
                        .TablaPosX = 35
                        .TablaPosY = 4
                        .TablaMaskPosX = 35
                        .TablaMaskPosY = 28
                    Case eAcciones.Quieto
                        .Alto = 23
                        .Ancho = 16
                        .TablaPosX = 1
                        .TablaPosY = 1
                        .TablaMaskPosX = 1
                        .TablaMaskPosY = 25
                    
                    Case eAcciones.Saltar
                        Select Case mAnimacion
                            Case 1
                                .Alto = 19
                                .Ancho = 16
                                .TablaPosX = 85
                                .TablaPosY = 5
                                .TablaMaskPosX = 85
                                .TablaMaskPosY = 29
                            Case 2
                                .Alto = 23
                                .Ancho = 16
                                .TablaPosX = 102
                                .TablaPosY = 1
                                .TablaMaskPosX = 102
                                .TablaMaskPosY = 25
                        End Select
                    Case eAcciones.SaltoDiagonal
                            .Alto = 20
                            .Ancho = 14
                            .TablaPosX = 169
                            .TablaPosY = 4
                            .TablaMaskPosX = 169
                            .TablaMaskPosY = 28
                    
                End Select
                
            Case eDireccion.Izquieda
                Select Case mAccion
                    Case eAcciones.Agacharse
                        .Alto = 7
                        .Ancho = 28
                        .TablaPosX = 262
                        .TablaPosY = 17
                        .TablaMaskPosX = 262
                        .TablaMaskPosY = 41
                    Case eAcciones.Caer
                        .Alto = 23
                        .Ancho = 16
                        .TablaPosX = 102
                        .TablaPosY = 1
                        .TablaMaskPosX = 102
                        .TablaMaskPosY = 25
                    Case eAcciones.Correr
                        Select Case mAnimacion
                            Case 1
                                .Alto = 21
                                .Ancho = 14
                                .TablaPosX = 407
                                .TablaPosY = 3
                                .TablaMaskPosX = 407
                                .TablaMaskPosY = 27
                            Case 2
                                .Alto = 21
                                .Ancho = 10
                                .TablaPosX = 422
                                .TablaPosY = 3
                                .TablaMaskPosX = 422
                                .TablaMaskPosY = 27
                        End Select
                    Case eAcciones.DañoDePie
                        .Alto = 21
                        .Ancho = 15
                        .TablaPosX = 307
                        .TablaPosY = 3
                        .TablaMaskPosX = 307
                        .TablaMaskPosY = 27

                    Case eAcciones.DañoTumbado
                        .Alto = 13
                        .Ancho = 26
                        .TablaPosX = 350
                        .TablaPosY = 11
                        .TablaMaskPosX = 350
                        .TablaMaskPosY = 35
                    
                    Case eAcciones.Escalar
                        .Alto = 22
                        .Ancho = 16
                        .TablaPosX = 216
                        .TablaPosY = 2
                        .TablaMaskPosX = 216
                        .TablaMaskPosY = 26
                    
                    Case eAcciones.Patada
                        .Alto = 17
                        .Ancho = 24
                        .TablaPosX = 144
                        .TablaPosY = 7
                        .TablaMaskPosX = 144
                        .TablaMaskPosY = 31

                    Case eAcciones.Puñetazo
                        .Alto = 20
                        .Ancho = 24
                        .TablaPosX = 60
                        .TablaPosY = 4
                        .TablaMaskPosX = 60
                        .TablaMaskPosY = 28
                    Case eAcciones.Quieto
                        .Alto = 23
                        .Ancho = 16
                        .TablaPosX = 18
                        .TablaPosY = 1
                        .TablaMaskPosX = 18
                        .TablaMaskPosY = 25
                    
                    Case eAcciones.Saltar
                        Select Case mAnimacion
                            Case 1
                                .Alto = 19
                                .Ancho = 16
                                .TablaPosX = 85
                                .TablaPosY = 5
                                .TablaMaskPosX = 85
                                .TablaMaskPosY = 29
                            Case 2
                                .Alto = 23
                                .Ancho = 16
                                .TablaPosX = 102
                                .TablaPosY = 1
                                .TablaMaskPosX = 102
                                .TablaMaskPosY = 25
                        End Select
                    Case eAcciones.SaltoDiagonal
                            .Alto = 20
                            .Ancho = 14
                            .TablaPosX = 184
                            .TablaPosY = 4
                            .TablaMaskPosX = 184
                            .TablaMaskPosY = 28
                    
                End Select
        
        End Select
    End With
End Sub

Private Function NumeroAnimaciones() As Integer
    Select Case mAccion
        Case eAcciones.Correr, eAcciones.Saltar
            NumeroAnimaciones = 2
        Case Else
            NumeroAnimaciones = 1
    End Select
End Function

Private Sub CambiarFrame()
  Dim TotalAnimaciones As Integer
  Dim TempFrame As Integer
  Static AnteriorEscalar As Integer
  Static test As Single

    test = test + (Elapsed)
    If test < 75 Then
        Actualizar_tSprites
        Exit Sub
        
    Else
        test = 0
        
    End If
    TotalAnimaciones = NumeroAnimaciones
    If mAccion = Escalar Then
       If pUp = True Or pDown = True Or pLeft = True Or pRight = True Then

            If AnteriorEscalar = 0 Then
                mDireccion = Derecha
                AnteriorEscalar = 1
            Else
                mDireccion = Izquieda
                AnteriorEscalar = 0
            End If
        End If
    End If
    If TotalAnimaciones = 1 Then
        mAnimacion = 1
    Else
        TempFrame = FrameAnterior + 1
        If TempFrame <= TotalAnimaciones Then
            mAnimacion = TempFrame
        Else
            If mAccion = Saltar Then
                mAnimacion = 2
                TempFrame = 2
            Else
                mAnimacion = 1
                TempFrame = 1
            End If
        End If
    
    End If
    FrameAnterior = TempFrame
    Actualizar_tSprites
End Sub

Public Property Get GameOver() As Boolean
    GameOver = mGameOver
End Property

Public Property Let Vidas(ByVal pVidas As Long)
    mVidas = pVidas
End Property

Public Property Get Vidas() As Long
    Vidas = mVidas
End Property

Public Property Let Puntos(ByVal pPuntos As Long)
    mPuntos = pPuntos
End Property

Public Property Get Puntos() As Long
    Puntos = mPuntos
End Property

Public Property Get Muere() As Boolean
    Muere = mMuere
End Property

Public Property Let Muere(ByVal pMuere As Boolean)
    mMuere = pMuere
End Property

Public Property Get Habitacion() As Integer
    Habitacion = mHabitacion
End Property

Public Property Let Habitacion(ByVal pHabitacion As Integer)
    mHabitacion = pHabitacion
End Property

Public Property Get CayendoDiagonal() As Boolean
    CayendoDiagonal = mCayendoDiagonal
End Property

Public Property Let CayendoDiagonal(ByVal pCayendoDiagonal As Boolean)
    mCayendoDiagonal = pCayendoDiagonal
End Property

Public Property Get OrigenSaltoDiagonalX() As Integer
    OrigenSaltoDiagonalX = mOrigenSaltoDiagonalX
End Property

Public Property Let OrigenSaltoDiagonalX(ByVal pOrigenSaltoDiagonalX As Integer)
    mOrigenSaltoDiagonalX = pOrigenSaltoDiagonalX
End Property

Public Property Get OrigenSaltoDiagonalY() As Integer
    OrigenSaltoDiagonalY = mOrigenSaltoDiagonalY
End Property

Public Property Let OrigenSaltoDiagonalY(ByVal pOrigenSaltoDiagonalY As Integer)
    mOrigenSaltoDiagonalY = pOrigenSaltoDiagonalY
End Property

Public Property Get OrigenSalto() As Integer
    OrigenSalto = mOrigenSalto
End Property

Public Property Let OrigenSalto(ByVal pOrigenSalto As Integer)
    mOrigenSalto = pOrigenSalto
End Property

Public Property Get OrigenPatada() As Integer
    OrigenPatada = mOrigenPatada
End Property

Public Property Let OrigenPatada(ByVal pOrigenPatada As Integer)
    mOrigenPatada = pOrigenPatada
End Property

Public Property Get PosX() As Integer
    PosX = mX
End Property

Public Property Let PosX(ByVal pPosicion As Integer)
    mX = pPosicion
End Property

Public Property Get PosY() As Integer
    PosY = mY
End Property

Public Property Let PosY(ByVal pPosicion As Integer)
    mY = pPosicion
End Property

Public Property Get Accion() As eAcciones
    Accion = mAccion
End Property

Public Property Let Accion(ByVal pAccion As eAcciones)
    mAccion = pAccion
    FrameAnterior = 0
End Property

Public Property Get Direccion() As eDireccion
    Direccion = mDireccion
End Property

Public Property Let Direccion(ByVal pDireccion As eDireccion)
    mDireccion = pDireccion
End Property

Private Sub Class_Initialize()
    FrameAnterior = 0
    mDireccion = Derecha
    mAccion = Quieto
    mX = 30
    mY = 175
    mHabitacion = 1
    mPuntos = 0
    mVidas = 4
    QueryPerformanceFrequency perfFreq
End Sub

Private Function CheckCollision(Rect1 As RECT, Rect2 As RECT) As Boolean
    Dim emptyrect As RECT
    
    If IntersectRect(emptyrect, Rect1, Rect2) <> 0 Then
        CheckCollision = True
    End If
End Function

Private Function BruceTocoVidas(queHabitacion) As Boolean
  Dim i As Integer
  Dim j As Integer
  Dim BruceRECT As RECT
  Dim LampRECT As RECT

  BruceTocoVidas = False
    
    With BruceRECT
        .Left = mX
        .Right = mX + Sprite.Ancho
        If mAccion = Patada Then
            .Top = mY - Sprite.Alto - 7
            .Bottom = mY - 7
        Else
            .Top = mY - Sprite.Alto
            .Bottom = mY
        End If
    End With
    
    With LampRECT
        .Left = 288
        .Top = 128
        .Right = .Left + 16
        .Bottom = .Top + 32
        
    End With
    If CheckCollision(BruceRECT, LampRECT) = True Then
        BruceTocoVidas = True
        Exit Function
    End If

End Function

Private Function BruceTocoLampara(queHabitacion) As Boolean
  Dim i As Integer
  Dim j As Integer
  Dim BruceRECT As RECT
  Dim LampRECT As RECT

    BruceTocoLampara = False
    
    With BruceRECT
        .Left = mX
        .Right = mX + Sprite.Ancho
        If mAccion = Patada Then
            .Top = mY - Sprite.Alto - 7
            .Bottom = mY - 7
        Else
            .Top = mY - Sprite.Alto
            .Bottom = mY
        End If
        
    End With
    
    For i = 1 To Habitaciones(mHabitacion).NumeroLamparas
        If Habitaciones(mHabitacion).Lamparas(i).Estado = True Then
            With LampRECT
                .Left = Habitaciones(mHabitacion).Lamparas(i).PosX
                .Top = Habitaciones(mHabitacion).Lamparas(i).PosY
                Select Case Habitaciones(mHabitacion).Lamparas(i).Tipo
                    Case 0, 1
                        .Right = .Left + 8
                        .Bottom = .Top + 10
                    Case 2
                        .Right = .Left + 6
                        .Bottom = .Top + 12
                End Select
            End With

            If CheckCollision(BruceRECT, LampRECT) = True Then
                BruceTocoLampara = True
                Habitaciones(mHabitacion).Lamparas(i).Estado = False
                Exit Function
            End If
        End If
    Next i
End Function

Public Function CheckCambioHabitacion() As Boolean
  Dim i As Integer
  Dim j As Integer

  Dim BruceRECT As RECT
  Dim SalidaRECT As RECT

    With BruceRECT
        .Left = mX
        .Right = mX + Sprite.Ancho
        If mAccion = Patada Then
            .Top = mY - Sprite.Alto - 7
            .Bottom = mY - 7
        Else
            .Top = mY - Sprite.Alto
            .Bottom = mY
        End If
    End With

    For j = 1 To Habitaciones(mHabitacion).NumeroSalidas
        With SalidaRECT
            .Left = Habitaciones(mHabitacion).Salidas(j).OrgX1
            .Top = Habitaciones(mHabitacion).Salidas(j).OrgY1
            .Right = Habitaciones(mHabitacion).Salidas(j).OrgX2
            .Bottom = Habitaciones(mHabitacion).Salidas(j).OrgY2
        End With
        
        If CheckCollision(BruceRECT, SalidaRECT) = True Then
            mX = Habitaciones(mHabitacion).Salidas(j).DestinoX
            mY = Habitaciones(mHabitacion).Salidas(j).DestinoY
            mHabitacion = Habitaciones(mHabitacion).Salidas(j).HabitacionDestino
            CheckCambioHabitacion = True
            Exit Function
        End If
    Next j
End Function

Public Sub ChequeaCogerVidas()
  Dim toco As Boolean

    If mHabitacion = 15 Then
        If HaCogidoVida = False Then
            toco = BruceTocoVidas(15)
            If toco = True Then
                HaCogidoVida = True
                If VisitasVidas > 0 Then
                    mVidas = mVidas + 1
                End If
            End If
        End If
    End If
End Sub

Public Sub ChequeaLamparas()
  Dim toco As Boolean
  Dim retval As Long
    
    toco = BruceTocoLampara(mHabitacion)
    If toco = True Then
        TotalLamparasRecogidas = TotalLamparasRecogidas + 1
        mPuntos = mPuntos + 125
        If CanPlayWave = True Then
            retval = PlaySound(App.Path & "\sonidos\collect.wav", 0, SND_FILENAME Or SND_ASYNC)
        End If
    End If
End Sub

Public Sub ChequeaRayos()
  Dim i As Integer
  Dim j As Integer

  Dim BruceRECT As RECT
  Dim RayosRECT As RECT

    If Habitaciones(mHabitacion).NumeroRayos = 0 Then
        Exit Sub
    End If
    
    With BruceRECT
        .Left = mX
        .Right = mX + Sprite.Ancho
        If mAccion = Patada Then
            .Top = mY - Sprite.Alto - 7
            .Bottom = mY - 7
        Else
            .Top = mY - Sprite.Alto
            .Bottom = mY
        End If
    End With

    For j = 1 To Habitaciones(mHabitacion).NumeroRayos
        With RayosRECT
            .Left = Habitaciones(mHabitacion).Rayos(j).tempX
            .Top = Habitaciones(mHabitacion).Rayos(j).PosY1
            .Right = .Left + 8
            .Bottom = .Top + 10
        End With
        
        If Habitaciones(mHabitacion).Rayos(j).Visible = True Then
            If CheckCollision(BruceRECT, RayosRECT) = True Then
                If mMuere = False Then
                    mMuere = True
                    Habitaciones(mHabitacion).Rayos(j).Stop = True
                End If
                Exit Sub
            End If
        End If
        
    Next j
End Sub

Public Sub ChequeaLasers()
  Dim i As Integer
  Dim j As Integer

  Dim BruceRECT As RECT
  Dim LaserRECT As RECT

    If Habitaciones(mHabitacion).NumeroLasers = 0 Then
        Exit Sub
    End If
    
    With BruceRECT
        .Left = mX
        .Right = mX + Sprite.Ancho
        If mAccion = Patada Then
            .Top = mY - Sprite.Alto - 7
            .Bottom = mY - 7
        Else
            .Top = mY - Sprite.Alto
            .Bottom = mY
        End If
    End With

    For j = 1 To Habitaciones(mHabitacion).NumeroLasers
        With LaserRECT
            .Left = Habitaciones(mHabitacion).Laser(j).tempX
            .Top = Habitaciones(mHabitacion).Laser(j).PosY1 - 1
            .Right = .Left + 8
            .Bottom = .Top + 1
        End With
        
        If CheckCollision(BruceRECT, LaserRECT) = True Then
            If mMuere = False Then
                mMuere = True
                Habitaciones(mHabitacion).Laser(j).Stop = True
            End If
            Exit Sub
        End If
    Next j
End Sub

Public Sub ChequeaFuego()
  Dim i As Integer
  Dim j As Integer

  Dim BruceRECT As RECT
  Dim FuegoRECT As RECT

    If Habitaciones(mHabitacion).NumeroFuegos = 0 Then
        Exit Sub
    End If
    
    With BruceRECT
        .Left = mX
        .Right = mX + Sprite.Ancho
        If mAccion = Patada Then
            .Top = mY - Sprite.Alto - 7
            .Bottom = mY - 7
        Else
            .Top = mY - Sprite.Alto
            .Bottom = mY
        End If
    End With

    For j = 1 To Habitaciones(mHabitacion).NumeroFuegos
        With FuegoRECT
            If Habitaciones(mHabitacion).Fuegos(j).Activo = False Then
                .Left = Habitaciones(mHabitacion).Fuegos(j).PosX
                .Top = Habitaciones(mHabitacion).Fuegos(j).PosY - 2
                .Right = Habitaciones(mHabitacion).Fuegos(j).PosX + 8
                .Bottom = Habitaciones(mHabitacion).Fuegos(j).PosY + 1
            Else
                .Left = Habitaciones(mHabitacion).Fuegos(j).PosX
                
                Select Case Habitaciones(mHabitacion).Fuegos(j).Tipo
                    Case 1
                        .Top = Habitaciones(mHabitacion).Fuegos(j).PosY - 8
                    Case 2
                        .Top = Habitaciones(mHabitacion).Fuegos(j).PosY - 18
                End Select
                .Right = Habitaciones(mHabitacion).Fuegos(j).PosX + 8
                .Bottom = Habitaciones(mHabitacion).Fuegos(j).PosY
            End If
        End With
        
        If CheckCollision(BruceRECT, FuegoRECT) = True Then
            If Habitaciones(mHabitacion).Fuegos(j).Activo = False Then
                Habitaciones(mHabitacion).Fuegos(j).Activo = True
            Else
                If Habitaciones(mHabitacion).Fuegos(j).Espera >= 3 Then
                    If mMuere = False Then
                        mMuere = True
                    End If
                End If
            End If
            Exit Sub
        End If

        
    Next j
End Sub

Public Sub ReposicionaMuerteBruce()
  Dim i As Integer
    
    If mGameOver = False Then
        mVidas = mVidas - 1
    End If
    If mVidas < 0 Then
        mGameOver = True
        If frmMain.ImgGameOver.Visible = False Then
            frmMain.ImgGameOver.Visible = True
        End If
    End If
    
    For i = 1 To Habitaciones(mHabitacion).NumeroLasers
        Habitaciones(mHabitacion).Laser(i).Stop = False
    Next i
    
    For i = 1 To Habitaciones(mHabitacion).NumeroRayos
        Habitaciones(mHabitacion).Rayos(i).Stop = False
    Next i

    Select Case mHabitacion
        Case 4
            If mX < 270 Then
                mX = Habitaciones(2).Salidas(4).DestinoX
                mY = Habitaciones(2).Salidas(4).DestinoY
            Else
                mX = Habitaciones(9).Salidas(1).DestinoX
                mY = Habitaciones(9).Salidas(1).DestinoY
            End If
        
        Case 5
            mX = Habitaciones(4).Salidas(1).DestinoX
            mY = 155
        
        Case 6
            If mY < 100 Then
                mX = Habitaciones(5).Salidas(2).DestinoX
                mY = Habitaciones(5).Salidas(2).DestinoY
            Else
                mX = Habitaciones(7).Salidas(2).DestinoX
                mY = Habitaciones(7).Salidas(2).DestinoY
            End If
        
        Case 7
            mX = Habitaciones(6).Salidas(1).DestinoX
            mY = Habitaciones(6).Salidas(1).DestinoY
        
        Case 8
            If mY > 160 Then
                mX = Habitaciones(7).Salidas(3).DestinoX
                mY = Habitaciones(7).Salidas(3).DestinoY
            Else
                mX = Habitaciones(7).Salidas(4).DestinoX
                mY = Habitaciones(7).Salidas(4).DestinoY
            End If
        
        Case 9
            mX = 263
            mY = 154
        
        Case 10
            mX = Habitaciones(3).Salidas(2).DestinoX
            mY = Habitaciones(3).Salidas(2).DestinoY
        
        Case 11
            mX = Habitaciones(10).Salidas(2).DestinoX
            mY = Habitaciones(10).Salidas(2).DestinoY
        
        Case 12
            mX = Habitaciones(11).Salidas(2).DestinoX
            mY = Habitaciones(11).Salidas(2).DestinoY
        
        Case 13
            mX = Habitaciones(12).Salidas(2).DestinoX
            mY = Habitaciones(12).Salidas(2).DestinoY
        
        Case 14
            If mY < 115 Then
                mX = Habitaciones(12).Salidas(3).DestinoX
                mY = Habitaciones(12).Salidas(3).DestinoY
            Else
                mX = Habitaciones(13).Salidas(1).DestinoX
                mY = Habitaciones(13).Salidas(1).DestinoY
            
            End If
        
        Case 15
            If mY < 115 Then
                mX = Habitaciones(12).Salidas(4).DestinoX
                mY = Habitaciones(12).Salidas(4).DestinoY
            Else
                mX = Habitaciones(14).Salidas(2).DestinoX
                mY = Habitaciones(14).Salidas(2).DestinoY
            
            End If
        
        Case 16
            mX = Habitaciones(13).Salidas(3).DestinoX
            mY = Habitaciones(13).Salidas(3).DestinoY
        
        Case 17
            mX = Habitaciones(16).Salidas(1).DestinoX
            mY = Habitaciones(16).Salidas(1).DestinoY
        
        Case 18
            mX = Habitaciones(17).Salidas(2).DestinoX
            mY = Habitaciones(17).Salidas(2).DestinoY
    End Select
    
    If mGameOver = True Then
        mMuere = True
    Else
        mMuere = False
    End If
    
    mAccion = Quieto
    mCayendoDiagonal = False
End Sub

Public Sub Render()
    
    SDC = frmMain.Picture1.hDC
    DDC = frmMain.PicWork.hDC

    Timer
    
    If mMuere = False Then
        CambiarFrame
    End If
    
    With Sprite
        If mAccion = Patada Then
            BitBlt DDC, mX, mY - .Alto - 7, .Ancho, .Alto, SDC, .TablaMaskPosX, .TablaMaskPosY, SRCAND
            BitBlt DDC, mX, mY - .Alto - 7, .Ancho, .Alto, SDC, .TablaPosX, .TablaPosY, SRCPAINT
        Else
            If mAccion = Puñetazo And mDireccion = Izquieda Then
                BitBlt DDC, mX - 8, mY - .Alto, .Ancho, .Alto, SDC, .TablaMaskPosX, .TablaMaskPosY, SRCAND
                BitBlt DDC, mX - 8, mY - .Alto, .Ancho, .Alto, SDC, .TablaPosX, .TablaPosY, SRCPAINT
            Else
                BitBlt DDC, mX, mY - .Alto, .Ancho, .Alto, SDC, .TablaMaskPosX, .TablaMaskPosY, SRCAND
                BitBlt DDC, mX, mY - .Alto, .Ancho, .Alto, SDC, .TablaPosX, .TablaPosY, SRCPAINT
            End If
        End If
    End With
       
End Sub

Private Sub Timer()
    ''This subroutine will perform all of our timing
    ''The results are used for animation and movement
    
    'Get current time
    QueryPerformanceCounter perfStart
    
    'Calculate elapsed time
    Elapsed = (perfStart - perfEnd) / perfFreq * 1000
    
    'Now, we start the timer again, this will be used to calculate
    'how much time has passed since we last called this subroutine.
    QueryPerformanceCounter perfEnd
    
    
End Sub

Public Sub ChequeaPinchos()
On Error GoTo ErrorHandler
  Dim i As Integer
  Dim j As Integer
  Dim docheck As Boolean
    
    For j = 1 To Sprite.Ancho
    For i = -1 To Sprite.Alto
        
        If (mX + j > 0 And mX + j < bmpBuff.bmWidth - 1) And (mY + i > 0 And (bmpBuff.bmHeight - mY + i) < bmpBuff.bmHeight) Then
            If pictBuff(mX + j, bmpBuff.bmHeight - mY + i) = 3 Then
                mMuere = True
                Exit For
            End If
        End If

    Next i
    Next j
    Exit Sub
ErrorHandler:
    'Debug.Print "XXX"

End Sub

Public Function ImposibleMoverIzquierda() As Boolean
On Error Resume Next
  Dim i As Integer
  Dim j As Integer
    
    ImposibleMoverIzquierda = False
    For j = 1 To 3
        For i = 1 To Sprite.Alto
            If pictBuff(mX - j, bmpBuff.bmHeight - mY + i) = 0 Then
                ImposibleMoverIzquierda = True
            End If
            Exit Function
        Next i
    Next j
End Function

Public Function ImposibleMoverDerecha() As Boolean
On Error Resume Next
  Dim i As Integer
    ImposibleMoverDerecha = False
    For i = 1 To 3
        If pictBuff(mX + Sprite.Ancho + i, bmpBuff.bmHeight - mY) = 0 Then
            ImposibleMoverDerecha = True
            Exit For
        End If
    Next i
End Function

Public Sub SubeUnaLinea()
On Error Resume Next
Dim i As Integer
    For i = 1 To Sprite.Ancho
        If pictBuff(mX + i, bmpBuff.bmHeight - mY - 1) = 0 Then
            mY = mY - 1
            Exit For
        End If
    Next i

End Sub

Public Function EstaCayendo() As Boolean
On Error Resume Next
Dim i As Integer
    EstaCayendo = True
    For i = 1 To Sprite.Ancho
        If pictBuff(mX + i, bmpBuff.bmHeight - mY - 1) <> 4 Then
            EstaCayendo = False
            Exit For
        End If
    Next i
End Function

Public Function EstaSobreCatarata() As Boolean
On Error GoTo ErrorHandler 'Resume Next
  Dim i As Integer
  Dim j As Integer
    
    EstaSobreCatarata = False
    For j = 1 To Sprite.Ancho
        For i = -1 To Sprite.Alto
            If pictBuff(mX + j, bmpBuff.bmHeight - mY + i) = 2 Then
                EstaSobreCatarata = True
                Exit For
            End If
        Next i
    Next j
ErrorHandler:
End Function

Public Function EstaEscalando() As Boolean
On Error Resume Next
Dim i As Integer
    EstaEscalando = False
    For i = 1 To Sprite.Ancho
        If pictBuff(mX + i, bmpBuff.bmHeight - mY + Sprite.Alto) = 1 Or pictBuff(mX + i, bmpBuff.bmHeight - mY + Sprite.Alto) = 2 Then
            EstaEscalando = True
            Exit For
        End If
    
        If pictBuff(mX + i, bmpBuff.bmHeight + (Sprite.Alto \ 2) - mY + -1) = 1 Or pictBuff(mX + i, bmpBuff.bmHeight + (Sprite.Alto \ 2) - mY + -1) = 2 Then
            EstaEscalando = True
            Exit For
        End If
        
        If pictBuff(mX + i, bmpBuff.bmHeight - mY + -1) = 1 Or pictBuff(mX + i, bmpBuff.bmHeight - mY + -1) = 2 Then
            EstaEscalando = True
            Exit For
        End If
    Next i
End Function
