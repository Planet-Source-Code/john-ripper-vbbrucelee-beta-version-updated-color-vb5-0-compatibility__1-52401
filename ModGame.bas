Attribute VB_Name = "ModGame"
Option Explicit
Dim sSplitDataVB5() As String
Public ColorMode As Boolean
Public TopScore As Long
Public TotalLamparasRecogidas As Integer
Public SDC As Long
Public DDC As Long
Public QuedanVidas As Boolean
Public HaCogidoVida As Boolean
Public VisitasVidas As Integer

Public Const AUDIO_NONE = 0
Public Const AUDIO_WAVE = 1
Public Const AUDIO_MIDI = 2
Public Const AUDIO_BOTH = 3
Public CanPlayMidi       As Boolean
Public CanPlayWave       As Boolean

Public pUp As Boolean
Public pDown As Boolean
Public pLeft As Boolean
Public pRight As Boolean
Public pSpaceBar As Boolean

Public Const TotalHabitaciones = 20

Private Type tSalidas
    OrgX1 As Integer
    OrgY1 As Integer
    OrgX2 As Integer
    OrgY2 As Integer
    DestinoX As Integer
    DestinoY As Integer
    HabitacionDestino As Integer
End Type

Private Type tLamparas
    Estado As Boolean
    PosX As Integer
    PosY As Integer
    Tipo As Integer
End Type

Private Type tRayos
    PosX1 As Integer
    PosY1 As Integer
    PosX2 As Integer
    InvisX1 As Integer
    InvisX2 As Integer
    tempX As Integer
    TempY As Integer
    Velo As Integer
    Visible As Boolean
    Stop As Boolean
End Type

Public Enum eDireccionLaser
    deIzquierdaADerecha = 0
    deDerechaAIzquierda = 1
End Enum

Private Type tLaser
    PosX1 As Integer
    tempX As Integer
    PosY1 As Integer
    Distancia As Integer
    Direccion As eDireccionLaser
    Stop As Boolean
End Type

Public Type tFuego
    PosX As Integer
    PosY As Integer
    Tipo As Integer
    Activo As Boolean
    Frame As Integer
    Espera As Integer
End Type

Public Enum eDireccionCataratas
    HaciaArriba = 0
    HaciaAbajo
End Enum

Public Type tCataratas
    PosX As Integer
    PosY As Integer
    PosY2 As Integer
    TempY As Integer
    Tipo As Integer
    Velocidad As Integer
    Direccion As eDireccionCataratas
End Type

Private Type tHabitaciones
    NumeroLamparas As Integer
    NumeroSalidas As Integer
    NumeroRayos As Integer
    NumeroLasers As Integer
    NumeroFuegos As Integer
    NumeroCataratas As Integer
    MapPic As String
    MapMask As String
    Lamparas(13) As tLamparas
    Salidas(4) As tSalidas
    Rayos(4) As tRayos
    Laser(7) As tLaser
    Fuegos(6) As tFuego
    Cataratas(10) As tCataratas
End Type

Public Habitaciones(TotalHabitaciones) As tHabitaciones

'Public Sub IniciarHabitaciones()
''Lee el fichero HabData.txt e inicializa la esctructura Hatitaciones()
'
'  Dim nf As Integer
'  Dim Numero As Integer
'  Dim HabActual As Integer
'  Dim sSplitData() As String
'  Dim sData As String
'  Dim i As Integer
'
'    nf = FreeFile
'
'    Open App.Path & "\pants\HabData.txt" For Input As #nf
'
'    Line Input #nf, sData   ';formato salidas
'    Line Input #nf, sData   ';formato lamparas
'    Line Input #nf, sData   ';formato rayos
'    Line Input #nf, sData   ';formato Laser
'    Line Input #nf, sData   ';formato Fuego
'    Line Input #nf, sData   ';formato Cataratas
'    Line Input #nf, sData   'vbcrlf
'    Line Input #nf, sData   'TotalHabitaciones
'    Line Input #nf, sData   'vbcrlf
'
'    Do Until EOF(nf)
'        Line Input #nf, sData   '#######
'        If sData = "FIN" Then Exit Do
'
'        Line Input #nf, sData   ';Habitacion=x
'        sSplitData = Split(sData, "=")
'
'        Numero = sSplitData(1)
'        HabActual = Numero
'
'        If ColorMode = False Then
'            Habitaciones(Numero).MapMask = "\pants\p" & Numero & "_mask.gif"
'            Habitaciones(Numero).MapPic = "\pants\p" & Numero & "_320x176.gif"
'        Else
'            Habitaciones(Numero).MapMask = "\pants\p" & Numero & "_mask.gif"
'            Habitaciones(Numero).MapPic = "\pants\p" & Numero & "C_320x176.gif"
'        End If
'        Line Input #nf, sData   'Salidas=x
'        sSplitData = Split(sData, "=")
'        Numero = sSplitData(1)
'        Habitaciones(HabActual).NumeroSalidas = Numero
'
'        Line Input #nf, sData   'Lamparas=x
'        sSplitData = Split(sData, "=")
'        Numero = sSplitData(1)
'        Habitaciones(HabActual).NumeroLamparas = Numero
'
'        Line Input #nf, sData   'Rayos=x
'        sSplitData = Split(sData, "=")
'        Numero = sSplitData(1)
'        Habitaciones(HabActual).NumeroRayos = Numero
'
'        Line Input #nf, sData   'Laser=x
'        sSplitData = Split(sData, "=")
'        Numero = sSplitData(1)
'        Habitaciones(HabActual).NumeroLasers = Numero
'
'        Line Input #nf, sData   'Fuego=x
'        sSplitData = Split(sData, "=")
'        Numero = sSplitData(1)
'        Habitaciones(HabActual).NumeroFuegos = Numero
'
'        Line Input #nf, sData   'Cataratas=x
'        sSplitData = Split(sData, "=")
'        Numero = sSplitData(1)
'        Habitaciones(HabActual).NumeroCataratas = Numero
'
'
'        Line Input #nf, sData   '--- Salidas ---
'
'        For i = 1 To Habitaciones(HabActual).NumeroSalidas 'NumSal
'            Line Input #nf, sData
'            sSplitData = Split(sData, ",")
'            Habitaciones(HabActual).Salidas(i).OrgX1 = sSplitData(0)
'            Habitaciones(HabActual).Salidas(i).OrgY1 = sSplitData(1)
'            Habitaciones(HabActual).Salidas(i).OrgX2 = sSplitData(2)
'            Habitaciones(HabActual).Salidas(i).OrgY2 = sSplitData(3)
'            Habitaciones(HabActual).Salidas(i).DestinoX = sSplitData(4)
'            Habitaciones(HabActual).Salidas(i).DestinoY = sSplitData(5)
'            Habitaciones(HabActual).Salidas(i).HabitacionDestino = sSplitData(6)
'        Next i
'
'        Line Input #nf, sData   '--- Lamparas ---
'
'        For i = 1 To Habitaciones(HabActual).NumeroLamparas 'NumLam
'            Line Input #nf, sData
'            sSplitData = Split(sData, ",")
'            Habitaciones(HabActual).Lamparas(i).Estado = True
'            Habitaciones(HabActual).Lamparas(i).PosX = sSplitData(0)
'            Habitaciones(HabActual).Lamparas(i).PosY = sSplitData(1)
'            Habitaciones(HabActual).Lamparas(i).Tipo = sSplitData(2)
'        Next i
'
'        Line Input #nf, sData   '--- Rayos ---
'
'        For i = 1 To Habitaciones(HabActual).NumeroRayos 'NumRay
'            Line Input #nf, sData
'            sSplitData = Split(sData, ",")
'            Habitaciones(HabActual).Rayos(i).PosX1 = sSplitData(0)
'            Habitaciones(HabActual).Rayos(i).PosY1 = sSplitData(1)
'            Habitaciones(HabActual).Rayos(i).TempY = sSplitData(1)
'            Habitaciones(HabActual).Rayos(i).PosX2 = sSplitData(2)
'            Habitaciones(HabActual).Rayos(i).InvisX1 = sSplitData(3)
'            Habitaciones(HabActual).Rayos(i).tempX = sSplitData(3)
'            Habitaciones(HabActual).Rayos(i).InvisX2 = sSplitData(4)
'            Habitaciones(HabActual).Rayos(i).Velo = sSplitData(5)
'            Habitaciones(HabActual).Rayos(i).Visible = False
'            Habitaciones(HabActual).Rayos(i).Stop = False
'        Next i
'
'        Line Input #nf, sData   '--- Lasers ---
'
'        For i = 1 To Habitaciones(HabActual).NumeroLasers 'NumLas
'            Line Input #nf, sData
'            sSplitData = Split(sData, ",")
'            Habitaciones(HabActual).Laser(i).PosX1 = sSplitData(0)
'            Habitaciones(HabActual).Laser(i).tempX = sSplitData(0)
'            Habitaciones(HabActual).Laser(i).PosY1 = sSplitData(1)
'            Habitaciones(HabActual).Laser(i).Distancia = sSplitData(2)
'            Habitaciones(HabActual).Laser(i).Direccion = sSplitData(3)
'            Habitaciones(HabActual).Laser(i).Stop = False
'        Next i
'
'        Line Input #nf, sData   '--- Fuegos ---
'
'        For i = 1 To Habitaciones(HabActual).NumeroFuegos 'NumFue
'            Line Input #nf, sData
'            sSplitData = Split(sData, ",")
'            Habitaciones(HabActual).Fuegos(i).PosX = sSplitData(0)
'            Habitaciones(HabActual).Fuegos(i).PosY = sSplitData(1)
'            Habitaciones(HabActual).Fuegos(i).Tipo = sSplitData(2)
'            Habitaciones(HabActual).Fuegos(i).Activo = False
'            Habitaciones(HabActual).Fuegos(i).Frame = 0
'            Habitaciones(HabActual).Fuegos(i).Espera = 0
'        Next i
'
'        Line Input #nf, sData   '--- Cataratas ---
'
'        For i = 1 To Habitaciones(HabActual).NumeroCataratas 'NumCat
'            Line Input #nf, sData
'            sSplitData = Split(sData, ",")
'            Habitaciones(HabActual).Cataratas(i).PosX = sSplitData(0)
'            Habitaciones(HabActual).Cataratas(i).PosY = sSplitData(1)
'            Habitaciones(HabActual).Cataratas(i).TempY = 0
'            Habitaciones(HabActual).Cataratas(i).PosY2 = sSplitData(2)
'            Habitaciones(HabActual).Cataratas(i).Tipo = sSplitData(3)
'            Habitaciones(HabActual).Cataratas(i).Velocidad = 1
'            Habitaciones(HabActual).Cataratas(i).Direccion = HaciaArriba
'        Next i
'    Loop
'Close #nf
'End Sub

Public Sub IniciarHabitacionesVB5()
'Lee el fichero HabData.txt e inicializa la esctructura Hatitaciones()

  Dim nf As Integer
  Dim Numero As Integer
  Dim HabActual As Integer
  Dim sData As String
  Dim i As Integer
    
    nf = FreeFile
    
    Open App.Path & "\pants\HabData.txt" For Input As #nf
    
    Line Input #nf, sData   ';formato salidas
    Line Input #nf, sData   ';formato lamparas
    Line Input #nf, sData   ';formato rayos
    Line Input #nf, sData   ';formato Laser
    Line Input #nf, sData   ';formato Fuego
    Line Input #nf, sData   ';formato Cataratas
    Line Input #nf, sData   'vbcrlf
    Line Input #nf, sData   'TotalHabitaciones
    Line Input #nf, sData   'vbcrlf
    
    Do Until EOF(nf)
        Line Input #nf, sData   '#######
        If sData = "FIN" Then Exit Do
        
        Line Input #nf, sData   ';Habitacion=x
        SplitVB5 sData, "="

        Numero = sSplitDataVB5(1)
        HabActual = Numero
        
        If ColorMode = False Then
            Habitaciones(Numero).MapMask = "\pants\p" & Numero & "_mask.gif"
            Habitaciones(Numero).MapPic = "\pants\p" & Numero & "_320x176.gif"
        Else
            Habitaciones(Numero).MapMask = "\pants\p" & Numero & "_mask.gif"
            Habitaciones(Numero).MapPic = "\pants\p" & Numero & "C_320x176.gif"
        End If
        Line Input #nf, sData   'Salidas=x
        SplitVB5 sData, "="
        Numero = sSplitDataVB5(1)
        Habitaciones(HabActual).NumeroSalidas = Numero

        Line Input #nf, sData   'Lamparas=x
        SplitVB5 sData, "="
        Numero = sSplitDataVB5(1)
        Habitaciones(HabActual).NumeroLamparas = Numero

        Line Input #nf, sData   'Rayos=x
        SplitVB5 sData, "="
        Numero = sSplitDataVB5(1)
        Habitaciones(HabActual).NumeroRayos = Numero

        Line Input #nf, sData   'Laser=x
        SplitVB5 sData, "="
        Numero = sSplitDataVB5(1)
        Habitaciones(HabActual).NumeroLasers = Numero

        Line Input #nf, sData   'Fuego=x
        SplitVB5 sData, "="
        Numero = sSplitDataVB5(1)
        Habitaciones(HabActual).NumeroFuegos = Numero

        Line Input #nf, sData   'Cataratas=x
        SplitVB5 sData, "="
        Numero = sSplitDataVB5(1)
        Habitaciones(HabActual).NumeroCataratas = Numero


        Line Input #nf, sData   '--- Salidas ---

        For i = 1 To Habitaciones(HabActual).NumeroSalidas 'NumSal
            Line Input #nf, sData
            SplitVB5 sData, ","
            Habitaciones(HabActual).Salidas(i).OrgX1 = sSplitDataVB5(0)
            Habitaciones(HabActual).Salidas(i).OrgY1 = sSplitDataVB5(1)
            Habitaciones(HabActual).Salidas(i).OrgX2 = sSplitDataVB5(2)
            Habitaciones(HabActual).Salidas(i).OrgY2 = sSplitDataVB5(3)
            Habitaciones(HabActual).Salidas(i).DestinoX = sSplitDataVB5(4)
            Habitaciones(HabActual).Salidas(i).DestinoY = sSplitDataVB5(5)
            Habitaciones(HabActual).Salidas(i).HabitacionDestino = sSplitDataVB5(6)
        Next i

        Line Input #nf, sData   '--- Lamparas ---

        For i = 1 To Habitaciones(HabActual).NumeroLamparas 'NumLam
            Line Input #nf, sData
            SplitVB5 sData, ","
            Habitaciones(HabActual).Lamparas(i).Estado = True
            Habitaciones(HabActual).Lamparas(i).PosX = sSplitDataVB5(0)
            Habitaciones(HabActual).Lamparas(i).PosY = sSplitDataVB5(1)
            Habitaciones(HabActual).Lamparas(i).Tipo = sSplitDataVB5(2)
        Next i

        Line Input #nf, sData   '--- Rayos ---

        For i = 1 To Habitaciones(HabActual).NumeroRayos 'NumRay
            Line Input #nf, sData
            SplitVB5 sData, ","
            Habitaciones(HabActual).Rayos(i).PosX1 = sSplitDataVB5(0)
            Habitaciones(HabActual).Rayos(i).PosY1 = sSplitDataVB5(1)
            Habitaciones(HabActual).Rayos(i).TempY = sSplitDataVB5(1)
            Habitaciones(HabActual).Rayos(i).PosX2 = sSplitDataVB5(2)
            Habitaciones(HabActual).Rayos(i).InvisX1 = sSplitDataVB5(3)
            Habitaciones(HabActual).Rayos(i).tempX = sSplitDataVB5(3)
            Habitaciones(HabActual).Rayos(i).InvisX2 = sSplitDataVB5(4)
            Habitaciones(HabActual).Rayos(i).Velo = sSplitDataVB5(5)
            Habitaciones(HabActual).Rayos(i).Visible = False
            Habitaciones(HabActual).Rayos(i).Stop = False
        Next i

        Line Input #nf, sData   '--- Lasers ---

        For i = 1 To Habitaciones(HabActual).NumeroLasers 'NumLas
            Line Input #nf, sData
            SplitVB5 sData, ","
            Habitaciones(HabActual).Laser(i).PosX1 = sSplitDataVB5(0)
            Habitaciones(HabActual).Laser(i).tempX = sSplitDataVB5(0)
            Habitaciones(HabActual).Laser(i).PosY1 = sSplitDataVB5(1)
            Habitaciones(HabActual).Laser(i).Distancia = sSplitDataVB5(2)
            Habitaciones(HabActual).Laser(i).Direccion = sSplitDataVB5(3)
            Habitaciones(HabActual).Laser(i).Stop = False
        Next i

        Line Input #nf, sData   '--- Fuegos ---

        For i = 1 To Habitaciones(HabActual).NumeroFuegos 'NumFue
            Line Input #nf, sData
            SplitVB5 sData, ","
            Habitaciones(HabActual).Fuegos(i).PosX = sSplitDataVB5(0)
            Habitaciones(HabActual).Fuegos(i).PosY = sSplitDataVB5(1)
            Habitaciones(HabActual).Fuegos(i).Tipo = sSplitDataVB5(2)
            Habitaciones(HabActual).Fuegos(i).Activo = False
            Habitaciones(HabActual).Fuegos(i).Frame = 0
            Habitaciones(HabActual).Fuegos(i).Espera = 0
        Next i

        Line Input #nf, sData   '--- Cataratas ---

        For i = 1 To Habitaciones(HabActual).NumeroCataratas 'NumCat
            Line Input #nf, sData
            SplitVB5 sData, ","
            Habitaciones(HabActual).Cataratas(i).PosX = sSplitDataVB5(0)
            Habitaciones(HabActual).Cataratas(i).PosY = sSplitDataVB5(1)
            Habitaciones(HabActual).Cataratas(i).TempY = 0
            Habitaciones(HabActual).Cataratas(i).PosY2 = sSplitDataVB5(2)
            Habitaciones(HabActual).Cataratas(i).Tipo = sSplitDataVB5(3)
            Habitaciones(HabActual).Cataratas(i).Velocidad = 1
            Habitaciones(HabActual).Cataratas(i).Direccion = HaciaArriba
        Next i
    Loop
Close #nf
End Sub



Public Function CanPlaySound() As Integer
  Dim i As Integer

      i = AUDIO_NONE
      If waveOutGetNumDevs > 0 Then
         i = AUDIO_WAVE
      End If
    
      If midiOutGetNumDevs > 0 Then
         i = i + AUDIO_MIDI
      End If
    
      CanPlaySound = i

End Function





Public Sub SplitVB5(ByVal sIn As String, _
    Optional sDelim As String = " ", _
    Optional nLimit As Long = -1, _
    Optional bCompare As VbCompareMethod = vbBinaryCompare)


    Dim nC As Long, nPos As Long, nDelimLen As Long
    If sDelim <> "" Then
        nDelimLen = Len(sDelim)
        nPos = InStr(1, sIn, sDelim, bCompare)
        Do While nPos
            ReDim Preserve sSplitDataVB5(nC)
            sSplitDataVB5(nC) = Left(sIn, nPos - 1)
            sIn = Mid(sIn, nPos + nDelimLen)
            nC = nC + 1
            If nLimit <> -1 And nC >= nLimit Then Exit Do
            nPos = InStr(1, sIn, sDelim, bCompare)
        Loop
    End If

    ReDim Preserve sSplitDataVB5(nC)
    sSplitDataVB5(nC) = sIn
    
End Sub

