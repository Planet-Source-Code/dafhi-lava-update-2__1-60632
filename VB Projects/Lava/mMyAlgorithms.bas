Attribute VB_Name = "mMyAlgorithms"
Option Explicit

'Dependencies: mRenderMelt.bas, mGeneral.bas
'optional dependencies:  mParticle.bas, mIterator.bas

'******************************************'
'*                                        *'
'*  module for routines you write         *'
'*                                        *'
'******************************************'

'change this based on how many effects are programmed
'in LavaFlows() so they can all be utilized
Public Const cFlowFX As Long = 7

Dim dx_  As Single 'system

Dim cx   As Single 'custom
Dim cy   As Single
Dim ngl  As Single
Dim ngl2 As Single
Dim ngl3 As Single
Dim ngl4 As Single
Dim rad  As Single
Dim rad2 As Single
Dim rad3 As Single
Dim rad4 As Single
Dim dx   As Single
Dim dy   As Single
Dim dx1  As Single
Dim dy1  As Single
Dim dx2  As Single
Dim dy2  As Single
Dim sx   As Single
Dim sy   As Single
Dim tmp1 As Single
Dim tmp2 As Single

Dim a0   As Single
Dim a1   As Single
Dim a2   As Single
Dim a3   As Single
Dim a4   As Single

Dim d0   As Single
Dim d1   As Single

Dim srcR As Single
Dim r    As Single
Dim srcT As Single

Dim xa_  As Single 'skew grid
Dim xb_  As Single
Dim xc_  As Single
Dim xd_  As Single
Dim ya_  As Single
Dim yb_  As Single
Dim yc_  As Single
Dim yd_  As Single
Dim x_trackL     As Single
Dim y_trackL     As Single
Dim x_trackR     As Single
Dim y_trackR     As Single
Dim x_alpha      As Single
Dim y_alpha      As Single

Dim m_hue    As Single
Dim Particle As AaParticle 'Option - mParticle.bas

'Array stuff
Dim Y_LEF   As Long
Dim Lng1D() As Long
Dim SA1     As SAFEARRAY1D


'Two subs provide framework for drawing to
'the 'melt renderer'. They are called from
'mRenderMelt.

Public Sub MyDraw(Ary2D() As Long, SurfDest As SurfaceDescriptor)
Dim sng1      As Single
Dim DrawColor As Long
Dim X1        As Long
Dim sHigh     As Single
Dim AaDraw    As Long

    '*********************************'
    '*                               *'
    '*  Draw particles               *'
    '*                               *'
    '*********************************'
    
    If speed <= 0 Then speed = 1 'frame per second variable (mGeneral.bas)
    
    'triangle wave for adjusting the color saturation
    'Returns: -0.25 <= Triangle <= 0.25
    sng1 = Triangle(Tick / 800) 'Tick from mGeneral.bas
    
    DrawColor = ARGBHSV(m_hue, 0.5 + 2 * sng1, 55 + Rnd * 200)
    m_hue = m_hue + 20 * speed 'speed from mGeneral.bas
    m_hue = m_hue - 1530 * Int(m_hue / 1530)
  
    'creating a pointer to Ary2D so I can use DrawSpot() which draws an antialiased particle.
    'DrawSpot() processes via 1D
    m_CreatePtr2Dto1D Ary2D
    
    'optional - mParticle.bas
    Particle.Alpha = 1.3
    Particle.slope = 0.8
    Particle.Color = DrawColor
  
    'triangle wave raises and lowers the wave bar
    sng1 = Triangle(Tick / 5000)
    
    AaDraw = Int(Rnd * 3)
   '1 = normal
   '2 = add
   '3 = subtract
    
    '*****************************************
    '*
    '* Two sets of dimensions:
    '*
    '*  1. graphics array beyond the window
    '*  (0 to XWIDEM, 0 to XHIGHM)
    '*
    '*  2. the actual window dimensions
    '*  (ViewCorner to ViewRgt, ViewCorner to ViewTop)
    '*
    '**********************************
    
    For X1 = 0 To XWIDEM
        sy = Int(Rnd * XHIGH)
        If Rnd < 0.3 Then Ary2D(X1, sy) = DrawColor
    Next
    
    'Draw the wave
    sHigh = XHIGH / 2 + sng1 * XHIGH * 2
    
    For X1 = 0 To XWIDEM
        sy = sHigh + (Rnd - 0.5) * 25
        
'        If sy >= 0 And sy <= XHIGH Then
'            Ary2D(X1, sy) = DrawColor
'        End If
    
        DrawSpot SurfDest, Lng1D, Particle, _
            X1, sy, 2.1, 2.1
       'brought to you by mParticle.bas
        '- have you had your particles today?

    Next
    
    'clear the pointer
    CopyMemory ByVal VarPtrArray(Lng1D), 0&, 4

End Sub

Public Sub LavaFlows(pAryX() As Single, pAryY() As Single, Optional ByVal pDelta_x!, Optional ByVal pDelta_y!, Optional ByVal Algo_ As Long = -1)

    '*********************************'
    '*                               *'
    '*  Flow pattern algorithms      *'
    '*                               *'
    '*********************************'
    
    Melt_HookArrays pAryX, pAryY 'first line  .. gives WritePixel()
    'in mRenderMelt.bas pointed access to pAryX, pAryY to allow
    'every below algorithm to have less typed code
    
    If Algo_ < 1 Or Algo_ > cFlowFX Then
        Algo_ = Int(Rnd * cFlowFX) + 1
    End If
    
    'surface center - not required for every pattern
    cx = XWIDE / 2
    cy = XHIGH / 2
    
    global_y = pDelta_y + BorderPadding 'need this line
    dx_ = pDelta_x + BorderPadding 'need this line
    
'    Algo_ = 1
'    Form1.Caption = Algo_
    
    Select Case Algo_
    Case 1
        RotoSkew
    Case 2
        Whirly
    Case 3
        Vortex
    Case 4
        Dome
    Case 5
        RotoTiles
    Case 6
        IteratorRings Int(Rnd - 0.5) 'True or False
    Case 7
        Glassfiber
    End Select
    
    Melt_UnhookArrays

End Sub

Private Sub RotoSkew()

    'zoom
    g_sk_zoom = 0.5 * speed * (Rnd - 0.5)
    ' ..           'speed', based on GetTickCount
    'rotation
    g_sk_angle = (Rnd - 0.5)
    
    SkewCorner xa_, ya_, 0.5 + Rnd, 3 / 8, 0.5
    SkewCorner xb_, yb_, 0.5 + Rnd, 1 / 8, 0.5
    SkewCorner xc_, yc_, 0.5 + Rnd, 5 / 8, 0.5
    SkewCorner xd_, yd_, 0.5 + Rnd, 7 / 8, 0.5

    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    y_alpha = Y_LEF / TopLeft
    x_trackL = xc_ + y_alpha * (xa_ - xc_)
    y_trackL = yc_ + y_alpha * (ya_ - yc_)
    x_trackR = xd_ + y_alpha * (xb_ - xd_)
    y_trackR = yd_ + y_alpha * (yb_ - yd_)
    x_alpha = 0
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
    
        melt_dx = x_trackL + x_alpha * (x_trackR - x_trackL)
        melt_dy = y_trackL + x_alpha * (y_trackR - y_trackL)
        
        WritePixel  'need this line
        x_alpha = x_alpha + 1 / BlitWid
        
    Next
    global_y = global_y + 1 'need this line
    Next
    
End Sub
Private Sub IteratorRings(Switch_ As Boolean)
    
    rad4 = 0.02 * (0.5 + Rnd) * (2 * Int(Rnd - 0.5) + 1) * Sqr(BlitWid * BlitWid + BlitHgt * BlitHgt)
    a3 = 0.014 * (3 + Rnd) * rad4 'rotation
    a2 = a3 * (Rnd + 1)         'phase
    a0 = 2 + Rnd               'complexity
    
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
        
        'radius and angle are obtained this way
        dx = cx - global_x
        dy = cy - global_y
        rad = Sqr(dx * dx + dy * dy)
        
        a1 = a3 * Iterator(a0, rad + a2, rad4, 1)
        If Switch_ Then
        ngl = GetAngle(dx, dy)
        Else
        ngl = GetAngle2(dx, dy)
        End If
        melt_dx = a1 * Cos(ngl)
        melt_dy = a1 * Sin(ngl)
        
        WritePixel  'need this line
        
    Next
    sy = sy + 1 / BlitHgt
    global_y = global_y + 1 'need this line
    Next

End Sub
Private Sub Glassfiber()

    a0 = 3.5
    a1 = Rnd * 30 + 10
    a2 = Rnd * 0.02 + 0.02
    a3 = 0.98
    a4 = 0.03
    
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
        
        dx = cx - global_x
        dy = cy - global_y
        
        ngl = GetAngle(dx, dy)
        
        rad = Sqr(dx * dx + dy * dy)
        
        d0 = rad / a0
        d1 = 1 - d0
        
        rad2 = (1 - rad * a2) * d0 + a3 * d1
        
        ngl2 = ngl + (Sin(ngl * a1) * 0.1 * d0) + ((a4 * rad * 0.01) * d1)
        
        melt_dx = rad2 * Cos(ngl2)
        melt_dy = rad2 * Sin(ngl2)
        
        WritePixel  'need this line
    Next
    global_y = global_y + 1 'need this line
    Next

End Sub


Private Sub Dome()

    rad = 0.1 * (Rnd - 0.5)
    rad2 = 0.1 * (Rnd - 0.5)
    
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
    
        'example:  displacement from center (cx, cy)
        dx = cx - global_x
        dy = cy - global_y
        
        'result to melt_dx, melt_dy
        melt_dx = Sqr(dx * dx + dy * dy) * rad
        melt_dy = Sqr(dx * dx + dy * dy) * rad2
        
        WritePixel  'need this line
    Next
    global_y = global_y + 1 'need this line
    Next

End Sub

Private Sub Vortex()

    g_sk_zoom = (0.2 + Rnd * 0.2) * Sqr(BlitWid * BlitWid + BlitHgt * BlitHgt)

    'rotation base = something * (1 or -1) ..
    g_sk_angle = (Rnd * 0.1 + 5) * (2 * Int(Rnd - 0.5) + 1) * speed
    
    SkewCorner xa_, ya_, 1 + Rnd * 0.5, 3 / 8, 0.5
    SkewCorner xb_, yb_, 1 + Rnd * 0.5, 1 / 8, 0.5
    SkewCorner xc_, yc_, 1 + Rnd * 0.5, 5 / 8, 0.5
    SkewCorner xd_, yd_, 1 + Rnd * 0.5, 7 / 8, 0.5

    rad4 = 5 * (0.5 + Rnd * 0.2) '* speed
    
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    y_alpha = Y_LEF / TopLeft
    x_trackL = xa_ + y_alpha * (xc_ - xa_)
    y_trackL = ya_ + y_alpha * (yc_ - ya_)
    x_trackR = xb_ + y_alpha * (xd_ - xb_)
    y_trackR = yb_ + y_alpha * (yd_ - yb_)
    x_alpha = 0
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
    
        sx = x_trackL + x_alpha * (x_trackR - x_trackL)
        sy = y_trackL + x_alpha * (y_trackR - y_trackL)
        'ngl = Sqr(sx * sx + sy * sy)
        
        ngl = GetAngle(-sx, sy)
        melt_dx = rad4 * Cos(ngl)
        melt_dy = rad4 * Sin(ngl)
        
        WritePixel  'need this line
        x_alpha = x_alpha + 1 / BlitWid
        
    Next
    global_y = global_y + 1 'need this line
    Next
    
End Sub

Private Sub Whirly()
    
    ngl2 = 0.05 * (Rnd + 0.1) 'frequency
    
    'amplitude base = number * (1 or -1)
    rad3 = 150 * speed * (2 * Int(Rnd - 0.5) + 1)
    rad3 = rad3 * (1 + Rnd * 3)
    
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
        
        'radius and angle are obtained this way
        dx = cx - global_x
        dy = cy - global_y
        ngl = GetAngle(dx, dy)
        rad = Sqr(dx * dx + dy * dy)
        
        rad2 = rad3 * Triangle(rad * ngl2)
        melt_dx = rad2 * Cos(ngl)
        melt_dy = rad2 * Sin(ngl)
        
        WritePixel  'need this line
        
    Next
    sy = sy + 1 / BlitHgt
    global_y = global_y + 1 'need this line
    Next
End Sub

Private Sub RotoTiles()

    'zoom
    g_sk_zoom = RndPosNeg * (0.2 + Rnd)
    
    'rotation
    g_sk_angle = Rnd
    
    SkewCorner xa_, ya_, 1 + Rnd * 0.15, 3 / 8, 0.5
    SkewCorner xb_, yb_, 1 + Rnd * 0.15, 1 / 8, 0.5
    SkewCorner xc_, yc_, 1 + Rnd * 0.15, 5 / 8, 0.5
    SkewCorner xd_, yd_, 1 + Rnd * 0.15, 7 / 8, 0.5
    
    rad2 = 0.2 * g_sk_zoom * (1 + Rnd * 4)
    rad3 = 0.2 * g_sk_zoom * (1 + Rnd * 4)

    For Y_LEF = BotLeft To TopLeft Step XWIDE
    global_x = dx_ 'need this line
    y_alpha = Y_LEF / TopLeft
    x_trackL = xc_ + y_alpha * (xa_ - xc_)
    y_trackL = yc_ + y_alpha * (ya_ - yc_)
    x_trackR = xd_ + y_alpha * (xb_ - xd_)
    y_trackR = yd_ + y_alpha * (yb_ - yd_)
    x_alpha = 0
    For m_Pos1D = Y_LEF To Y_LEF + BlitWidM
    
        sx = x_trackL + x_alpha * (x_trackR - x_trackL)
        sy = y_trackL + x_alpha * (y_trackR - y_trackL)
        
        melt_dx = rad2 * Triangle(sx)
        melt_dy = rad3 * Triangle(sy)
    
        WritePixel  'need this line
        x_alpha = x_alpha + 1 / BlitWid
        
    Next
    global_y = global_y + 1 'need this line
    Next

End Sub


'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_


Private Sub m_CreatePtr2Dto1D(Ary2D() As Long)

    'creating a pointer to 2d array to create
    '1d array that AaParticle understands
    SA1.cbElements = 4
    SA1.cDims = 1
    SA1.cElements = Alice.UBound + 1
    SA1.pvData = VarPtr(Ary2D(0, 0))
    CopyMemory ByVal VarPtrArray(Lng1D), VarPtr(SA1), 4
    
End Sub

