Attribute VB_Name = "mParticle"
Option Explicit

'Dependencies:
'mGeneral.bas

Private Type CircleProc
'    Defined_        As Boolean
    cone_h          As Single
End Type

Public Type AaParticle
    Proc            As CircleProc
    Alpha           As Single
    slope           As Single
    Color           As Long
End Type

Dim ForeR           As Long
Dim ForeG           As Long
Dim ForeB           As Long
Dim ForeRGB         As Long
Dim XLEF            As Long
Dim XRIT            As Long
Dim yBot            As Long
Dim yTop            As Long
Dim sAlpha          As Single
Dim maxAlpha        As Single
Dim h1              As Single
Dim dxLeft1         As Single
Dim ix1             As Single
Dim iy1             As Single
Dim y1sq            As Single

Public Const ParticleEffect_NORMAL       As Long = 0
Public Const ParticleEffect_ADD          As Long = 1
Public Const ParticleEffect_SUBTRACT     As Long = 2
Public Const ParticleEffect_PROJECTION   As Long = 3

Private Const GrayScaleA     As Long = 1 + 256 + 65536
Private Const CUBE_256       As Long = &H1000000
Private Const SQUARE_256     As Long = &H10000

Public Sub DrawSpot(SfcDesc As SurfaceDescriptor, Lng1D() As Long, SpotA As AaParticle, ByVal x_ As Single, y_ As Single, ByVal sizex As Single, ByVal sizey As Single, Optional RenderType_0_To_2 As Long, Optional SwapRB As Boolean)
Dim X1       As Long
Dim Y1       As Long
Dim WideM    As Long

    'RECT
    XLEF = Int(x_ - sizex + 0.5)
    yBot = Int(y_ - sizey + 0.5)
    XRIT = Int(x_ + sizex + 0.5)
    yTop = Int(y_ + sizey + 0.5)
    
    'CLIP
    If XLEF < 0 Then XLEF = 0
    If yBot < 0 Then yBot = 0
    If XRIT > SfcDesc.WM Then XRIT = SfcDesc.WM
    If yTop > SfcDesc.HM Then yTop = SfcDesc.HM
    
    WideM = XRIT - XLEF
    
    ForeRGB = SpotA.Color
    
    If SwapRB Then
        X1 = ForeRGB And &HFF&
        ForeRGB = (ForeRGB And &HFF00&) + 256& * (X1 * 256&) + (ForeRGB \ 256&) \ 256&
    End If
    
    ForeR = ForeRGB And MaskHIGH
    ForeG = ForeRGB And &HFF00&
    ForeB = ForeRGB And &HFF&
    
    h1 = SpotA.Alpha * SpotA.slope
    
    If sizex <> 0 Then ix1 = h1 / sizex
    If sizey <> 0 Then iy1 = h1 / sizey
    
    dxLeft1 = (XLEF - x_) * ix1
    y_ = (yBot - y_) * iy1
    
    maxAlpha = SpotA.Alpha
    If maxAlpha > 1 Then maxAlpha = 1
    
    If RenderType_0_To_2 = ParticleEffect_NORMAL Then
        
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            X1 = SfcDesc.Wide * Y1 + XLEF
            For X1 = X1 To X1 + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_alpha Lng1D(X1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    ElseIf RenderType_0_To_2 = ParticleEffect_ADD Then
    
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            X1 = SfcDesc.Wide * Y1 + XLEF
            For X1 = X1 To X1 + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha_Add Lng1D(X1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    ElseIf RenderType_0_To_2 = ParticleEffect_PROJECTION Then
    
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            X1 = SfcDesc.Wide * Y1 + XLEF
            For X1 = X1 To X1 + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha_Projector Lng1D(X1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    ElseIf RenderType_0_To_2 = ParticleEffect_SUBTRACT Then
    
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            X1 = SfcDesc.Wide * Y1 + XLEF
            For X1 = X1 To X1 + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha_Subtract Lng1D(X1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    End If
    
End Sub


Private Sub m_alpha(Pixel_ As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = Pixel_ And MaskHIGH
    BackG = Pixel_ And &HFF00&
    BackB = Pixel_ And &HFF&
    
    BackR = sAlpha * ((ForeR - BackR) \ 256&) \ 256&
    BackG = sAlpha * (ForeG - BackG) \ 256&
    BackB = sAlpha * (ForeB - BackB) '\ instead of / for faster non-float
    
    Pixel_ = Pixel_ + BackB + 256& * (BackG + 256& * BackR)

End Sub
Private Sub m_Alpha_Add(Pixel_ As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = ((ForeR * sAlpha) And MaskHIGH) + (Pixel_ And MaskHIGH)
    BackG = (ForeG * sAlpha And &HFF00&) + (Pixel_ And &HFF00&)
    BackB = ForeB * sAlpha + (Pixel_ And &HFF&)
    
    If BackR > MaskHIGH Then BackR = MaskHIGH
    If BackG > &HFF00& Then BackG = &HFF00&
    If BackB > 255& Then BackB = 255&
    
    Pixel_ = BackB + BackG + BackR

End Sub
Private Sub m_Alpha_Subtract(Pixel_ As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = (Pixel_ And MaskHIGH) - ((ForeR * sAlpha) And MaskHIGH)
    BackG = (Pixel_ And &HFF00&) - (ForeG * sAlpha And &HFF00&)
    BackB = (Pixel_ And &HFF&) - ForeB * sAlpha
    
    If BackR < SQUARE_256 Then BackR = SQUARE_256
    If BackG < 256& Then BackG = 256&
    If BackB < 0& Then BackB = 0&
    
    Pixel_ = BackB + BackG + BackR

End Sub
Private Sub m_Alpha_Projector(Pixel_ As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = Pixel_ And MaskHIGH
    BackG = Pixel_ And &HFF00&
    BackB = Pixel_ And &HFF&
    
    BackR = sAlpha * ((ForeR + BackR) \ 256&) \ 256&
    BackG = sAlpha * (ForeG + BackG) \ 256&
    BackB = sAlpha * (ForeB + BackB)
    
    If BackR > 255 Then BackR = 255
    If BackG > 255 Then BackG = 255
    If BackB > 255 Then BackB = 255
    
    Pixel_ = BackB + 256& * (BackG + 256& * BackR)
    
End Sub


Public Sub DrawSpot2(SfcDesc As SurfaceDescriptor, PicBox As Form, SpotA As AaParticle, ByVal x_ As Single, y_ As Single, ByVal sizex As Single, ByVal sizey As Single, Optional RenderType_0_To_2 As Long, Optional SwapRB As Boolean)
Dim X1       As Long
Dim Y1       As Long
Dim WideM    As Long

    'RECT
    XLEF = Int(x_ - sizex + 0.5)
    yBot = Int(y_ - sizey + 0.5)
    XRIT = Int(x_ + sizex + 0.5)
    yTop = Int(y_ + sizey + 0.5)
    
    'CLIP
    If XLEF < 0 Then XLEF = 0
    If yBot < 0 Then yBot = 0
    If XRIT > SfcDesc.WM Then XRIT = SfcDesc.WM
    If yTop > SfcDesc.HM Then yTop = SfcDesc.HM
    
    WideM = XRIT - XLEF
    
    ForeRGB = SpotA.Color
    
    If SwapRB Then
        X1 = ForeRGB And &HFF&
        ForeRGB = (ForeRGB And &HFF00&) + 256& * (X1 * 256&) + (ForeRGB \ 256&) \ 256&
    End If
    
    ForeR = ForeRGB And MaskHIGH
    ForeG = ForeRGB And &HFF00&
    ForeB = ForeRGB And &HFF&

    h1 = SpotA.Alpha * SpotA.slope
    
    If sizex <> 0 Then ix1 = h1 / sizex
    If sizey <> 0 Then iy1 = h1 / sizey
    
    dxLeft1 = (XLEF - x_) * ix1
    y_ = (yBot - y_) * iy1
    
    maxAlpha = SpotA.Alpha
    If maxAlpha > 1 Then maxAlpha = 1
    
    If RenderType_0_To_2 = ParticleEffect_NORMAL Then
        
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            For X1 = XLEF To XLEF + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha2 PicBox, X1, Y1, PicBox.Point(X1, Y1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    ElseIf RenderType_0_To_2 = ParticleEffect_ADD Then
    
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            For X1 = XLEF To XLEF + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha_Add2 PicBox, X1, Y1, PicBox.Point(X1, Y1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
    
    ElseIf RenderType_0_To_2 = ParticleEffect_PROJECTION Then
    
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            For X1 = XLEF To XLEF + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha_Projector2 PicBox, X1, Y1, PicBox.Point(X1, Y1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    ElseIf RenderType_0_To_2 = ParticleEffect_SUBTRACT Then
    
        For Y1 = yBot To yTop
            x_ = dxLeft1
            y1sq = y_ * y_
            For X1 = XLEF To XLEF + WideM
                sAlpha = h1 - Sqr(x_ * x_ + y1sq)
                m_Alpha_Subtract2 PicBox, X1, Y1, PicBox.Point(X1, Y1)
                x_ = x_ + ix1
            Next
            y_ = y_ + iy1
        Next
        
    End If
        
End Sub


Private Sub m_Alpha2(Pic As Form, X, Y, BackColor As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = BackColor And MaskHIGH
    BackG = BackColor And &HFF00&
    BackB = BackColor And &HFF&
    
    BackR = sAlpha * ((ForeR - BackR) \ 256&) \ 256&
    BackG = sAlpha * (ForeG - BackG) \ 256&
    BackB = sAlpha * (ForeB - BackB) '\ instead of / for faster non-float
    
    Pic.PSet (X, Y), BackColor + BackB + 256& * (BackG + 256& * BackR)

End Sub
Private Sub m_Alpha_Add2(Pic As Form, X, Y, BackColor As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = ((ForeR * sAlpha) And MaskHIGH) + (BackColor And MaskHIGH)
    BackG = (ForeG * sAlpha And &HFF00&) + (BackColor And &HFF00&)
    BackB = ForeB * sAlpha + (BackColor And &HFF&)
    
    If BackR > MaskHIGH Then BackR = MaskHIGH
    If BackG > &HFF00& Then BackG = &HFF00&
    If BackB > 255& Then BackB = 255&
    
    Pic.PSet (X, Y), BackR + BackG + BackB

End Sub
Private Sub m_Alpha_Subtract2(Pic As Form, X, Y, BackColor As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = (BackColor And MaskHIGH) - ((ForeR * sAlpha) And MaskHIGH)
    BackG = (BackColor And &HFF00&) - (ForeG * sAlpha And &HFF00&)
    BackB = (BackColor And &HFF&) - ForeB * sAlpha
    
    If BackR < SQUARE_256 Then BackR = SQUARE_256
    If BackG < 256& Then BackG = 256&
    If BackB < 0& Then BackB = 0&
    
    Pic.PSet (X, Y), BackR + BackG + BackB

End Sub
Private Sub m_Alpha_Projector2(Pic As Form, X, Y, BackColor As Long)
Dim BackR    As Long
Dim BackG    As Long
Dim BackB    As Long

    If sAlpha < 0 Then
        sAlpha = 0
    ElseIf sAlpha > maxAlpha Then
        sAlpha = maxAlpha
    End If
    
    BackR = BackColor And MaskHIGH
    BackG = BackColor And &HFF00&
    BackB = BackColor And &HFF&
    
    BackR = sAlpha * ((ForeR + BackR) \ 256&) \ 256&
    BackG = sAlpha * (ForeG + BackG) \ 256&
    BackB = sAlpha * (ForeB + BackB)
    
    If BackR > 255 Then BackR = 255
    If BackG > 255 Then BackG = 255
    If BackB > 255 Then BackB = 255
    
    Pic.PSet (X, Y), BackB + 256& * (BackG + 256& * BackR)

End Sub
