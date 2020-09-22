Attribute VB_Name = "mRenderMelt"
Option Explicit

'mRenderMelt.bas by dafhi

'Dependencies:  mGeneral.bas, mMyAlgorithms.bas

Public global_x As Single 'meant as read-only from mMyLava
Public global_y As Single

Public melt_dx As Single
Public melt_dy As Single

Public XWIDE      As Long
Public XHIGH      As Long
Public XWIDEM     As Long
Public XHIGHM     As Long

Public BlitWid    As Long
Public BlitHgt    As Long
Public BlitWidM   As Long
Public BlitHgtM   As Long
Public LimRgt     As Long
Public LimTop     As Long
Public ViewRgt    As Long
Public ViewTop    As Long

Public Alice      As SurfaceDescriptor   'surfaces cleared in Form_Unload
Public Carmen     As SurfaceDescriptor
Dim Flipped       As Boolean

Public BotLeft    As Long '1d array stuff :P
Public TopLeft    As Long

Dim navigation As Single
Dim nav_result As Single
Dim nav_speed  As Single
Dim hold_time  As Single
Public NavDirectn As Long
Dim Navigating As Boolean
Dim LookX1()   As Single
Dim LookY1()   As Single
Dim LookX2()   As Single
Dim LookY2()   As Single
Dim sx         As Single
Dim sy         As Single

Public AryX()     As Single 'Pointer stuff
Public AryY()     As Single

Public m_Pos1D    As Long
Dim SA1 As SAFEARRAY1D
Dim SA2 As SAFEARRAY1D

Dim m_SA_DibSrc As SAFEARRAY1D
Dim m_SA_DibDest As SAFEARRAY1D

Dim m_DibSrc() As Long
Dim m_DibDest() As Long

Public Const BorderPadding As Long = 5
Public Const ViewCorner    As Long = BorderPadding
Public Const ViewCrnrP1    As Long = BorderPadding + 1

Public Sub RenderMelt(hDC As Long)

    If Flipped Then
        RenderMelt_Unwrapped hDC, Carmen.Dib32, Alice.Dib32, Alice, Carmen
    Else
        RenderMelt_Unwrapped hDC, Alice.Dib32, Carmen.Dib32, Carmen, Alice
    End If
    
    Flipped = Not Flipped

End Sub

Private Sub RenderMelt_Unwrapped(hDC As Long, Dest() As Long, Src() As Long, SurfSrc As SurfaceDescriptor, SurfDest As SurfaceDescriptor)
Dim PelA&
Dim PelB&
Dim PelC&
Dim PelD&
Dim I1&
Dim Y_LEF As Long
Dim Red_ As Long
Dim Grn_ As Long
Dim Blu_ As Long
Dim LefCol As Long
Dim BotRow As Long
Dim wu_edge_left As Single
Dim wu_edge_bottom As Single
Dim alpha_left As Single
Dim alpha_bottom As Single
Dim src_x As Single
Dim src_y As Single

    If BlitWid < 2 Or BlitHgt < 2 Then Exit Sub
  
    'changing melt pattern
    If navigation >= hold_time Then
        navigation = 0
        SetLooks
        NavDirectn = 1 - NavDirectn
    End If
    If NavDirectn = 1 Then
        nav_result = 1 - navigation
        If nav_result < 0 Then nav_result = 0
    Else
        nav_result = navigation
        If nav_result > 1 Then nav_result = 1
    End If
    
'    Form1.Caption = Round(nav_result, 2)

    m_MeltHookArraysDibSurface m_SA_DibSrc, m_SA_DibDest, Src, Dest
  
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    For I1 = Y_LEF To Y_LEF + BlitWidM
    
    sx = LookX1(I1)
    src_x = sx + nav_result * (LookX2(I1) - sx)
            
    sy = LookY1(I1)
    src_y = sy + nav_result * (LookY2(I1) - sy)
    
    wu_edge_left = src_x - 0.5
    wu_edge_bottom = src_y - 0.5
     
    LefCol = Int(src_x)
    BotRow = Int(src_y)
     
    alpha_left = LefCol + 0.5 - wu_edge_left
    alpha_bottom = BotRow + 0.5 - wu_edge_bottom
    
    LefCol = BotRow * XWIDE + LefCol
    PelC = m_DibSrc(LefCol)
    
    LefCol = LefCol + 1
    PelD = m_DibSrc(LefCol)
    
    LefCol = LefCol + XWIDE
    PelB = m_DibSrc(LefCol)
    
    LefCol = LefCol - 1
    PelA = m_DibSrc(LefCol)
    
    Red_ = ((PelA And MaskHIGH) - (PelB And MaskHIGH)) \ L65536
    Grn_ = ((PelA And &HFF00&) - (PelB And &HFF00&)) \ 256&
    Blu_ = (PelA And &HFF&) - (PelB And &HFF&)
    Red_ = alpha_left * Red_
    Grn_ = alpha_left * Grn_
    Blu_ = alpha_left * Blu_
    PelA = PelB + Blu_ + 256& * (Grn_ + 256& * Red_)
       
    Red_ = ((PelC And MaskHIGH) - (PelD And MaskHIGH)) \ L65536
    Grn_ = ((PelC And &HFF00&) - (PelD And &HFF00&)) \ 256&
    Blu_ = (PelC And &HFF&) - (PelD And &HFF&)
    Red_ = alpha_left * Red_
    Grn_ = alpha_left * Grn_
    Blu_ = alpha_left * Blu_
    PelC = PelD + Blu_ + 256& * (Grn_ + 256& * Red_)
    
    Red_ = ((PelC And MaskHIGH) - (PelA And MaskHIGH)) \ L65536
    Grn_ = ((PelC And &HFF00&) - (PelA And &HFF00&)) \ 256&
    Blu_ = (PelC And &HFF&) - (PelA And &HFF&)
    Red_ = alpha_bottom * Red_
    Grn_ = alpha_bottom * Grn_
    Blu_ = alpha_bottom * Blu_
    m_DibDest(I1) = PelA + Blu_ + 256& * (Grn_ + 256& * Red_)
    
    Next
    Next
    
    CopyMemory ByVal VarPtrArray(m_DibSrc), 0&, 4
    CopyMemory ByVal VarPtrArray(m_DibDest), 0&, 4
  
    MyDraw Dest, SurfDest
    
    navigation = navigation + nav_speed
  
    StretchDIBits hDC, 0, 0, BlitWid, BlitHgt, BorderPadding, BorderPadding, BlitWid, BlitHgt, Dest(0, 0), Alice.BIH, DIB_RGB_COLORS, vbSrcCopy
  
End Sub
Private Sub m_MeltHookArraysDibSurface(SA_Src As SAFEARRAY1D, SA_Dest As SAFEARRAY1D, pArySrc() As Long, pAryDest() As Long)

    SA_Src.cDims = 1
    SA_Src.cbElements = 4
    SA_Src.cElements = XWIDE * XHIGH
    SA_Dest = SA_Src
    
    SA_Src.pvData = VarPtr(pArySrc(0, 0))
    CopyMemory ByVal VarPtrArray(m_DibSrc), VarPtr(SA_Src), 4
    SA_Dest.pvData = VarPtr(pAryDest(0, 0))
    CopyMemory ByVal VarPtrArray(m_DibDest), VarPtr(SA_Dest), 4

End Sub

Public Sub SetLooks(Optional ByVal dx_!, Optional ByVal dy_!, Optional ByVal NavD As Long = -1, Optional ByVal EffectNum As Long = -1)
    
    If NavD = -1 Then NavD = NavDirectn
    
    If NavD > 0 Then
        LavaFlows LookX2, LookY2, dx_, dy_, EffectNum
    Else
        LavaFlows LookX1, LookY1, dx_, dy_, EffectNum
    End If

End Sub
Public Sub SummonMelt(Optional ByVal morph_quickness! = 0.02, Optional ByVal display_time As Single = 50, Optional ByVal EffectNum As Long = -1, Optional ByVal dx_!, Optional ByVal dy_!)
    
    m_MeltPosition LookX2, LookY2, LookX1, LookY1
    SetLooks dx_, dy_, 1, EffectNum
    navigation = 0
    NavDirectn = 0
    
    If morph_quickness < 0 Then morph_quickness = 0
    nav_speed = morph_quickness
    
    hold_time = display_time * nav_speed + 1
    If hold_time < 1 Then hold_time = 1

End Sub
Private Sub m_MeltPosition(s_xfg!(), s_yfg!(), s_xbg!(), s_ybg!())
Dim I1    As Long
Dim Y_LEF As Long
    
    For Y_LEF = BotLeft To TopLeft Step XWIDE
    For I1 = Y_LEF To Y_LEF + BlitWidM
        sx = s_xbg(I1)
        sy = s_ybg(I1)
        s_xbg(I1) = sx + nav_result * (s_xfg(I1) - sx)
        s_ybg(I1) = sy + nav_result * (s_yfg(I1) - sy)
    Next
    Next
    
End Sub

Public Sub SizeMelt(Wid&, Hgt&, Optional ByVal speed_multiplier! = 1)

    If Wid < 2 Or Hgt < 2 Then Exit Sub
    If Wid = BlitWid And Hgt = BlitHgt Then Exit Sub
    
    BlitWid = Wid
    BlitHgt = Hgt
 
    BlitWidM = BlitWid - 1
    BlitHgtM = BlitHgt - 1
    
    ViewRgt = BorderPadding + BlitWidM
    ViewTop = BorderPadding + BlitHgtM
    
    XWIDE = Wid + 2 * BorderPadding
    XHIGH = Hgt + 2 * BorderPadding
    
    XWIDEM = XWIDE - 1
    XHIGHM = XHIGH - 1
    
    LimRgt = ViewRgt - 1
    LimTop = ViewTop - 1
    
    SetSurfaceDesc Alice, Alice.Dib32, XWIDE, XHIGH
    SetSurfaceDesc Carmen, Carmen.Dib32, XWIDE, XHIGH
    
    m_ReDim LookX1, Alice.UBound
    m_ReDim LookY1, Alice.UBound
    m_ReDim LookX2, Alice.UBound
    m_ReDim LookY2, Alice.UBound
    
    BotLeft = XWIDE * BorderPadding + BorderPadding
    TopLeft = BotLeft + BlitHgtM * XWIDE
    
    Flipped = False
    
    SetLooks , , 1
    SetLooks , , 0
    
    SummonMelt
    
End Sub
Private Sub m_ReDim(Ary As Variant, UBound_&)
    Erase Ary
    ReDim Ary(UBound_)
End Sub

Public Sub WritePixel()
    melt_dx = melt_dx + global_x
    melt_dy = melt_dy + global_y
    If melt_dx < ViewCorner Then
        melt_dx = ViewCorner
    ElseIf melt_dx > LimRgt Then
        melt_dx = LimRgt
    End If
    If melt_dy < ViewCorner Then
        melt_dy = ViewCorner
    ElseIf melt_dy > LimTop Then
        melt_dy = LimTop
    End If
    AryX(m_Pos1D) = melt_dx
    AryY(m_Pos1D) = melt_dy
    global_x = global_x + 1
End Sub
Public Sub Melt_HookArrays(pAryX() As Single, pAryY() As Single)
    
    'create pointers to pAryX, pAryY so that
    'WriteField() can access them as AryX, AryY
    SA1.cDims = 1
    SA1.cbElements = 4
    SA1.cElements = UBound(pAryX) + 1
    SA2 = SA1
    
    SA1.pvData = VarPtr(pAryX(0))
    CopyMemory ByVal VarPtrArray(AryX), VarPtr(SA1), 4
    SA2.pvData = VarPtr(pAryY(0))
    CopyMemory ByVal VarPtrArray(AryY), VarPtr(SA2), 4
    
End Sub
Public Sub Melt_UnhookArrays()
    CopyMemory ByVal VarPtrArray(AryX), 0&, 4
    CopyMemory ByVal VarPtrArray(AryY), 0&, 4
End Sub
