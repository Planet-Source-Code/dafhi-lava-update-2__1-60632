Attribute VB_Name = "mGeneral"
Option Explicit

'mGeneral.bas by dafhi  July 29, 2005

'This module contains type declarations, subs, variables, constants,
'and functions I use alot.  This module will probably become huge.

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type RGBTriple
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Type RGBQUAD
 Blue  As Byte
 Green As Byte
 Red   As Byte
 Alpha As Byte
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

Public Type SurfaceDescriptor
    BIH             As BITMAPINFOHEADER
    Wide            As Long
    High            As Long
    WM              As Long
    HM              As Long
    hDC             As Long
    UBound      As Long
    cDims     As Long
    Dib32()         As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Dim I As Long
Dim J As Long

Public Const pi As Double = 3.14159265358979
Public Const TwoPi As Double = 2 * pi
Public Const piBy2 As Single = pi / 2
Public Const halfPi As Single = piBy2

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const NOTE_1OF12 As Double = 2 ^ (1 / 12)

Public Const ASC_DOUBLE_QUOTE As Integer = 34

Dim LBA  As Long
Dim UBA  As Long
Dim LenA As Long

Dim MinusFnd  As Boolean
Dim PointFnd  As Boolean

Private StrTemp As String

Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'ARGBHSV() Function
Public Blu_&
Public Grn_&
Public Red_&
Public subt!

Public Const DIB_RGB_COLORS As Long = 0
Public Const GrayScaleRGB As Long = 1 + 256& + 65536

Public Const MaskHIGH       As Long = &HFF0000
Public Const MaskMID        As Long = &HFF00
Public Const MaskLOW        As Long = &HFF
Public Const MaskRB         As Long = &HFF00FF
Public Const L65536         As Long = 65536

'skew corner
Public g_sk_zoom   As Single
Public g_sk_angle  As Single

'CheckFPS()
Public Tick       As Long
Public FrameCount As Long
Public Elapsed    As Long
Public speed      As Single
Public sFPS       As Single

Private PrevTick   As Long
Private NextTick   As Long
Private FirstTick As Long
Private TickSum   As Long

Private Const Interval_Micro As Long = 4

Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)
Declare Function StretchDIBits Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         lpBits As Any, _
         lpBitsInfo As BITMAPINFOHEADER, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long


'****************'
'*              *'
'*   Graphics   *'
'*              *'
'****************'

Sub FPS_Init() 'right before game loop
    PrevTick = timeGetTime
'    PrevMicro = PrevTick
    FirstTick = PrevTick
    NextTick = PrevTick + Interval_Micro
'    NextMicro = NextTick
End Sub
Function CheckFPS(Optional RetFPS, Optional ByVal speed_coefficient As Single = 1, Optional Interval_Millisec& = 200) As Boolean
    
'CODE SAMPLE
'1. Paste comments below to Form
'2. hit ctrl-h
'3. line 1 says [comment mark][1 space] (2 characters total)
'4. line 2 says nothing
'5. Replace All
'6. be sure to reference mGeneral.bas
    
' Private Sub Form_Load()
    ' FPS_Init 'initialize time variables
    ' Do While DoEvents '"very simple game loop"
        
        ' Cls
        ' Print "posx = posx + dx * speed
        ' print "speed is smaller for faster CPU
        
        ' If CheckFPS(FPS, speed_multiplier, 200) Then
        '    Caption = "FPS: " & FPS
        ' End If
    ' Loop
' End Sub
    
    Tick = timeGetTime

    
    FrameCount = FrameCount + 1
    TickSum = Tick - FirstTick
    speed = speed_coefficient * (TickSum / FrameCount)
    If Tick >= NextTick Then
        RetFPS = 1000 * FrameCount / TickSum
        PrevTick = Tick
        sFPS = RetFPS
        NextTick = Tick + Interval_Millisec
        CheckFPS = True
    Else
        CheckFPS = False
    End If

End Function

Public Function RGBHSV(hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat As Single
Dim value1      As Single
Dim diff1       As Single
Dim maxim       As Single

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    Blu_ = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     Red_ = Int(value1)
     Grn_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     Grn_ = Int(value1)
     Red_ = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    Red_ = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     Grn_ = Int(value1)
     Blu_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     Blu_ = Int(value1)
     Grn_ = Int(value1 - hue_and_sat)
    End If
   Else
    Grn_ = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     Blu_ = Int(value1)
     Red_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     Red_ = Int(value1)
     Blu_ = Int(value1 - hue_and_sat)
    End If
   End If
   RGBHSV = Red_ Or Grn_ * 256& Or Blu_ * 65536
  Else 'saturation_0_To_1 <= 0
   RGBHSV = Int(value1) * CLng(65793) '1 + 256 + 65536
  End If
 Else 'value_0_To_255 <= 0
  RGBHSV = 0&
 End If
End Function
Public Function ARGBHSV(hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat As Single
Dim value1      As Single
Dim diff1       As Single
Dim maxim       As Single

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    Blu_ = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     Red_ = Int(value1)
     Grn_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     Grn_ = Int(value1)
     Red_ = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    Red_ = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     Grn_ = Int(value1)
     Blu_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     Blu_ = Int(value1)
     Grn_ = Int(value1 - hue_and_sat)
    End If
   Else
    Grn_ = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     Blu_ = Int(value1)
     Red_ = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     Red_ = Int(value1)
     Blu_ = Int(value1 - hue_and_sat)
    End If
   End If
   ARGBHSV = Red_ * 65536 Or Grn_ * 256& Or Blu_
  Else 'saturation_0_To_1 <= 0
   ARGBHSV = Int(value1) * CLng(65793) '1 + 256 + 65536
  End If
 Else 'value_0_To_255 <= 0
  ARGBHSV = 0&
 End If
End Function
Public Function FlipRB(Color_ As Long) As Long
Dim LBlu As Long
    LBlu = Color_ And &HFF&
    FlipRB = (Color_ And &HFF00&) + 256& * (LBlu * 256&) + (Color_ \ 256&) \ 256&
End Function


Public Sub SetSurfaceDesc(SDesc1 As SurfaceDescriptor, Ary_() As Long, ByVal Wide As Long, ByVal High As Long, Optional ByVal lhDC As Long, Optional Set1D As Boolean = False)
    If Wide * High < 10000000 And Wide * High > 0 Then
        SDesc1.High = High
        SDesc1.Wide = Wide
        SDesc1.WM = SDesc1.Wide - 1
        SDesc1.HM = SDesc1.High - 1
        SDesc1.hDC = lhDC
        SDesc1.BIH.biHeight = High
        SDesc1.BIH.biWidth = Wide
        SDesc1.BIH.biPlanes = 1
        SDesc1.BIH.biBitCount = 32
        SDesc1.BIH.biSize = Len(SDesc1.BIH)
        SDesc1.BIH.biSizeImage = 4 * Wide * High
        SDesc1.UBound = Wide * High - 1
        If Set1D Then
            Erase Ary_
            ReDim Ary_(SDesc1.UBound)
            SDesc1.cDims = 1
        Else
            Erase Ary_
            ReDim Ary_(SDesc1.WM, SDesc1.HM)
            SDesc1.cDims = 2
        End If
    End If
End Sub

'********************'
'*                  *'
'*   String stuff   *'
'*                  *'
'********************'

Function Eval(ByVal StrIn As String, Optional ByVal Pos_ As Long = 1, Optional RetPos As Variant) As Double
Dim numerat   As Double
Dim denominat As Double
Dim NuFnd     As Boolean
Dim DeFnd     As Boolean

    'simple expression evaluation recognizes:
    'ONE minus sign ONE decimal, ONE division
    'with garbage handling
    
    ' "/-2" = -0.5
    ' "0/0" = 1
    ' "-.2.2.-1" = -0.221

    LenA = Len(StrIn)
    
    mEval1 StrIn, NuFnd, Pos_
    numerat = Val(Mid$(StrIn, Pos_, I + NuFnd))
    If NuFnd Then
        If numerat = 0 Then numerat = 1
        Add I, 1
        J = I
        mEval1 StrIn, NuFnd, J
        denominat = Val(Right$(StrIn, I - J))
        RetPos = J
    Else
        RetPos = I
    End If
    
    If denominat = 0 Then denominat = 1
    Eval = numerat / denominat
    
End Function
Private Sub mEval1(StrIn As String, TmpBool As Boolean, StartPos As Long)
    PointFnd = False
    MinusFnd = False
    I = StartPos
    Do While I <= LenA
        StrTemp = Mid$(StrIn, I, 1)
        If StrTemp = "/" Or StrTemp = "\" Then
            TmpBool = True
            Exit Do
        ElseIf StrTemp = "." Then
            If PointFnd Then
                mEval2 StrIn
            Else
                Add I, 1
            End If
            PointFnd = True
        ElseIf StrTemp = "-" Then
            If MinusFnd Then
                mEval2 StrIn
            Else
                Add I, 1
            End If
            MinusFnd = True
        ElseIf Asc(StrTemp) < 48 Or Asc(StrTemp) > 57 Then
'            mEval2 StrIn
            Exit Do
        Else
            Add I, 1
        End If
    Loop
End Sub
Private Sub mEval2(StrIn As String)
    StrIn = Left$(StrIn, I - 1) + Right$(StrIn, LenA - I)
    Add LenA, -1
End Sub

Sub FillBytesFromString(Bytes1() As Byte, ByVal Str1 As String)
    LBA = LBound(Bytes1)
    UBA = UBound(Bytes1)
    StrTemp = Left$(Str1, UBA - LBA + 1)
    For I = LBA To UBA
        Bytes1(I) = Asc(Mid$(StrTemp, I + 1, 1))
    Next
End Sub
Function StringFromBytes(Bytes() As Byte) As String
Dim J1&
    LenA = UBound(Bytes) - LBound(Bytes) + 1
    If LenA > 0 Then
        StringFromBytes = Bytes
        StringFromBytes = StringFromBytes + StringFromBytes
        J1 = 1
        For I = LBound(Bytes) To UBound(Bytes)
            Mid$(StringFromBytes, J1, 1) = Chr$(Bytes(I))
            J1 = J1 + 1
        Next
    End If
End Function
Function GetLine(StrInput As String, ByVal Pos_ As Long, Optional RetPos As Long)
    If Pos_ > Len(StrInput) Then
        GetLine = ""
        RetPos = Pos_
        Exit Function
    End If
    For I = Pos_ To Len(StrInput)
        J = Asc(Mid$(StrInput, I, 1))
        If I = 10 Or I = 13 Then Exit For
    Next
    GetLine = Mid$(StrInput, Pos_, I - Pos_)
    RetPos = Pos_
End Function


' == File ==
Function IsFile(strFileSpec As String) As Boolean
    If strFileSpec = "" Then Exit Function
    If Len(Dir$(strFileSpec)) > 0 Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function
Function ValidFile(strFullFileSpec As String) As Boolean
Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    ValidFile = FS.fileexists(strFullFileSpec)
End Function

' various math
Sub Add(Varia1 As Variant, ByVal value_ As Double)
    Varia1 = Varia1 + value_
End Sub
Sub Mul(Varia1 As Variant, ByVal value_ As Double)
    Varia1 = Varia1 * value_
End Sub
Sub LinearAlg(ret_ As Single, from_!, to_!, perc_!)
    ret_ = from_ + perc_ * (to_ - from_)
End Sub
Sub TruncVar(ByRef RetVal As Variant)
    RetVal = RetVal - Int(RetVal)
End Sub
Function Triangle(ByVal In_dbl#) As Double
    Triangle = In_dbl - Int(In_dbl)
    If Triangle > 0.75 Then
        Triangle = Triangle - 1
    ElseIf Triangle > 0.25 Then
        Triangle = 0.5 - Triangle
    End If
End Function
Function LMax(ByVal sVar1 As Single, ByVal sVar2 As Single)
    If sVar1 < sVar2 Then
        LMax = Int(sVar2 + 0.5)
    Else
        LMax = Int(sVar1 + 0.5)
    End If
End Function
Function LMin(ByVal sVar1 As Single, ByVal sVar2 As Single)
    If sVar1 < sVar2 Then
        LMin = Int(sVar1 + 0.5)
    Else
        LMin = Int(sVar2 + 0.5)
    End If
End Function
Sub SkewCorner(pRetSX!, pRetSY!, ByVal pRad!, ByVal pAngle_0_To_1!, Optional ByVal p_rnd_quadrant_swing_mult! = 0)
    pRad = pRad * g_sk_zoom
    pAngle_0_To_1 = (g_sk_angle + pAngle_0_To_1 + p_rnd_quadrant_swing_mult * (Rnd - 0.5) * 0.25) * TwoPi
    pRetSX = pRad * Cos(pAngle_0_To_1)
    pRetSY = pRad * Sin(pAngle_0_To_1)
End Sub
Sub RadianModulus(ByRef retAngle As Variant)
 retAngle = retAngle - TwoPi * Int(retAngle / TwoPi)
End Sub
Function TriangleModulus(ByVal in1 As Single, ByVal modulus As Single) As Single
Dim mod4!
  
    mod4 = modulus * 4
    
    'mod operation
    TriangleModulus = in1 - mod4 * Int(in1 / mod4)
    
    'triangle constraint
    If TriangleModulus > modulus * 3 Then
        TriangleModulus = TriangleModulus - mod4
    ElseIf TriangleModulus > modulus Then
        TriangleModulus = modulus * 2 - TriangleModulus
    End If
  
End Function
Function RndPosNeg() As Long
    RndPosNeg = 2 * Int(Rnd - 0.5) + 1
End Function
Function GetAngle(sngDX!, sngDY!) As Single
 If sngDY = 0 Then
  If sngDX < 0 Then
   GetAngle = pi * 3& / 2&
  ElseIf sngDX > 0 Then
   GetAngle = pi / 2&
  End If
 Else
  If sngDY > 0 Then
   GetAngle = pi - Atn(sngDX / sngDY)
  Else
   GetAngle = Atn(sngDX / -sngDY)
  End If
 End If
End Function
Function GetAngle2(sngDX!, sngDY!) As Single
 If sngDX = 0! Then
  If sngDY < 0! Then
   GetAngle2 = pi * 1.5!
  ElseIf sngDY > 0 Then
   GetAngle2 = pi * 0.5!
  End If
 Else
  If sngDX > 0! Then
   GetAngle2 = Atn(sngDY / sngDX)
  Else
   GetAngle2 = pi - Atn(sngDY / -sngDX)
  End If
 End If
End Function
