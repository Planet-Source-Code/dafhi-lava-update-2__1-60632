Attribute VB_Name = "mIterator"
Option Explicit

'mIterator.bas - experimental complex gradient generator
'July 29, 2005

Public Type IterVARS
    Depth         As Long
    mod_origin    As Single
    mod_adjust    As Single
    variance_  As Single
    output_base   As Single
    input_base    As Single
End Type

Public Function Iterwrapper(IVAR As IterVARS, ByVal input_ As Double, Optional ByVal Constrain As Boolean = False, Optional ByVal output_scale As Single = 1, Optional ByVal output_base As Single = 1) As Double
Dim sLimitLow  As Single
Dim sLimitHigh As Single

    output_base = output_base * IVAR.mod_origin
    Iterwrapper = Iterator(IVAR.Depth, input_ + IVAR.input_base, IVAR.mod_adjust, IVAR.variance_) * output_scale + output_base
    
    If Constrain Then
        sLimitLow = IVAR.mod_origin
        sLimitHigh = IVAR.output_base + sLimitLow
        sLimitLow = IVAR.output_base - sLimitLow
        If Iterwrapper < sLimitLow Then
            Iterwrapper = sLimitLow
        ElseIf Iterwrapper > sLimitHigh Then
            Iterwrapper = sLimitHigh
        End If
    End If

End Function

Public Function Iterator(ByVal Complexity As Long, ByVal input_ As Double, ByVal modulus As Double, Optional ByVal variance_ As Single = 0) As Double
Dim mod1!
Dim mod4!
Dim sng1!
Dim LC&

 Iterator = 0
 
 For LC = 1 To Complexity
 
  mod1 = modulus / LC + variance_
  mod4 = mod1 * 4!
  
  'mod operation
  sng1 = input_ - mod4 * Int(input_ / mod4)
  
  'triangle constraint
  If sng1 > mod1 * 3 Then
   sng1 = sng1 - mod4
  ElseIf sng1 > mod1 Then
   sng1 = mod1 * 2 - sng1
  End If
  
  Iterator = Iterator + sng1
 
 Next
 
End Function

Public Sub Iterator_Init(IVAR As IterVARS, Optional ByVal mod_ As Double = 1, Optional ByVal Depth As Long = 1, Optional ByVal input_base As Single = 0.5, Optional ByVal variance_ As Single = 0.1, Optional ByVal output_base As Single = 0.5)

    If Depth < 1 Then Depth = 1
    If mod_ <= 0 Then mod_ = 1
    IVAR.output_base = mod_ * output_base
    IVAR.input_base = input_base * mod_ * Depth
    
    mod_ = mod_ * 0.5
    If variance_ <= 0 Then variance_ = 0
    variance_ = variance_ * mod_
    
    IVAR.Depth = Depth
    IVAR.mod_origin = mod_
    IVAR.variance_ = variance_

    mod_ = 0
    For Depth = 1 To Depth
        mod_ = mod_ + 1 / Depth
    Next
    
    IVAR.mod_adjust = (IVAR.mod_origin - Depth * variance_) / mod_
    
End Sub

