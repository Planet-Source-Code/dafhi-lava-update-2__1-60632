VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Lava by dafhi
'Update 2 - July 29, 2005

'See module 'mMyAlgorithms'
'for pixel plotting and flow algorithms

Private Sub Form_Load()
 
    Randomize
    
    Move 100, 1000, 2900, 2500
    
    Show
    
    FPS_Init
 
    ForeColor = vbWhite
    
    Do While DoEvents
        
        RenderMelt hDC
        
        CurrentY = 0
        Print "spacebar, fast transition"
        Print " z, slow"
        
        If CheckFPS(, 0.003) Then
            Caption = "FPS: " & Round(sFPS, 1)
        End If
        
        Sleep 1
    
    Loop
 
End Sub

Private Sub Form_KeyDown(IntKey As Integer, Shift As Integer)
Dim sTemp As Single
    Select Case IntKey
    Case vbKeyEscape
        Unload Me
    Case vbKeySpace
        sTemp = Rnd * TwoPi
        SummonMelt 0.3 * speed, 30, , 0 * Cos(sTemp), 0 * Sin(sTemp)
    Case vbKeyZ, vbKeyX, vbKeyC, vbKeyV, vbKeyB
        sTemp = Rnd * TwoPi
        SummonMelt 0.05 * (1 + Rnd * 0.5) * speed, , , 6 * Cos(sTemp), 6 * Sin(sTemp)
    End Select
End Sub

Private Sub Form_Resize()

    ScaleMode = vbPixels
    SizeMelt ScaleWidth, ScaleHeight

End Sub
