VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2280
      Top             =   3180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const WAIT      As Long = 10        ' milliseconds
Const G         As Double = 0.25 ' G Force in milliseconds * WAIT
Const WIND      As Double = -0.5 ' Wind factor

Private blnAnimateStart As Boolean
Private blnAnimationSet As Boolean


Private Type aPiece
    xStart              As Double
    yStart              As Double
    X                   As Double       ' x position on surface
    Y                   As Double       ' y position on surface
    Vx                  As Double       ' move after Vx milliseconds
    Vy                  As Double       ' move after Vy milliseconds
    VxNext              As Double
    VyNext              As Double
    VxMax               As Double
    VyMax               As Double
    VxMaxNext           As Double
    VyMaxNext           As Double
    Color               As Long
    ColorR              As Long         ' Red color of piece
    ColorG              As Long         ' Red color of piece
    ColorB              As Long         ' Red color of piece
    Fade                As Long         ' Will the piece fade to black or not
    xLast               As Double       ' Last Displayed X Pos
    yLast               As Double       ' Last Displayed y Pos
    Size                As Long         ' Size
    SizeNext            As Long
    Type                As Long
    TypeNext            As Long
End Type

Private arrPieces() As aPiece





Private Function CreatePieces(ByVal lngNumPieces As Long, _
                              ByVal lngFadeIn As Long, _
                              ByVal lngVxMax As Double, _
                              ByVal lngVyMax As Double, _
                              ByVal lngSize As Long, _
                              ByVal X As Double, _
                              ByVal Y As Double) As Long

    Dim lngPiece    As Long
    Dim lngX        As Long
    Dim lngY        As Long
    
    ReDim arrPieces(lngNumPieces - 1)
    
    Me.Cls
    
    Call Timer1_Timer
    
    Exit Function

End Function


Private Sub Form_Load()
    'Call AlwaysOnTop(Me, True)
    'Call ShowCursor(False)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'End
    If blnAnimateStart = True Then
        Unload Me
    Else
        blnAnimateStart = True
        Call CreatePieces(1000, True, 5, 5, 5, X, Y)
        Call AnimatePieces
    End If

End Sub


Function AnimatePieces()

    Dim lngPiece        As Long
    Dim lngMaxSpeed     As Long
    Dim lngExp          As Long
    Dim lngCol          As Long
    
Go:
    
    'Me.Cls
    
    For lngPiece = 0 To UBound(arrPieces)
    
        With arrPieces(lngPiece)
            
            .Color = .Color - 8
            If .Color < 0 Then .Color = 0
            
            .Vy = .Vy + G
            .Vx = .Vx + WIND
            .X = .X + .Vx
            .Y = .Y + .Vy
            
            .Vx = .Vx * 26 / 30
            .Vy = .Vy * 26 / 30
            
            'Circle (.xLast, .yLast), .Size, 0
            'PSet (.xLast, .yLast), 0
            Line (.xLast, .yLast)-(.xLast + .Size - 1, .yLast + .Size - 1), 0, B
            lngCol = .Color '(128 + Int(Rnd * 255)) * (.Color / 255)
            Line (.X, .Y)-(.X + .Size - 1, .Y + .Size - 1), RGB(.ColorR * .Color / 255, .ColorG * .Color / 255, .ColorB * .Color / 255), B
            
            'PSet (.x, .y), RGB(.Color, .Color, .Color)
            'Circle (.x, .y), .Size, RGB(.Color, .Color, .Color)
            .xLast = .X
            .yLast = .Y
            
            'If .y >= Me.ScaleHeight - .Size Then
            '    .Vy = -.Vy
            'End If
            
            '.Lived >= .LiveFor Or
            If .Color <= 16 Then
                If .Type = 2 Then
                    .Color = 64
                    .Vx = -5 + Rnd * 2 * 5
                    .Vy = -5 '-20 + Rnd * 2 * 20
                    .Type = 3
                ElseIf .Type = 4 Or .Type = 1 Then
                    .Color = 64
                    .Vy = 0
                    .Vx = 0
                    .Type = 3
                Else
                    Randomize Timer
                    .VxMax = .VxMaxNext
                    .VyMax = .VyMaxNext
                    .Vx = .VxNext
                    .Vy = .VyNext
                    .X = .xStart
                    .Y = .yStart
                    .Color = 160 + Rnd * 95
                    .ColorR = 128 + Int(Rnd * 128)
                    .ColorG = 128 + Int(Rnd * 128)
                    .ColorB = 128 + Int(Rnd * 128)
                    .Size = 3
                    .Type = .TypeNext
                End If
            End If
            
        End With
    
    Next lngPiece
    
    DoEvents
    Call Sleep(WAIT)
    
    GoTo Go

End Function


Private Sub Timer1_Timer()
    Dim lngPiece    As Long
    Dim lngX        As Long
    Dim lngY        As Long
    Dim lngType     As Long
    Dim lngSide     As Long
    Dim dblExpY     As Long
    Dim dblVReal    As Double
    Dim dblSin      As Double
    Dim dblCos      As Double
    Dim lngNegative As Long
    
    If Not blnAnimateStart Then Exit Sub
    
    lngType = Int(Rnd * 4) + 1
    If lngType = 4 Then lngType = 1
    If lngType = 3 Then lngType = 2
    
    Select Case lngType
    
        Case 1 ' Anywhere
            lngX = Int(Rnd * Me.ScaleWidth)
            lngY = Int(Rnd * Me.ScaleHeight)
            dblExpY = (Me.ScaleHeight / 2) * Rnd
        Case 2 ' Magma
            lngX = Int(Rnd * Me.ScaleWidth)
            lngY = Me.ScaleHeight
        Case 3 ' Rain
            lngX = Int(Rnd * Me.ScaleWidth)
            lngY = 0
        Case 4 ' Bullets
            lngSide = Me.ScaleWidth * Int((0.5 + Rnd * 1))
            lngX = 0
            lngY = Int(Rnd * Me.ScaleHeight)
    End Select
    
    For lngPiece = 0 To UBound(arrPieces)
    
        With arrPieces(lngPiece)
            .xStart = lngX
            .yStart = lngY
            Select Case lngType
                Case 1 ' Anywhere, explosion
                    .yStart = dblExpY
                    .VxMaxNext = 20
                    .VyMaxNext = 20
                Case 2 ' Magma
                    .VxMaxNext = 0
                    .VyMaxNext = -80
                Case 3 ' Rain
                    '.xStart = Int(Rnd * Me.ScaleWidth)
                    .VxMaxNext = 0.5
                    .VyMaxNext = 20
                Case 4 ' Bullets
                    '.yStart = Rnd * Me.ScaleHeight
                    .xStart = lngSide
                    If .xStart = 0 Then
                        .VxMaxNext = 25
                    Else
                        .VxMaxNext = -25
                    End If
                    .VyMaxNext = 0
            End Select
            .VxNext = -.VxMaxNext + Rnd * 2 * .VxMaxNext
            .VyNext = -.VyMaxNext + Rnd * 2 * .VyMaxNext
            If lngType = 2 Then
                .VyNext = -80 * Rnd
            ElseIf lngType = 1 Then
                dblVReal = Rnd * 20
                If lngPiece Mod 2 = 0 Then
                    dblSin = Rnd
                    dblCos = (1 - dblSin ^ 2) ^ 0.5
                Else
                    dblCos = Rnd
                    dblSin = (1 - dblCos ^ 2) ^ 0.5
                End If
                'lngNegative = 1
                If Rnd < 0.5 Then lngNegative = -1 Else lngNegative = 1
                .VxNext = lngNegative * dblVReal * dblSin
                If Rnd < 0.5 Then lngNegative = -1 Else lngNegative = 1
                .VyNext = lngNegative * dblVReal * dblCos
                'Debug.Print .VxNext ^ 2 + .VyNext ^ 2
            End If
            .TypeNext = lngType
            .SizeNext = 0
        End With
    Next lngPiece

End Sub


Private Sub AlwaysOnTop(FrmID As Form, OnTop As Boolean)
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    If OnTop Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim X As Integer
    
    'Show the mouse cursor again
    'Call ShowCursor(1)
    
    'End the program
    End

End Sub
