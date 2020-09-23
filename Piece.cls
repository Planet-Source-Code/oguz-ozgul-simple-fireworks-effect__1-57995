VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Piece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type aPiece
    x               As Long ' x position on surface
    y               As Long ' y position on surface
    Vx              As Long ' move after Vx milliseconds
    VY              As Long ' move after Vy milliseconds
    WaitX           As Long ' steps to move horizontally
    waity           As Long ' steps to move vertically
    color           As Long ' color of piece
    xLast           As Long ' Last Displayed X Pos
    yLast           As Long ' Last Displayed y Pos
End Type

Private typPiece()  As aPiece

Private frmSurface  As Form




Public Function Initialize(ByVal x As Long, _
                           ByVal y As Long, _
                           ByVal Vx As Long, _
                           ByVal VY As Long, _
                           ByVal color As Long, _
                           ByVal FadeToBlack As Boolean, _
                           ByVal NumberOfTracks As Long, _
                           ByVal TracksWait As Long)
                           
    Dim lngPiece As Long
    
    ReDim typPiece(NumberOfTracks)
    
    With typPiece(0)
        .x = x
        .y = y
        .Vx = Vx
        .VY = VY
        .WaitX = Vx
        .waity = VY
        .color = color
        .xLast = -1
        .yLast = -1
    End With
    
    For lngPiece = 1 To NumberOfTracks
        With typPiece(lngPiece)
            
        
        End With
    Next lngPiece
                       
                       

End Function




Public Property Let Surface(ByVal frmForm As Form)



End Property
