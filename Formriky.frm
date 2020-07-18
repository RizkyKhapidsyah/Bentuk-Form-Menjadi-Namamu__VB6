VERSION 5.00
Begin VB.Form rikyMetallll 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   8640
   ClientLeft      =   2250
   ClientTop       =   1680
   ClientWidth     =   14685
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   14685
End
Attribute VB_Name = "rikyMetallll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "Rizky Khapidsyah"
Private Const RegisteredTo = "Indonesian IT Programmer"
Private ResultRegion As Long
Private Function RubahBentukForm(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, LakukanAksii As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

    ObjectRegion = CreateRectRgn(133 * ScaleX * 15 / STPPX + OffsetX, 151 * ScaleY * 15 / STPPY + OffsetY, 180 * ScaleX * 15 / STPPX + OffsetX, 408 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion

    ObjectRegion = CreateRectRgn(133 * ScaleX * 15 / STPPX + OffsetX, 151 * ScaleY * 15 / STPPY + OffsetY, 325 * ScaleX * 15 / STPPX + OffsetX, 190 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(302 * ScaleX * 15 / STPPX + OffsetX, 155 * ScaleY * 15 / STPPY + OffsetY, 333 * ScaleX * 15 / STPPX + OffsetX, 264 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(158 * ScaleX * 15 / STPPX + OffsetX, 249 * ScaleY * 15 / STPPY + OffsetY, 328 * ScaleX * 15 / STPPX + OffsetX, 281 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 4)
    For Counter = 0 To 4
        PolyPoints(Counter).X = GP0X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP0Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 5, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(351 * ScaleX * 15 / STPPX + OffsetX, 259 * ScaleY * 15 / STPPY + OffsetY, 388 * ScaleX * 15 / STPPX + OffsetX, 391 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(410 * ScaleX * 15 / STPPX + OffsetX, 260 * ScaleY * 15 / STPPY + OffsetY, 529 * ScaleX * 15 / STPPX + OffsetX, 285 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(415 * ScaleX * 15 / STPPX + OffsetX, 371 * ScaleY * 15 / STPPY + OffsetY, 523 * ScaleX * 15 / STPPX + OffsetX, 395 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 0)
    For Counter = 0 To 0
        PolyPoints(Counter).X = GP1X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP1Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 1, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 4)
    For Counter = 0 To 4
        PolyPoints(Counter).X = GP2X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP2Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 5, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(543 * ScaleX * 15 / STPPX + OffsetX, 261 * ScaleY * 15 / STPPY + OffsetY, 574 * ScaleX * 15 / STPPX + OffsetX, 401 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 7)
    For Counter = 0 To 7
        PolyPoints(Counter).X = GP3X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP3Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 8, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 8)
    For Counter = 0 To 8
        PolyPoints(Counter).X = GP4X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP4Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 9, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 1)
    For Counter = 0 To 1
        PolyPoints(Counter).X = GP5X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP5Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 2, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(316 * ScaleX * 15 / STPPX + OffsetX, 439 * ScaleY * 15 / STPPY + OffsetY, 338 * ScaleX * 15 / STPPX + OffsetX, 516 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 6)
    For Counter = 0 To 6
        PolyPoints(Counter).X = GP6X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP6Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 7, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(391 * ScaleX * 15 / STPPX + OffsetX, 444 * ScaleY * 15 / STPPY + OffsetY, 404 * ScaleX * 15 / STPPX + OffsetX, 511 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(401 * ScaleX * 15 / STPPX + OffsetX, 474 * ScaleY * 15 / STPPY + OffsetY, 427 * ScaleX * 15 / STPPX + OffsetX, 485 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(424 * ScaleX * 15 / STPPX + OffsetX, 445 * ScaleY * 15 / STPPY + OffsetY, 436 * ScaleX * 15 / STPPX + OffsetX, 510 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 8)
    For Counter = 0 To 8
        PolyPoints(Counter).X = GP7X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP7Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 9, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 2)
    For Counter = 0 To 2
        PolyPoints(Counter).X = GP8X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP8Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 3, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 4)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(523 * ScaleX * 15 / STPPX + OffsetX, 448 * ScaleY * 15 / STPPY + OffsetY, 538 * ScaleX * 15 / STPPX + OffsetX, 508 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 4)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(526 * ScaleX * 15 / STPPX + OffsetX, 449 * ScaleY * 15 / STPPY + OffsetY, 538 * ScaleX * 15 / STPPX + OffsetX, 510 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(528 * ScaleX * 15 / STPPX + OffsetX, 448 * ScaleY * 15 / STPPY + OffsetY, 568 * ScaleX * 15 / STPPX + OffsetX, 458 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(534 * ScaleX * 15 / STPPX + OffsetX, 469 * ScaleY * 15 / STPPY + OffsetY, 565 * ScaleX * 15 / STPPX + OffsetX, 478 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(581 * ScaleX * 15 / STPPX + OffsetX, 452 * ScaleY * 15 / STPPY + OffsetY, 591 * ScaleX * 15 / STPPX + OffsetX, 517 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(600 * ScaleX * 15 / STPPX + OffsetX, 450 * ScaleY * 15 / STPPY + OffsetY, 649 * ScaleX * 15 / STPPX + OffsetX, 462 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(620 * ScaleX * 15 / STPPX + OffsetX, 459 * ScaleY * 15 / STPPY + OffsetY, 631 * ScaleX * 15 / STPPX + OffsetX, 510 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(674 * ScaleX * 15 / STPPX + OffsetX, 452 * ScaleY * 15 / STPPY + OffsetY, 726 * ScaleX * 15 / STPPX + OffsetX, 466 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(678 * ScaleX * 15 / STPPX + OffsetX, 463 * ScaleY * 15 / STPPY + OffsetY, 687 * ScaleX * 15 / STPPX + OffsetX, 491 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(681 * ScaleX * 15 / STPPX + OffsetX, 485 * ScaleY * 15 / STPPY + OffsetY, 724 * ScaleX * 15 / STPPX + OffsetX, 497 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(713 * ScaleX * 15 / STPPX + OffsetX, 494 * ScaleY * 15 / STPPY + OffsetY, 725 * ScaleX * 15 / STPPX + OffsetX, 520 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(676 * ScaleX * 15 / STPPX + OffsetX, 514 * ScaleY * 15 / STPPY + OffsetY, 716 * ScaleX * 15 / STPPX + OffsetX, 522 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 8)
    For Counter = 0 To 8
        PolyPoints(Counter).X = GP9X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP9Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 9, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 8)
    For Counter = 0 To 8
        PolyPoints(Counter).X = GP10X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP10Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 9, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 2)
    For Counter = 0 To 2
        PolyPoints(Counter).X = GP11X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).y = GP11Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 3, 1)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 4)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(910 * ScaleX * 15 / STPPX + OffsetX, 432 * ScaleY * 15 / STPPY + OffsetY, 926 * ScaleX * 15 / STPPX + OffsetX, 517 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(923 * ScaleX * 15 / STPPX + OffsetX, 459 * ScaleY * 15 / STPPY + OffsetY, 966 * ScaleX * 15 / STPPX + OffsetX, 476 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(958 * ScaleX * 15 / STPPX + OffsetX, 432 * ScaleY * 15 / STPPY + OffsetY, 979 * ScaleX * 15 / STPPX + OffsetX, 520 * ScaleY * 15 / STPPY + OffsetY)
    LakukanAksii = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    LakukanAksii = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    RubahBentukForm = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 183
    Case 1
        GP0X = 305
    Case 2
        GP0X = 332
    Case 3
        GP0X = 228
    Case 4
        GP0X = 184
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 276
    Case 1
        GP0Y = 386
    Case 2
        GP0Y = 360
    Case 3
        GP0Y = 276
    Case 4
        GP0Y = 273
    End Select
End Function
Private Function GP1X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP1X = 540
    End Select
End Function
Private Function GP1Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP1Y = 259
    End Select
End Function
Private Function GP2X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP2X = 516
    Case 1
        GP2X = 413
    Case 2
        GP2X = 453
    Case 3
        GP2X = 525
    Case 4
        GP2X = 526
    End Select
End Function
Private Function GP2Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP2Y = 265
    Case 1
        GP2Y = 373
    Case 2
        GP2Y = 373
    Case 3
        GP2Y = 300
    Case 4
        GP2Y = 270
    End Select
End Function
Private Function GP3X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP3X = 564
    Case 1
        GP3X = 625
    Case 2
        GP3X = 641
    Case 3
        GP3X = 576
    Case 4
        GP3X = 646
    Case 5
        GP3X = 625
    Case 6
        GP3X = 570
    Case 7
        GP3X = 562
    End Select
End Function
Private Function GP3Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP3Y = 307
    Case 1
        GP3Y = 253
    Case 2
        GP3Y = 272
    Case 3
        GP3Y = 326
    Case 4
        GP3Y = 395
    Case 5
        GP3Y = 408
    Case 6
        GP3Y = 354
    Case 7
        GP3Y = 308
    End Select
End Function
Private Function GP4X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP4X = 661
    Case 1
        GP4X = 692
    Case 2
        GP4X = 726
    Case 3
        GP4X = 748
    Case 4
        GP4X = 777
    Case 5
        GP4X = 723
    Case 6
        GP4X = 700
    Case 7
        GP4X = 723
    Case 8
        GP4X = 661
    End Select
End Function
Private Function GP4Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP4Y = 260
    Case 1
        GP4Y = 258
    Case 2
        GP4Y = 317
    Case 3
        GP4Y = 255
    Case 4
        GP4Y = 259
    Case 5
        GP4Y = 416
    Case 6
        GP4Y = 416
    Case 7
        GP4Y = 343
    Case 8
        GP4Y = 262
    End Select
End Function
Private Function GP5X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP5X = 405
    Case 1
        GP5X = 401
    End Select
End Function
Private Function GP5Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP5Y = 463
    Case 1
        GP5Y = 576
    End Select
End Function
Private Function GP6X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP6X = 328
    Case 1
        GP6X = 359
    Case 2
        GP6X = 373
    Case 3
        GP6X = 337
    Case 4
        GP6X = 383
    Case 5
        GP6X = 366
    Case 6
        GP6X = 328
    End Select
End Function
Private Function GP6Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP6Y = 467
    Case 1
        GP6Y = 443
    Case 2
        GP6Y = 443
    Case 3
        GP6Y = 477
    Case 4
        GP6Y = 512
    Case 5
        GP6Y = 512
    Case 6
        GP6Y = 481
    End Select
End Function
Private Function GP7X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP7X = 449
    Case 1
        GP7X = 479
    Case 2
        GP7X = 513
    Case 3
        GP7X = 493
    Case 4
        GP7X = 487
    Case 5
        GP7X = 473
    Case 6
        GP7X = 466
    Case 7
        GP7X = 451
    Case 8
        GP7X = 472
    End Select
End Function
Private Function GP7Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP7Y = 510
    Case 1
        GP7Y = 448
    Case 2
        GP7Y = 511
    Case 3
        GP7Y = 508
    Case 4
        GP7Y = 495
    Case 5
        GP7Y = 497
    Case 6
        GP7Y = 512
    Case 7
        GP7Y = 511
    Case 8
        GP7Y = 464
    End Select
End Function
Private Function GP8X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP8X = 476
    Case 1
        GP8X = 484
    Case 2
        GP8X = 475
    End Select
End Function
Private Function GP8Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP8Y = 469
    Case 1
        GP8Y = 479
    Case 2
        GP8Y = 479
    End Select
End Function
Private Function GP9X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP9X = 741
    Case 1
        GP9X = 765
    Case 2
        GP9X = 780
    Case 3
        GP9X = 792
    Case 4
        GP9X = 814
    Case 5
        GP9X = 778
    Case 6
        GP9X = 771
    Case 7
        GP9X = 780
    Case 8
        GP9X = 748
    End Select
End Function
Private Function GP9Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP9Y = 452
    Case 1
        GP9Y = 449
    Case 2
        GP9Y = 486
    Case 3
        GP9Y = 449
    Case 4
        GP9Y = 452
    Case 5
        GP9Y = 538
    Case 6
        GP9Y = 530
    Case 7
        GP9Y = 500
    Case 8
        GP9Y = 453
    End Select
End Function
Private Function GP10X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP10X = 846
    Case 1
        GP10X = 811
    Case 2
        GP10X = 833
    Case 3
        GP10X = 842
    Case 4
        GP10X = 865
    Case 5
        GP10X = 880
    Case 6
        GP10X = 899
    Case 7
        GP10X = 861
    Case 8
        GP10X = 848
    End Select
End Function
Private Function GP10Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP10Y = 445
    Case 1
        GP10Y = 526
    Case 2
        GP10Y = 526
    Case 3
        GP10Y = 502
    Case 4
        GP10Y = 502
    Case 5
        GP10Y = 534
    Case 6
        GP10Y = 534
    Case 7
        GP10Y = 419
    Case 8
        GP10Y = 443
    End Select
End Function
Private Function GP11X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP11X = 858
    Case 1
        GP11X = 849
    Case 2
        GP11X = 864
    End Select
End Function
Private Function GP11Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP11Y = 454
    Case 1
        GP11Y = 480
    Case 2
        GP11Y = 481
    End Select
End Function

Private Sub Form_Load()
    Dim LakukanAksii As Long
    LakukanAksii = SetWindowRgn(Me.hWnd, RubahBentukForm(1, 1, 0, 0), True)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub
