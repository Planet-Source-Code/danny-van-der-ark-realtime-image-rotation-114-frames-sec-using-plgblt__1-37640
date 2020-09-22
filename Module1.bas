Attribute VB_Name = "Module1"
Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, _
                                    lpPoint As POINTS2D, _
                                    ByVal hdcSrc As Long, _
                                    ByVal nXSrc As Long, _
                                    ByVal nYSrc As Long, _
                                    ByVal nWidth As Long, _
                                    ByVal nHeight As Long, _
                                    ByVal hbmMask As Long, _
                                    ByVal xMask As Long, _
                                    ByVal yMask As Long) As Long

Global Const NotPI = 3.14159265238 / 180

'--------------------------------------------------------------------------------
Public Type POINTS2D
   x As Long
   y As Long
End Type


