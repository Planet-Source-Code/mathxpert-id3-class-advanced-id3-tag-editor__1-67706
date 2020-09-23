Attribute VB_Name = "modHimetrics"
Option Explicit

Public Function HimetricToPixelsX(ByVal Himetrics As Long) As Long
    HimetricToPixelsX = Fix(CDbl(Himetrics) / 2540 * 1440 / CDbl(Screen.TwipsPerPixelX))
End Function

Public Function HimetricToPixelsY(ByVal Himetrics As Long) As Long
    HimetricToPixelsY = Fix(CDbl(Himetrics) / 2540 * 1440 / CDbl(Screen.TwipsPerPixelY))
End Function
