Attribute VB_Name = "BLMMod3"
      #If Win32 Then
      Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
      Private iBKMode As Long
   #Else
      Private Declare Function SetBkMode Lib "GDI" (ByVal hDC As Integer, ByVal nBkMode As Integer) As Integer
      Private iBKMode As Integer
   #End If
      Private Const TRANSPARENT = 1
      Private Const OPAQUE = 2
Public PrtposX%(2), PrtposY%(2), PrtPosX2%, PrtPosY2%
Public CX0%, CY0%, CX1%, CY1%, CX2%, CY2%, CX3%, CY3%

Public Sub PrintFront()
CDC1.Pic1(0).Picture = CDC1.Pic1(0).Image
CDC1.Pic1(1).Picture = CDC1.Pic1(1).Image
CY1 = CDC1.Pic1(1).ScaleHeight * Screen.TwipsPerPixelY
CX2 = (CDC1.Pic1(1).ScaleWidth + CDC1.Pic1(0).ScaleWidth) * Screen.TwipsPerPixelX
CX3 = CDC1.Pic1(0).ScaleWidth * Screen.TwipsPerPixelX
CDC1.ScaleMode = 1
For xx = 0 To 1
CDC1.Pic1(xx).ScaleMode = 1 'set to twips
Next xx
 Printer.ScaleMode = 1
Printer.Orientation = 2
PrtPosX2 = PrtposX(0)
PrtPosY2 = PrtposY(0)
Printer.Line (PrtPosX2 - (20 * Screen.TwipsPerPixelX), PrtPosY2)-(PrtPosX2 + CX2 + (40 * Screen.TwipsPerPixelX), PrtPosY2), 0
Printer.Line (PrtPosX2 - (20 * Screen.TwipsPerPixelX), PrtPosY2 + CY1)-(PrtPosX2 + CX2 + (40 * Screen.TwipsPerPixelX), PrtPosY2 + CY1), 0
Printer.Line (PrtPosX2, PrtPosY2 - (20 * Screen.TwipsPerPixelY))-(PrtPosX2, PrtPosY2 + CY1 + (20 * Screen.TwipsPerPixelY)), 0
Printer.Line (PrtPosX2 + CX2, PrtPosY2 - (20 * Screen.TwipsPerPixelY))-(PrtPosX2 + CX2, PrtPosY2 + CY1 + (20 * Screen.TwipsPerPixelY)), 0
'middenlijn
Printer.Line (PrtPosX2 + CX3, PrtPosY2 - (20 * Screen.TwipsPerPixelY))-(PrtPosX2 + CX3, PrtPosY2 + CY1 + (20 * Screen.TwipsPerPixelY)), 0
'zijkantjes
Printer.Line (PrtPosX2 + CX3, PrtPosY2 - (20 * Screen.TwipsPerPixelY))-(PrtPosX2 + CX2, PrtPosY2 - (20 * Screen.TwipsPerPixelY)), 0
Printer.Line (PrtPosX2 + CX3, PrtPosY2 + CY1 + (20 * Screen.TwipsPerPixelY))-(PrtPosX2 + CX2, PrtPosY2 + CY1 + (20 * Screen.TwipsPerPixelY)), 0
Printer.Line (PrtPosX2 + CX2 + (40 * Screen.TwipsPerPixelX), PrtPosY2)-(PrtPosX2 + CX2 + (40 * Screen.TwipsPerPixelX), PrtPosY2 + CY1), 0

Printer.PaintPicture CDC1.Pic1(1).Picture, PrtPosX2, PrtposY(0)
       Printer.FontTransparent = True
       iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
        PrintBackTxt
PrtPosX2 = (PicW * Screen.TwipsPerPixelX) + PrtposX(0)
PrtPosY2 = PrtposY(0)
Printer.PaintPicture CDC1.Pic1(0).Picture, PrtPosX2, PrtposY(0)
       Printer.FontTransparent = True
       iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
        PrintFrontTxt
Printer.EndDoc
Printer.Orientation = 1
CDC1.ScaleMode = 3
For xx = 0 To 1
CDC1.Pic1(xx).ScaleMode = 3 'set to pixel
Next xx
End Sub

Public Sub PrintBackTxt()
With CDC1
For xx = 0 To 49
If .SLab1(xx).Visible = True Then
    Printer.Font = .SLab1(xx).Font
    Printer.FontSize = .SLab1(xx).FontSize
    Printer.FontBold = .SLab1(xx).FontBold
    Printer.FontItalic = .SLab1(xx).FontItalic
    Printer.FontUnderline = .SLab1(xx).FontUnderline
    Printer.ForeColor = .SLab1(xx).ForeColor
    Printer.CurrentX = .SLab1(xx).Left + PrtPosX2
    Printer.CurrentY = .SLab1(xx).Top + PrtposY(0)
    Printer.Print .SLab1(xx).Caption
End If
If .Lab1(xx).Visible = True Then
    Printer.Font = .Lab1(xx).Font
    Printer.FontSize = .Lab1(xx).FontSize
    Printer.FontBold = .Lab1(xx).FontBold
    Printer.FontItalic = .Lab1(xx).FontItalic
    Printer.FontUnderline = .Lab1(xx).FontUnderline
    Printer.ForeColor = .Lab1(xx).ForeColor
    Printer.CurrentX = .Lab1(xx).Left + PrtPosX2
    Printer.CurrentY = .Lab1(xx).Top + PrtposY(0)
    Printer.Print .Lab1(xx).Caption
End If
Next xx
End With
End Sub

Public Sub PrintFrontTxt()
With CDC1
For xx = 0 To 49
If .SLab0(xx).Visible = True Then
    Printer.Font = .SLab0(xx).Font
    Printer.FontSize = .SLab0(xx).FontSize
    Printer.FontBold = .SLab0(xx).FontBold
    Printer.FontItalic = .SLab0(xx).FontItalic
    Printer.FontUnderline = .SLab0(xx).FontUnderline
    Printer.ForeColor = .SLab0(xx).ForeColor
    Printer.CurrentX = .SLab0(xx).Left + PrtPosX2
    Printer.CurrentY = .SLab0(xx).Top + PrtposY(0)
    Printer.Print .SLab0(xx).Caption
End If
If .Lab0(xx).Visible = True Then
    Printer.Font = .Lab0(xx).Font
    Printer.FontSize = .Lab0(xx).FontSize
    Printer.FontBold = .Lab0(xx).FontBold
    Printer.FontItalic = .Lab0(xx).FontItalic
    Printer.FontUnderline = .Lab0(xx).FontUnderline
    Printer.ForeColor = .Lab0(xx).ForeColor
    Printer.CurrentX = .Lab0(xx).Left + PrtPosX2
    Printer.CurrentY = .Lab0(xx).Top + PrtposY(0)
    Printer.Print .Lab0(xx).Caption
End If
Next xx
End With
End Sub

