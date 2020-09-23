Attribute VB_Name = "BLMMod1"
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum DirecGrad
Horiz
Vertic
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum
Public xx%, yy%, Directory$, Newx!, Newy!, PicIdx%, PicW%, PicH%, q%
Public Const CDCTitle$ = "CD-Cover V1.0"

Public PiMem(3) As Picture, Im As Picture, Re(1) As Boolean

Public z!, R&, G&, B&, Temp$, Faktor!, NW&, NH&, ff%, PrTitle$, Num&, Cidx%, BgIdx%
Public ShX(1, 49), ShY(1, 49), Alpha&(9), ProjectName$, Sc%, ShapeColor&(15), SideW%, Showfont%

'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Public Const AC_SRC_OVER = &H0
Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hDC As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Const LF_FACESIZE = 32
Public Const LOGPIXELSY = 90

'Module-level private variables
Public mobjDevice As Object
Public msngSX1 As Single
Public msngSY1 As Single
Public msngXRatio As Single
Public msngYRatio As Single
Public mlfFont As LOGFONT
Public mintAngle As Integer

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lsngStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lsngPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim LF As LOGFONT, FontName As String, ZeroPos As Long
    CopyMemory LF, ByVal lplf, LenB(LF)
    FontName = StrConv(LF.lfFaceName, vbUnicode)
    ZeroPos = InStr(1, FontName, Chr$(0))
    If ZeroPos > 0 Then FontName = Left$(FontName, ZeroPos - 1)
    CDC2.Combo1.AddItem FontName
    CDC10.List1.AddItem FontName
    EnumFontProc = 1
End Function

   
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R%, G%, B%, R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FC&, T3Dxx%, SM%
On Error Resume Next

'global things
SM = Obj0.ScaleMode 'save scalemode
Obj0.ScaleMode = 3 'pixel
Obj.Borderstyle = 0 'no border
If IsMissing(Style3D) Then Style3D = 0
If Style3D > 4 Then Style3D = 3

'get formcolor
FC = Obj0.BackColor
'in case formcolor = systemcolor --> call the function RealColor
FC = RealColor(FC)
' convert to RGB
R = FC And &HFF
G = Int((FC And &HFF00&) / 256)
B = Int((FC And &HFF0000) / 65536)
'-------------------
If Style3D = 0 Then 'RaiseRaise
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R3 = R1
    R4 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G3 = G1
    G4 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B3 = B1
    B4 = B2
End If
'-------------------
If Style3D = 1 Then 'RaiseInset
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R4 = R1
    R3 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G4 = G1
    G3 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 2 Then 'InsetRaise
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R4 = R1
    R3 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G4 = G1
    G3 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 3 Then 'InsetInset
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R3 = R1
    R4 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G3 = G1
    G4 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B3 = B1
    B4 = B2
End If
If Style3D = 4 Then 'No Border
R1 = R: R2 = R: R3 = R: R4 = R
G1 = G: G2 = G: G3 = G: G4 = G
B1 = B: B2 = B: B3 = B: B4 = B
End If
Bev = Bev + 1
T3Dxx = Bev 'just in case Filled = 1

'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1 'in case T3DF1 (filled)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)

Obj0.ScaleMode = SM 'restore original scalemode
End Function
  
  ' if System Color then translate to 'normal color'
  ' else, do nothing
  Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function

Public Sub MakeDirectories()
On Error Resume Next
MkDir Directory & "\CDC_Pictures"
MkDir Directory & "\CDC_Tubes"
MkDir Directory & "\CDC-Projects"
MkDir Directory & "\CDC-SavedPictures"
End Sub

Public Sub Screensetup()
Alpha(0) = &H1A0000
Alpha(1) = &H330000
Alpha(2) = &H4D0000
Alpha(3) = &H660000
Alpha(4) = &H800000
Alpha(5) = &H990000
Alpha(6) = &HB30000
Alpha(7) = &HCC0000
Alpha(8) = &HE60000
Alpha(9) = &HFF0000
Directory = App.Path
CDC1.Caption = CDCTitle
With CDC1
.Label3 = ""
PicH = 472
PicW = 472
.Pic1(0).Move 309, 44, PicW, PicH
.Pic1(1).Move 309, 44, PicW, PicH
.Pic1(0).Visible = True
.Pic1(1).Visible = False
T3D CDC1, .Pic1(0), 5, T3dInsetRaise
T3D CDC1, .Pic1(1), 5, T3dInsetRaise
T3D CDC1, .Toolbar1, 3, T3dInsetRaise
T3D CDC1, .Label1, 3, T3dInsetRaise
T3D CDC1, .Label4, 3, T3dInsetRaise
.Label2.Top = .Pic1(0).Top
T3D CDC1, .Label2, 5, T3dInsetRaise
T3D CDC1, .Image1, 5, T3dInsetRaise
.Label1 = "FRONTSIDE"
PicIdx = 0
For xx = 1 To 49
Load .Lab0(xx)
Load .SLab0(xx)
Load .Lab1(xx)
Load .SLab1(xx)
Set .Lab0(xx).Container = .Pic1(0)
Set .SLab0(xx).Container = .Pic1(0)
Set .Lab1(xx).Container = .Pic1(1)
Set .SLab1(xx).Container = .Pic1(1)
Next xx
DefaultTextPositions
.Shape1.Visible = False
End With
'---------------------------------------
With CDC10
.Label2 = "number of fonts: " & .List1.ListCount
For xx = 1 To .List1.ListCount - 1
Load .Label1(xx)
.Label1(xx).Visible = True
.Label1(xx).Move 0, xx * .Label1(0).Height
.Label1(xx) = xx
Next xx
.Picture2.Move 0, 0, .Picture1.ScaleWidth - 1, .List1.ListCount * .Label1(0).Height
.VScroll1.Min = 0
.VScroll1.Max = (.List1.ListCount * -1) + 20
For xx = 0 To .List1.ListCount - 1
.Label1(xx) = .List1.List(xx)
.Label1(xx).Font = .List1.List(xx)
Next xx
.Label3 = "abcdefghijklmnopqrstuvwxyz" & vbCrLf & "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & vbCrLf & "0 1 2 3 4 5 6 7 8 9"
.Label3.Font = .List1.List(0)
.Label4 = .List1.List(0)
.Label1(0).BackColor = vbBlue
End With
End Sub

Public Sub Dimention(Ob As Object, Ob2 As Object, cX%, cY%)
Newx = cX
Newy = cY
T = 1.1
Do While Newx > 130 Or Newy > 130
Newx = cX / T
Newy = cY / T
T = T + 0.1
Loop
Ob.Move ((130 - Newx) / 2) + 18, ((130 - Newy) / 2) + 387, Newx, Newy
Ob.Picture = Ob2.Picture
End Sub

Public Sub Dimention2(Ob As Object, Ob2 As Object, cX%, cY%)
Newx = cX
Newy = cY
T = 1.1
Do While Newx > 130 Or Newy > 130
Newx = cX / T
Newy = cY / T
T = T + 0.1
Loop
Ob.Move ((130 - Newx) / 2) + 246, ((130 - Newy) / 2) + 36, Newx, Newy
Ob.Picture = Ob2.Picture
End Sub

Public Sub CopyToMap(Ob As Object, cX%, cY%, ccX%, ccY%)
On Error GoTo CopyToMap1
If PicPresent(CDC1) = False Then Exit Sub
SaveRedo
Ob.PaintPicture CDC1.Pic2.Picture, ccX, ccY, cX, cY
Ob.Refresh
Exit Sub
CopyToMap1:
MsgBox "Invalid picture...", vbCritical, CdTitle
End Sub

Public Function PicPresent(Ob As Object) As Boolean
PicPresent = False
For xx = 0 To Ob.File1.ListCount - 1
If Ob.File1.Selected(xx) = True Then
PicPresent = True
Exit Function
End If
Next xx
MsgBox "No picture selected", vbExclamation, BLMTitle
End Function

Public Sub SetText()
With CDC2
.Label1 = "Edit " & Choose(PicIdx + 1, "Frontside", "Backside")
.List1.Clear
For xx = 0 To 49
If PicIdx = 0 Then .List1.AddItem CDC1.Lab0(xx)
If PicIdx = 1 Then .List1.AddItem CDC1.Lab1(xx)
Next xx
.List1.Selected(0) = True
.Toolbar2.Buttons(PicIdx + 1).Value = tbrPressed
CDC1.Toolbar1.Buttons(PicIdx + 3).Value = tbrPressed
End With
End Sub

Public Sub DefaultTextPositions()
With CDC1
For xx = 0 To 49
.Lab0(xx).Visible = False
.Lab0(xx).Font = "Arial"
.Lab0(xx).FontSize = 12
.Lab0(xx).FontBold = False
.Lab0(xx).FontItalic = False
.Lab0(xx).FontUnderline = False
.Lab0(xx) = "Front" & Format(xx, "00")
.Lab0(xx).ForeColor = &HFFFF00
.Lab0(xx).Left = 10
.Lab0(xx).Top = 20 * xx
ShX(0, xx) = 2
ShY(0, xx) = 2
.SLab0(xx).Visible = False
.SLab0(xx).Font = "Arial"
.SLab0(xx).FontSize = 12
.SLab0(xx).FontBold = False
.SLab0(xx).FontItalic = False
.SLab0(xx).FontUnderline = False
.SLab0(xx) = "Front" & Format(xx, "00")
.SLab0(xx).ForeColor = 0
.SLab0(xx).Left = .Lab0(xx).Left + ShX(0, xx)
.SLab0(xx).Top = .Lab0(xx).Top + ShY(0, xx)
Set .Lab0(xx).Container = .Pic1(0)
Set .SLab0(xx).Container = .Pic1(0)

.Lab1(xx).Visible = False
.Lab1(xx).Font = "Arial"
.Lab1(xx).FontSize = 12
.Lab1(xx).FontBold = False
.Lab1(xx).FontItalic = False
.Lab1(xx).FontUnderline = False
.Lab1(xx) = "Back" & Format(xx, "00")
.Lab1(xx).ForeColor = &HFFFF00
.Lab1(xx).Left = 10
.Lab1(xx).Top = 20 * xx
ShX(1, xx) = 2
ShY(1, xx) = 2
.SLab1(xx).Visible = False
.SLab1(xx).Font = "Arial"
.SLab1(xx).FontSize = 12
.SLab1(xx).FontBold = False
.SLab1(xx).FontItalic = False
.SLab1(xx).FontUnderline = False
.SLab1(xx) = "Back" & Format(xx, "00")
.SLab1(xx).ForeColor = 0
.SLab1(xx).Left = .Lab1(xx).Left + ShX(1, xx)
.SLab1(xx).Top = .Lab1(xx).Top + ShY(1, xx)
Set .Lab1(xx).Container = .Pic1(1)
Set .SLab1(xx).Container = .Pic1(1)

Next xx
ProjectName = "CDC_Project.cdb"
.Label4 = ProjectName
End With
End Sub

Public Sub SaveRedo()
CDC1.TempMem.Picture = CDC1.Pic1(PicIdx).Image
Set PiMem(PicIdx) = CDC1.TempMem.Image
CDC1.Toolbar1.Buttons(1).Enabled = True
Re(PicIdx) = True
End Sub

Public Sub Redo()
CDC1.Pic1(PicIdx) = PiMem(PicIdx)
CDC1.Pic1(PicIdx).Refresh
Set PiMem(PicIdx) = Nothing
CDC1.Toolbar1.Buttons(1).Enabled = False
Re(PicIdx) = False
End Sub

Public Function Grad(Obj As Object, Col1 As Long, Col2 As Long, Optional Dgrad As DirecGrad)
Dim R1, R2, G1, G2, B1, B2, Sr, Sg, Sb, H%, H2%, xxx%
Dim R, G, B
Dim TmpScale%
On Error Resume Next
If IsMissing(Dgrad) Then Dgrad = Horiz
TmpScale = Obj.ScaleMode
Obj.ScaleMode = 3
Obj.AutoRedraw = True
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
If Dgrad = Horiz Then
H = Obj.ScaleHeight
Else
H = Obj.ScaleWidth
End If
Sr = (R2 - R1) / H
Sg = (G2 - G1) / H
Sb = (B2 - B1) / H
For xxx = 0 To H
If Dgrad = Horiz Then
Obj.Line (0, xxx)-(Obj.ScaleWidth, xxx), RGB(R1, G1, B1)
Else
Obj.Line (xxx, 0)-(xxx, Obj.ScaleHeight), RGB(R1, G1, B1)
End If
R1 = R1 + Sr
G1 = G1 + Sg
B1 = B1 + Sb
Next xxx
Obj.ScaleMode = TmpScale
End Function

Public Function Grad2(Obj As Object, Col1 As Long, Col2 As Long, Optional Dgrad As DirecGrad)
Dim R1, R2, G1, G2, B1, B2, Sr, Sg, Sb, H%, H2%, xxx%
Dim R, G, B
Dim TmpScale%
On Error Resume Next
If IsMissing(Dgrad) Then Dgrad = Horiz
TmpScale = Obj.ScaleMode
Obj.ScaleMode = 3
Obj.AutoRedraw = True
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
If Dgrad = Horiz Then
H = Obj.ScaleHeight / 2
Else
H = Obj.ScaleWidth / 2
End If
Sr = (R2 - R1) / H
Sg = (G2 - G1) / H
Sb = (B2 - B1) / H
For xxx = 0 To H
If Dgrad = Horiz Then
Obj.Line (0, xxx)-(Obj.ScaleWidth, xxx), RGB(R1, G1, B1)
Obj.Line (0, Obj.ScaleHeight - xxx)-(Obj.ScaleWidth, Obj.ScaleHeight - xxx), RGB(R1, G1, B1)
Else
Obj.Line (xxx, 0)-(xxx, Obj.ScaleHeight), RGB(R1, G1, B1)
Obj.Line (Obj.ScaleWidth - xxx, 0)-(Obj.ScaleWidth - xxx, Obj.ScaleHeight), RGB(R1, G1, B1)
End If
R1 = R1 + Sr
G1 = G1 + Sg
B1 = B1 + Sb
Next xxx
Obj.ScaleMode = TmpScale
End Function

Public Function Grad3(Obj As Object, Col1 As Long, Col2 As Long, Col3 As Long, Optional Dgrad As DirecGrad)
Dim R1, R2, R3, G1, G2, G3, B1, B2, B3, Sr, Sg, Sb, H%, H2%, xxx%
Dim R, G, B
Dim TmpScale%
On Error Resume Next
If IsMissing(Dgrad) Then Dgrad = Horiz
TmpScale = Obj.ScaleMode
Obj.ScaleMode = 3
Obj.AutoRedraw = True
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
R3 = Col3 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
G3 = (Col3 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
B3 = (Col3 And &H80FF0000) / &H10000
    If Dgrad = Horiz Then
    H = Obj.ScaleHeight / 2
    H2 = Obj.ScaleHeight
    Else
    H = Obj.ScaleWidth / 2
    H2 = Obj.ScaleWidth
    End If
    
    Sr = (R2 - R1) / H
    Sg = (G2 - G1) / H
    Sb = (B2 - B1) / H
    For xxx = 0 To H
    If Dgrad = Horiz Then
    Obj.Line (0, xxx)-(Obj.ScaleWidth, xxx), RGB(R1, G1, B1)
    Else
    Obj.Line (xxx, 0)-(xxx, Obj.ScaleHeight), RGB(R1, G1, B1)
    End If
    R1 = R1 + Sr
    G1 = G1 + Sg
    B1 = B1 + Sb
    Next xxx
    Sr = (R3 - R2) / H
    Sg = (G3 - G2) / H
    Sb = (B3 - B2) / H
    For xxx = H To H2
    If Dgrad = Horiz Then
    Obj.Line (0, xxx)-(Obj.ScaleWidth, xxx), RGB(R2, G2, B2)
    Else
    Obj.Line (xxx, 0)-(xxx, Obj.ScaleHeight), RGB(R2, G2, B2)
    End If
    R2 = R2 + Sr
    G2 = G2 + Sg
    B2 = B2 + Sb
    Next xxx
Obj.ScaleMode = TmpScale
End Function

Public Function CircleGradient(Obj As Object, StX%, StY%, Col1&, Col2&, cR%, Asp!)
Screen.MousePointer = 11
Dim CstepR As Single, CstepG As Single, CstepB As Single, cX%
Dim R1!, R2!, G1!, G2!, B1!, B2!
Obj.AutoRedraw = True
Obj.ScaleMode = 3
Obj.DrawWidth = 2
Obj.DrawStyle = 6 'inside solid
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
For cX = 0 To Obj.Height
Obj.Line (0, cX)-(Obj.Width, cX), RGB(R2, G2, B2)
Next cX
CstepR = (R2 - R1) / cR
CstepG = (G2 - G1) / cR
CstepB = (B2 - B1) / cR
For cX = 0 To cR
Obj.Circle (StX, StY), cX, RGB(R1, G1, B1), , , Asp
R1 = R1 + CstepR
G1 = G1 + CstepG
B1 = B1 + CstepB
Next cX
Obj.DrawWidth = 1
Screen.MousePointer = 1
End Function

Public Function CircleGradient2(Obj As Object, StX%, StY%, Col1&, Col2&, cR%, Asp!)
Screen.MousePointer = 11
Dim CstepR As Single, CstepG As Single, CstepB As Single, cX%
Dim R1!, R2!, G1!, G2!, B1!, B2!
Obj.AutoRedraw = True
Obj.ScaleMode = 3
Obj.DrawWidth = 2
Obj.DrawStyle = 6 'inside solid
R1 = Col1 And &H800000FF
R2 = Col2 And &H800000FF
G1 = (Col1 And &H8000FF00) / &H100
G2 = (Col2 And &H8000FF00) / &H100
B1 = (Col1 And &H80FF0000) / &H10000
B2 = (Col2 And &H80FF0000) / &H10000
For cX = 0 To Obj.Height
Obj.Line (0, cX)-(Obj.Width, cX), RGB(R1, G1, B1)
Next cX
CstepR = (R2 - R1) / cR * 2
CstepG = (G2 - G1) / cR * 2
CstepB = (B2 - B1) / cR * 2
For cX = 0 To cR / 2
Obj.Circle (StX, StY), cX, RGB(R1, G1, B1), , , Asp
Obj.Circle (StX, StY), cR - cX, RGB(R1, G1, B1), , , Asp
R1 = R1 + CstepR
G1 = G1 + CstepG
B1 = B1 + CstepB
Next cX
Obj.DrawWidth = 1
Screen.MousePointer = 1
End Function


