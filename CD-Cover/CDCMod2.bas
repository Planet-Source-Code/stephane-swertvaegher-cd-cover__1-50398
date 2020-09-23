Attribute VB_Name = "BLMMod2"
'Public Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Public Declare Function SetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Public Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Const ILLUMINANT_A = 1
Const HALFTONE = 4
Private Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0&

Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup
Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure

  Dim hdcNew As Long
  Dim ret As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim X As Long, Y As Long, Tx%
  Dim R As Long, G As Long, B As Long
Public EchoX%, EchoY%, EchoNr%, EchoRed%
Private Speed(0 To 765) As Long   'Speed up values
  Dim sF As Single

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type BITMAPINFOHEADER   '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Public Sub GetPicData(Ob As Object)
Ob.Picture = Ob.Image
 GetObject Ob, Len(PicInfo), PicInfo
  hdcNew = CreateCompatibleDC(0&)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  'redimension  (BGR+pad,x,y)
  ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  'get bytes
  ret = GetDIBits(hdcNew, Ob, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  ret = GetDIBits(hdcNew, Ob, 0, PicInfo.bmHeight, bDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
End Sub

Public Sub SetPicData(Ob As Object)
  'copy bytes to device
  ret = SetDIBits(hdcNew, Ob, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  DeleteDC hdcNew
  ReDim iDATA(1 To 4, 1 To 2, 1 To 2) As Byte
  ReDim bDATA(1 To 4, 1 To 2, 1 To 2) As Byte
End Sub

Public Sub KillComp(C%)
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
     iDATA(C, X, Y) = 0
    Next X, Y
End Sub

Public Sub NegativeImage(cc%)
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    If cc = 0 Or cc = 1 Then iDATA(1, X, Y) = 255 - iDATA(1, X, Y)
    If cc = 0 Or cc = 2 Then iDATA(2, X, Y) = 255 - iDATA(2, X, Y)
    If cc = 0 Or cc = 3 Then iDATA(3, X, Y) = 255 - iDATA(3, X, Y)
    Next X, Y
End Sub

Public Sub RBG()
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
iDATA(1, X, Y) = bDATA(2, X, Y)
iDATA(2, X, Y) = bDATA(1, X, Y)
    Next X, Y
End Sub

Public Sub GRB()
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
iDATA(3, X, Y) = bDATA(2, X, Y)
iDATA(2, X, Y) = bDATA(3, X, Y)
    Next X, Y
End Sub

Public Sub GBR()
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
iDATA(3, X, Y) = bDATA(2, X, Y)
iDATA(2, X, Y) = bDATA(1, X, Y)
iDATA(1, X, Y) = bDATA(3, X, Y)
    Next X, Y
End Sub

Public Sub BRG()
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
iDATA(3, X, Y) = bDATA(1, X, Y)
iDATA(2, X, Y) = bDATA(3, X, Y)
iDATA(1, X, Y) = bDATA(2, X, Y)
    Next X, Y
End Sub

Public Sub BGR()
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
iDATA(3, X, Y) = bDATA(1, X, Y)
iDATA(1, X, Y) = bDATA(3, X, Y)
    Next X, Y
End Sub

Public Sub GreyScale()
For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      R = (iDATA(3, X, Y) * 0.3) + (iDATA(2, X, Y) * 0.59) + (iDATA(1, X, Y) + 0.11)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      iDATA(1, X, Y) = R
      iDATA(2, X, Y) = R
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Emboss(V1%, V2%, V3%)
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      B = Abs(CLng(iDATA(1, X, Y)) - CLng(iDATA(1, X + 1, Y + 1)) + V1)
      G = Abs(CLng(iDATA(2, X, Y)) - CLng(iDATA(2, X + 1, Y + 1)) + V2)
      R = Abs(CLng(iDATA(3, X, Y)) - CLng(iDATA(3, X + 1, Y + 1)) + V3)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub EmbossHR()
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      B = Abs(CLng(iDATA(1, X, Y)) - CLng(iDATA(1, X + 1, Y + 1)) + 128)
      G = Abs(CLng(iDATA(2, X, Y)) - CLng(iDATA(2, X + 1, Y + 1)) + 128)
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
    Next X, Y
End Sub

Public Sub EmbossHG()
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      B = Abs(CLng(iDATA(1, X, Y)) - CLng(iDATA(1, X + 1, Y + 1)) + 128)
      R = Abs(CLng(iDATA(3, X, Y)) - CLng(iDATA(3, X + 1, Y + 1)) + 128)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub EmbossHB()
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      G = Abs(CLng(iDATA(2, X, Y)) - CLng(iDATA(2, X + 1, Y + 1)) + 128)
      R = Abs(CLng(iDATA(3, X, Y)) - CLng(iDATA(3, X + 1, Y + 1)) + 128)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Increase(cc%)
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      R = iDATA(3, X, Y)
      G = iDATA(2, X, Y)
      B = iDATA(1, X, Y)
      If cc = 1 Then R = R + 32
      If cc = 2 Then G = G + 32
      If cc = 3 Then B = B + 32
        If cc = 4 Then
        R = R + 32
        G = G + 32
        B = B + 32
        End If
      If R > 255 Then R = 255
      If G > 255 Then G = 255
      If B > 255 Then B = 255
      iDATA(3, X, Y) = R
      iDATA(2, X, Y) = G
      iDATA(1, X, Y) = B
    Next X, Y
End Sub

Public Sub Decrease(cc%)
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      R = iDATA(3, X, Y)
      G = iDATA(2, X, Y)
      B = iDATA(1, X, Y)
      If cc = 1 Then R = R - 32
      If cc = 2 Then G = G - 32
      If cc = 3 Then B = B - 32
        If cc = 4 Then
        R = R - 32
        G = G - 32
        B = B - 32
        End If
      If R < 0 Then R = 0
      If G < 0 Then G = 0
      If B < 0 Then B = 0
      iDATA(3, X, Y) = R
      iDATA(2, X, Y) = G
      iDATA(1, X, Y) = B
    Next X, Y
End Sub

Public Sub Engrave()
  For Y = 1 To PicInfo.bmHeight - 1
    For X = 1 To PicInfo.bmWidth - 1
      B = Abs(CLng(iDATA(1, X + 1, Y + 1)) - CLng(iDATA(1, X, Y)) + 128)
      G = Abs(CLng(iDATA(2, X + 1, Y + 1)) - CLng(iDATA(2, X, Y)) + 128)
      R = Abs(CLng(iDATA(3, X + 1, Y + 1)) - CLng(iDATA(3, X, Y)) + 128)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub EngraveMore()
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      B = CLng(bDATA(1, X + 1, Y - 1)) - CLng(bDATA(1, X - 1, Y - 1)) + _
          CLng(bDATA(1, X + 1, Y)) - CLng(bDATA(1, X - 1, Y)) + _
          CLng(bDATA(1, X + 1, Y + 1)) - CLng(bDATA(1, X - 1, Y + 1)) + 128
      G = CLng(bDATA(2, X + 1, Y - 1)) - CLng(bDATA(2, X - 1, Y - 1)) + _
          CLng(bDATA(2, X + 1, Y)) - CLng(bDATA(2, X - 1, Y)) + _
          CLng(bDATA(2, X + 1, Y + 1)) - CLng(bDATA(2, X - 1, Y + 1)) + 128
      R = CLng(bDATA(3, X + 1, Y - 1)) - CLng(bDATA(3, X - 1, Y - 1)) + _
          CLng(bDATA(3, X + 1, Y)) - CLng(bDATA(3, X - 1, Y)) + _
          CLng(bDATA(3, X + 1, Y + 1)) - CLng(bDATA(3, X - 1, Y + 1)) + 128
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Diffuse(Factor%)
  Dim aX As Long, aY As Long
  Dim hF As Long
  hF = Factor / 2
For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      aX = Rnd * Factor - hF
      aY = Rnd * Factor - hF
      If X + aX < 1 Then aX = 0
      If X + aX > PicInfo.bmWidth Then aX = 0
      If Y + aY < 1 Then aY = 0
      If Y + aY > PicInfo.bmHeight Then aY = 0
      iDATA(1, X, Y) = iDATA(1, X + aX, Y + aY)
      iDATA(2, X, Y) = iDATA(2, X + aX, Y + aY)
      iDATA(3, X, Y) = iDATA(3, X + aX, Y + aY)
    Next X, Y
End Sub

Public Sub Relief()
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      B = 2 * CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + CLng(bDATA(1, X, Y - 1)) - CLng(bDATA(1, X, Y + 1)) - CLng(bDATA(1, X + 1, Y)) - 2 * CLng(bDATA(1, X + 1, Y + 1))
      G = 2 * CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + CLng(bDATA(2, X, Y - 1)) - CLng(bDATA(2, X, Y + 1)) - CLng(bDATA(2, X + 1, Y)) - 2 * CLng(bDATA(2, X + 1, Y + 1))
      R = 2 * CLng(bDATA(3, X - 1, Y - 1)) + CLng(bDATA(3, X - 1, Y)) + CLng(bDATA(3, X, Y - 1)) - CLng(bDATA(3, X, Y + 1)) - CLng(bDATA(3, X + 1, Y)) - 2 * CLng(bDATA(3, X + 1, Y + 1))
      B = (CLng(bDATA(1, X, Y)) + B) \ 2 + 50
      G = (CLng(bDATA(2, X, Y)) + G) \ 2 + 50
      R = (CLng(bDATA(3, X, Y)) + R) \ 2 + 50
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Pixelize(ByVal PixSize As Long)
  Dim pX As Long, pY As Long
  Dim sX As Long, sY As Long
  Dim mC As Long
  B = 0: G = 0: R = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      If ((X - 1) Mod PixSize) = 0 Then
        sX = ((X - 1) \ PixSize) * PixSize + 1
        sY = ((Y - 1) \ PixSize) * PixSize + 1
        B = 0: G = 0: R = 0: mC = 0
        For pX = sX To sX + PixSize - 1
          For pY = sY To sY + PixSize - 1
            If (pX <= PicInfo.bmWidth) And (pY <= PicInfo.bmHeight) Then
              B = B + CLng(bDATA(1, pX, pY))
              G = G + CLng(bDATA(2, pX, pY))
              R = R + CLng(bDATA(3, pX, pY))
              mC = mC + 1
            End If
          Next pY
        Next pX
        If mC > 0 Then
          B = B \ mC
          G = G \ mC
          R = R \ mC
        End If
      End If
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Blur()
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, X - 1, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + _
        CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X, Y - 1)) + _
        CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y - 1)) + _
        CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X + 1, Y + 1))
      B = B \ 8
      G = CLng(iDATA(2, X - 1, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + _
        CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X, Y - 1)) + _
        CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y - 1)) + _
        CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X + 1, Y + 1))
      G = G \ 8
      R = CLng(iDATA(3, X - 1, Y - 1)) + CLng(iDATA(3, X - 1, Y)) + _
        CLng(iDATA(3, X - 1, Y + 1)) + CLng(iDATA(3, X, Y - 1)) + _
        CLng(iDATA(3, X, Y + 1)) + CLng(iDATA(3, X + 1, Y - 1)) + _
        CLng(iDATA(3, X + 1, Y)) + CLng(iDATA(3, X + 1, Y + 1))
      R = R \ 8
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub BlurMore()
On Error Resume Next
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, X - 2, Y - 2)) + CLng(iDATA(1, X - 2, Y - 1)) + CLng(iDATA(1, X - 2, Y)) + CLng(iDATA(1, X - 2, Y + 1)) + CLng(iDATA(1, X - 2, Y + 2)) + _
             CLng(iDATA(1, X - 1, Y - 2)) + CLng(iDATA(1, X - 1, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X - 1, Y + 2)) + _
             CLng(iDATA(1, X, Y - 2)) + CLng(iDATA(1, X, Y - 1)) + CLng(iDATA(1, X, Y)) + CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X, Y + 2)) + _
             CLng(iDATA(1, X + 1, Y - 2)) + CLng(iDATA(1, X + 1, Y - 1)) + CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X + 1, Y + 1)) + CLng(iDATA(1, X + 1, Y + 2)) + _
             CLng(iDATA(1, X + 2, Y - 2)) + CLng(iDATA(1, X + 2, Y - 1)) + CLng(iDATA(1, X + 2, Y)) + CLng(iDATA(1, X + 2, Y + 1)) + CLng(iDATA(1, X + 2, Y + 2))
      B = B \ 25
      G = CLng(iDATA(2, X - 2, Y - 2)) + CLng(iDATA(2, X - 2, Y - 1)) + CLng(iDATA(2, X - 2, Y)) + CLng(iDATA(2, X - 2, Y + 1)) + CLng(iDATA(2, X - 2, Y + 2)) + _
             CLng(iDATA(2, X - 1, Y - 2)) + CLng(iDATA(2, X - 1, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X - 1, Y + 2)) + _
             CLng(iDATA(2, X, Y - 2)) + CLng(iDATA(2, X, Y - 1)) + CLng(iDATA(2, X, Y)) + CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X, Y + 2)) + _
             CLng(iDATA(2, X + 1, Y - 2)) + CLng(iDATA(2, X + 1, Y - 1)) + CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X + 1, Y + 1)) + CLng(iDATA(2, X + 1, Y + 2)) + _
             CLng(iDATA(2, X + 2, Y - 2)) + CLng(iDATA(2, X + 2, Y - 1)) + CLng(iDATA(2, X + 2, Y)) + CLng(iDATA(2, X + 2, Y + 1)) + CLng(iDATA(2, X + 2, Y + 2))
      G = G \ 25
      R = CLng(iDATA(3, X - 2, Y - 2)) + CLng(iDATA(3, X - 2, Y - 1)) + CLng(iDATA(3, X - 2, Y)) + CLng(iDATA(3, X - 2, Y + 1)) + CLng(iDATA(3, X - 2, Y + 2)) + _
             CLng(iDATA(3, X - 1, Y - 2)) + CLng(iDATA(3, X - 1, Y - 1)) + CLng(iDATA(3, X - 1, Y)) + CLng(iDATA(3, X - 1, Y + 1)) + CLng(iDATA(3, X - 1, Y + 2)) + _
             CLng(iDATA(3, X, Y - 2)) + CLng(iDATA(3, X, Y - 1)) + CLng(iDATA(3, X, Y)) + CLng(iDATA(3, X, Y + 1)) + CLng(iDATA(3, X, Y + 2)) + _
             CLng(iDATA(3, X + 1, Y - 2)) + CLng(iDATA(3, X + 1, Y - 1)) + CLng(iDATA(3, X + 1, Y)) + CLng(iDATA(3, X + 1, Y + 1)) + CLng(iDATA(3, X + 1, Y + 2)) + _
             CLng(iDATA(3, X + 2, Y - 2)) + CLng(iDATA(3, X + 2, Y - 1)) + CLng(iDATA(3, X + 2, Y)) + CLng(iDATA(3, X + 2, Y + 1)) + CLng(iDATA(3, X + 2, Y + 2))
      R = R \ 25
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub NearestColorBW(Col&)
  Dim cB As Long, cG As Long, cR As Long
  cR = Col Mod 256
  cG = ((Col And &HFF00&) \ 256&) Mod 256&
  cB = (Col And &HFF0000) \ 65536

  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, X, Y))
      G = CLng(bDATA(2, X, Y))
      R = CLng(bDATA(3, X, Y))
      If (R < cR) And (G < cG) And (B < cB) Then
        iDATA(1, X, Y) = 0
        iDATA(2, X, Y) = 0
        iDATA(3, X, Y) = 0
      Else
        iDATA(1, X, Y) = 255
        iDATA(2, X, Y) = 255
        iDATA(3, X, Y) = 255
      End If
    Next X, Y
End Sub

Public Sub Charcoal() 'charcoal
Dim tCol&
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
            R = Abs(R * (G - B + G + R)) / 256
            G = Abs(R * (B - G + B + R)) / 256
            B = Abs(G * (B - G + B + R)) / 256
            tCol = RGB(R, G, B)
            R = Abs(tCol Mod 256)
            G = Abs((tCol \ 256) Mod 256)
            B = Abs(tCol \ 256 \ 256)
            R = (R + G + B) / 3
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = R
    iDATA(1, X, Y) = R
    Next X, Y
End Sub

Public Sub OrderedDither()
  Dim VecDither(1 To 4, 1 To 4) As Byte
  Dim cX As Long, cY As Long
    VecDither(1, 1) = 1:    VecDither(1, 2) = 9
    VecDither(1, 3) = 3:    VecDither(1, 4) = 11
    VecDither(2, 1) = 13:   VecDither(2, 2) = 5
    VecDither(2, 3) = 15:   VecDither(2, 4) = 7
    VecDither(3, 1) = 4:    VecDither(3, 2) = 12
    VecDither(3, 3) = 2:    VecDither(3, 4) = 10
    VecDither(4, 1) = 16:   VecDither(4, 2) = 8
    VecDither(4, 3) = 14:   VecDither(4, 4) = 6
  For X = 0 To 765
    Speed(X) = 1 + (X \ 3) \ 16
  Next X
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, X, Y))
      G = CLng(bDATA(2, X, Y))
      R = CLng(bDATA(3, X, Y))
      B = Speed(R + G + B)
      cX = 1 + ((X - 1) Mod 4)
      cY = 1 + ((Y - 1) Mod 4)
      If B < VecDither(cX, cY) Then
        iDATA(1, X, Y) = 0
        iDATA(2, X, Y) = 0
        iDATA(3, X, Y) = 0
      Else
        iDATA(1, X, Y) = 255
        iDATA(2, X, Y) = 255
        iDATA(3, X, Y) = 255
      End If
    Next X, Y
End Sub

Public Sub FloydSteinberg(ByVal PalWeight As Long)
  Dim Erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 4, -255 To 255) As Long
  
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  For X = -255 To 255
    PartErr(1, X) = (7 * X) \ 16
    PartErr(2, X) = (3 * X) \ 16
    PartErr(3, X) = (5 * X) \ 16
    PartErr(4, X) = (1 * X) \ 16
  Next X
  Erro = 0
  ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
  For X = 1 To PicInfo.bmWidth
    VecErro(1, X) = 0
    VecErro(2, X) = 0
  Next X
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, X, Y))
      G = CLng(bDATA(2, X, Y))
      R = CLng(bDATA(3, X, Y))
      B = Speed(R + G + B)
      mCol = mCol + B
      nCol = nCol + 1
    Next X, Y
  mCol = mCol \ nCol
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, X, Y))
      G = CLng(bDATA(2, X, Y))
      R = CLng(bDATA(3, X, Y))
      B = Speed(R + G + B)
      B = B + (VecErro(1, X) * 10) \ PalWeight
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If B < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, X, Y) = nCol
      iDATA(2, X, Y) = nCol
      iDATA(3, X, Y) = nCol
      Erro = B - nCol
      If X < PicInfo.bmWidth Then VecErro(1, X + 1) = VecErro(1, X + 1) + PartErr(1, Erro)
      If Y < PicInfo.bmHeight Then
        If X > 1 Then VecErro(2, X - 1) = VecErro(2, X - 1) + PartErr(2, Erro)
        VecErro(2, X) = VecErro(2, X) + PartErr(3, Erro)
        If X < PicInfo.bmWidth Then VecErro(2, X + 1) = VecErro(2, X + 1) + PartErr(4, Erro)
      End If
    Next X
    For X = 1 To PicInfo.bmWidth
      VecErro(1, X) = VecErro(2, X)
      VecErro(2, X) = 0
    Next X, Y
End Sub

Public Sub Contour(Col As Long)
  Dim cB As Long, cG As Long, cR As Long
  cR = Col Mod 256
  cG = ((Col And &HFF00&) \ 256&) Mod 256&
  cB = (Col And &HFF0000) \ 65536
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      B = CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + CLng(bDATA(1, X - 1, Y + 1)) + CLng(bDATA(1, X, Y - 1)) + _
          CLng(bDATA(1, X, Y + 1)) + CLng(bDATA(1, X + 1, Y - 1)) + CLng(bDATA(1, X + 1, Y)) + CLng(bDATA(1, X + 1, Y + 1))
      G = CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + CLng(bDATA(2, X - 1, Y + 1)) + CLng(bDATA(2, X, Y - 1)) + _
          CLng(bDATA(2, X, Y + 1)) + CLng(bDATA(2, X + 1, Y - 1)) + CLng(bDATA(2, X + 1, Y)) + CLng(bDATA(2, X + 1, Y + 1))
      R = CLng(bDATA(3, X - 1, Y - 1)) + CLng(bDATA(3, X - 1, Y)) + CLng(bDATA(3, X - 1, Y + 1)) + CLng(bDATA(3, X, Y - 1)) + _
          CLng(bDATA(3, X, Y + 1)) + CLng(bDATA(3, X + 1, Y - 1)) + CLng(bDATA(3, X + 1, Y)) + CLng(bDATA(3, X + 1, Y + 1))
      B = 8 * CLng(bDATA(1, X, Y)) - B + cB
      G = 8 * CLng(bDATA(2, X, Y)) - G + cG
      R = 8 * CLng(bDATA(3, X, Y)) - R + cR
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub EdgeEnhance(ByVal Factor As Long)
  Dim mf As Long, dF As Long
  mf = 9 + Factor
  dF = 1 + Factor
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, X - 1, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X, Y - 1)) + _
        CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y - 1)) + CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X + 1, Y + 1))
      B = (mf * CLng(iDATA(1, X, Y)) - B) \ dF
      G = CLng(iDATA(2, X - 1, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X, Y - 1)) + _
        CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y - 1)) + CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X + 1, Y + 1))
      G = (mf * CLng(iDATA(2, X, Y)) - G) \ dF
      R = CLng(iDATA(3, X - 1, Y - 1)) + CLng(iDATA(3, X - 1, Y)) + CLng(iDATA(3, X - 1, Y + 1)) + CLng(iDATA(3, X, Y - 1)) + _
        CLng(iDATA(3, X, Y + 1)) + CLng(iDATA(3, X + 1, Y - 1)) + CLng(iDATA(3, X + 1, Y)) + CLng(iDATA(3, X + 1, Y + 1))
      R = (mf * CLng(iDATA(3, X, Y)) - R) \ dF
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub ConnectedContour()
  Dim V As Long
  Dim I As Long
  Dim vMin As Long
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      For I = 1 To 3
        vMin = 255
        V = CLng(bDATA(I, X - 1, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X + 1, Y - 1))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(I, X - 1, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X + 1, Y))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(I, X - 1, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(I, X + 1, Y + 1))
        If V < vMin Then vMin = V
        
        iDATA(I, X, Y) = CLng(iDATA(I, X, Y)) - vMin
      Next I
    Next X, Y
End Sub

Public Sub AddNoise(ByVal Factor As Long)
  Dim V As Long
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      G = CLng(bDATA(2, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      R = CLng(bDATA(3, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Fog(C%)
Dim tt1%
  For Y = 1 To PicInfo.bmHeight
tt1 = (Rnd * C) - 2
    For X = 1 To PicInfo.bmWidth
        R = Abs(iDATA(3, X, Y) + tt1)
        G = Abs(iDATA(2, X, Y) + tt1)
        B = Abs(iDATA(1, X, Y) + tt1)
      If R > 255 Then R = 255
      If R < 0 Then R = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, X, Y) = B
      iDATA(2, X, Y) = G
      iDATA(3, X, Y) = R
    Next X, Y
End Sub

Public Sub Freeze(Strength As Single) 'freeze
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
R = Abs((R - G - B) * Strength)
G = Abs((G - B - R) * Strength)
B = Abs((B - R - G) * Strength)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub Brown(C%) 'brown
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
R = Abs(G * B) / C
G = Abs(B * R) / 256
B = Abs(R * G) / 256
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
    Next X, Y
End Sub

Public Sub Liquid() 'liquid
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
R = ((G - B) ^ 2) / 125
G = ((R - B) ^ 2) / 125
B = ((R - G) ^ 2) / 125
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
    Next X, Y
End Sub

Public Sub Yellow() 'yellow
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
B = ((G - R) ^ 2) / 125
R = ((G - B) ^ 2) / 125
G = ((B + R) ^ 2) / 125
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
    Next X, Y
End Sub

Public Sub DarkMoon() 'dark moon
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
R = Abs(R - 64)
G = Abs(R - 64)
B = Abs(R - 64)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub TotalEclipse() 'eclipse
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
R = Abs(G - 64)
G = Abs(G - 64)
B = Abs(G - 64)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub PurpleRain() 'purple
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
R = Abs(G + R / 2)
G = Abs(B + G / 2)
B = Abs(R + B / 2)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub Spooky() 'Spooky
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
G = Abs(R + G / 2)
B = Abs(R + B / 2)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub UnReal() 'unreal
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
        R = Abs(Sin(Atn(G / B)) * 125 + 20)
        G = Abs(Sin(Atn(R / B)) * 125 + 20)
        B = Abs(Sin(Atn(R / G)) * 125 + 20)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub Flame() 'flame
Dim C As Long
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
      C = (R + G + B) / 3
        If R > B Then
            R = Abs(R + C)
            B = Abs(B - C)
        End If
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub Effect0(Eff%)
On Error Resume Next
Dim C&
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
Select Case Eff
Case 0
    G = (R + G) / 2
    R = G
    B = B * (Atn(G) * 2)
Case 1
    G = (R + G) / 2
    R = G
Case 2
    R = (R + B) / 2
    B = R
Case 3
    G = (G + B) / 2
    B = G
Case 4
    G = (B + G) / 2
    B = G
    R = R * (Atn(G) * 2)
Case 5
    B = (B + R) / 2
    R = B
    G = G * (Atn(R) * 2)
Case 6
    B = Sin(B) * B
    R = Sin(R) * R
    G = Sin(G) * G
Case 7
    C = (R + G + B) / 12
    B = Abs(Not (G + C))
    R = Abs(Not (B + C))
    G = Abs(Not (R + C))
Case 8
    B = G
    G = R
Case 9
    R = R / 2
    B = G / 2
    G = R
Case 10
    R = R
    B = G / 2
    G = R / 2
Case 11
    R = R + Abs(Sin(R) * 64)
    G = G + Abs(Sin(G) * 64)
    B = B + Abs(Sin(B) * 64)
End Select
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub Aquarel() 'aquarel
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y)
    G = iDATA(2, X, Y)
    B = iDATA(1, X, Y)
If R < 128 And G < 128 And B < 128 Then
R = 2 * R: G = 2 * G: B = 2 * B
Else
R = R / 2: G = G / 2: B = B / 2
End If
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub Erode(pct%) 'erode
On Error Resume Next
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
    R = Abs(iDATA(3, X, Y) Xor pct)
    G = Abs(iDATA(2, X, Y) Xor pct)
    B = Abs(iDATA(1, X, Y) Xor pct)
    iDATA(3, X, Y) = R
    iDATA(2, X, Y) = G
    iDATA(1, X, Y) = B
Next X, Y
End Sub

Public Sub EffectX(Ob As Object, Strength%, Wave As Single, Eff%) 'wave x
On Error Resume Next
Dim Degree As Single
Wave = Wave / 10

  For Y = 1 To PicInfo.bmHeight
Degree = (Y * Wave / 180 * 3.14)
If Eff = 0 Then k = (Cos(Degree) * Strength)
If Eff = 1 Then k = Abs(Cos(Degree) * Strength)
If k < 0 Then
    For Tx = 1 To PicInfo.bmWidth
        iDATA(3, X, Y) = iDATA(3, PicInfo.bmWidth, Y)
        iDATA(2, X, Y) = iDATA(2, PicInfo.bmWidth, Y)
        iDATA(1, X, Y) = iDATA(1, PicInfo.bmWidth, Y)
    Next Tx
    Else
    For Tx = 1 To PicInfo.bmWidth
        iDATA(3, X, Y) = iDATA(3, 1, Y)
        iDATA(2, X, Y) = iDATA(2, 1, Y)
        iDATA(1, X, Y) = iDATA(1, 1, Y)
    Next Tx
End If
    For X = 1 To PicInfo.bmWidth
iDATA(3, X, Y) = bDATA(3, X + k, Y)
iDATA(2, X, Y) = bDATA(2, X + k, Y)
iDATA(1, X, Y) = bDATA(1, X + k, Y)
Next X, Y
End Sub

Public Sub EffectY(Ob As Object, Strength%, Wave As Single, Eff%) 'wave x
On Error Resume Next
Dim Degree As Single
Wave = Wave / 10

    For X = 1 To PicInfo.bmWidth
Degree = (X * Wave / 180 * 3.14)
If Eff = 0 Then k = (Cos(Degree) * Strength)
If Eff = 1 Then k = Abs(Cos(Degree) * Strength)
If k < 0 Then
    For Tx = 1 To PicInfo.bmHeight
        iDATA(3, X, Y) = iDATA(3, X, PicInfo.bmHeight)
        iDATA(2, X, Y) = iDATA(2, X, PicInfo.bmHeight)
        iDATA(1, X, Y) = iDATA(1, X, PicInfo.bmHeight)
    Next Tx
    Else
    For Tx = 1 To PicInfo.bmHeight
        iDATA(3, X, Y) = iDATA(3, X, 0)
        iDATA(2, X, Y) = iDATA(2, X, 0)
        iDATA(1, X, Y) = iDATA(1, X, 0)
    Next Tx
End If
  For Y = 1 To PicInfo.bmHeight
iDATA(3, X, Y) = bDATA(3, X, Y + k)
iDATA(2, X, Y) = bDATA(2, X, Y + k)
iDATA(1, X, Y) = bDATA(1, X, Y + k)
Next Y, X
End Sub

Public Sub Echo(Ob As Object, ENr%, ERed%, EX%, EY%) 'echo picture
Dim EchoW&, EchoH&
Dim EchoLeft%, EchoTop%, Phase%
    Dim CA As COLORADJUSTMENT
'On Error Resume Next
EchoW = CDC1.Pic3.Width - 1
EchoH = CDC1.Pic3.Height - 1
Phase = 0
For xx = 0 To ENr - 1
    GetColorAdjustment Ob.hDC, CA
    CA.caSize = Len(CA)
    CA.caBrightness = -100
    CA.caIlluminantIndex = ILLUMINANT_A
    If GetStretchBltMode(Ob.hDC) <> HALFTONE Then
        SetStretchBltMode Ob.hDC, HALFTONE
    End If
If CDC5.Check3.Value = 1 Then Phase = xx
EchoW = EchoW * (100 - ERed) / 100
EchoH = EchoH * (100 - ERed) / 100
EchoLeft = (CDC1.Pic3.Width / 2) - (EchoW / 2) + ((Phase + 1) * EX)
EchoTop = (CDC1.Pic3.Height / 2) - (EchoH / 2) + ((Phase + 1) * EY)
'Ob.PaintPicture BLM1.Pic3, EchoLeft, EchoTop, EchoW, EchoH
StretchBlt Ob.hDC, EchoLeft, EchoTop, EchoW, EchoH, CDC1.Pic3.hDC, 0, 0, CDC1.Pic3.Width, CDC1.Pic3.Height, vbSrcCopy
Next xx
End Sub

Public Sub Tile(Ob As Object, XTile%, YTile%)
On Error Resume Next
Dim TileX%, TileY%
    Dim CA As COLORADJUSTMENT
TileX = Int(Ob.Width / XTile)
TileY = Int(Ob.Height / YTile)
For xx = 0 To XTile '- 1
For yy = 0 To YTile '- 1
    GetColorAdjustment Ob.hDC, CA
    CA.caSize = Len(CA)
    CA.caBrightness = -100
    CA.caIlluminantIndex = ILLUMINANT_A
    If GetStretchBltMode(Ob.hDC) <> HALFTONE Then
        SetStretchBltMode Ob.hDC, HALFTONE
    End If
'Ob.PaintPicture BLM1.Pic3, xx * TileX, yy * TileY, TileX, TileY
StretchBlt Ob.hDC, xx * TileX, yy * TileY, TileX, TileY, CDC1.Pic3.hDC, 0, 0, CDC1.Pic3.Width, CDC1.Pic3.Height, vbSrcCopy
Next yy, xx
End Sub

Public Sub Blinds(pct1%, Reverse As Boolean, HV As Boolean)  ' blinds
Dim rt1%
On Error Resume Next
If Reverse = False Then
rt1 = 0
Else
rt1 = pct1
End If
If HV = False Then 'horizontal
    For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
R = iDATA(3, X, Y) - (rt1 * iDATA(3, X, Y) / pct1)
G = iDATA(2, X, Y) - (rt1 * iDATA(2, X, Y) / pct1)
B = iDATA(1, X, Y) - (rt1 * iDATA(1, X, Y) / pct1)
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
        iDATA(3, X, Y) = R
        iDATA(2, X, Y) = G
        iDATA(1, X, Y) = B
Next X
        If Reverse = False Then
            rt1 = rt1 + 1
            If rt1 = pct1 Then rt1 = 0
        Else
            rt1 = rt1 - 1
            If rt1 = 0 Then rt1 = pct1
        End If
        Next Y
Else 'hv=true
    For X = 1 To PicInfo.bmWidth
    For Y = 1 To PicInfo.bmHeight
R = iDATA(3, X, Y) - (rt1 * iDATA(3, X, Y) / pct1)
G = iDATA(2, X, Y) - (rt1 * iDATA(2, X, Y) / pct1)
B = iDATA(1, X, Y) - (rt1 * iDATA(1, X, Y) / pct1)
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
        iDATA(3, X, Y) = R
        iDATA(2, X, Y) = G
        iDATA(1, X, Y) = B
Next Y
        If Reverse = False Then
            rt1 = rt1 + 1
            If rt1 = pct1 Then rt1 = 0
        Else
            rt1 = rt1 - 1
            If rt1 = 0 Then rt1 = pct1
        End If
        Next X
End If
End Sub

Public Sub Blinds3(pct1%, HV As Boolean) ' bump blinds
Dim rt1%, Rtt As Boolean
On Error Resume Next
rt1 = 0
Rtt = False
If HV = False Then 'hor bump
    For Y = 1 To PicInfo.bmHeight
        For X = 1 To PicInfo.bmWidth
    R = iDATA(3, X, Y) - (rt1 * iDATA(3, X, Y) / pct1)
    G = iDATA(2, X, Y) - (rt1 * iDATA(2, X, Y) / pct1)
    B = iDATA(1, X, Y) - (rt1 * iDATA(1, X, Y) / pct1)
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
            iDATA(3, X, Y) = R
            iDATA(2, X, Y) = G
            iDATA(1, X, Y) = B
    Next X
        If Rtt = False Then
        rt1 = rt1 + 2
        Else
        rt1 = rt1 - 2
        End If
            If rt1 >= pct1 Then Rtt = True
            If rt1 <= 0 Then Rtt = False
    Next Y
Else 'vert bump
        For X = 1 To PicInfo.bmWidth
    For Y = 1 To PicInfo.bmHeight
    R = iDATA(3, X, Y) - (rt1 * iDATA(3, X, Y) / pct1)
    G = iDATA(2, X, Y) - (rt1 * iDATA(2, X, Y) / pct1)
    B = iDATA(1, X, Y) - (rt1 * iDATA(1, X, Y) / pct1)
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
            iDATA(3, X, Y) = R
            iDATA(2, X, Y) = G
            iDATA(1, X, Y) = B
    Next Y
        If Rtt = False Then
        rt1 = rt1 + 2
        Else
        rt1 = rt1 - 2
        End If
            If rt1 >= pct1 Then Rtt = True
            If rt1 <= 0 Then Rtt = False
    Next X
End If
End Sub

Public Sub Mozaic(Br%) 'mozaic
Dim Br2%, MColorR&, MColorG&, MColorB&, qq%, pp%
On Error Resume Next
Br2 = Int(Br / 2)
        For X = 1 To PicInfo.bmWidth Step Br
    For Y = 1 To PicInfo.bmHeight Step Br
MColorR = iDATA(3, X + Br2, Y + Br2)
MColorG = iDATA(2, X + Br2, Y + Br2)
MColorB = iDATA(1, X + Br2, Y + Br2)
    For qq = X To X + Br - 1
    For pp = Y To Y + Br - 1
    iDATA(3, qq, pp) = MColorR
    iDATA(2, qq, pp) = MColorG
    iDATA(1, qq, pp) = MColorB
    Next pp, qq
Next Y, X
End Sub

Public Sub Mozaic2(Br%)  'mozaic2
Dim Br2%, MColorR&, MColorG&, MColorB&, qq%, pp%, R1&, G1&, B1&
On Error Resume Next
Br2 = Int(Br / 2)
        For X = 1 To PicInfo.bmWidth Step Br
    For Y = 1 To PicInfo.bmHeight Step Br
R1 = iDATA(3, X + Br2, Y + Br2)
G1 = iDATA(2, X + Br2, Y + Br2)
B1 = iDATA(1, X + Br2, Y + Br2)
    For qq = X To X + Br - 1
    For pp = Y To Y + Br - 1
    If qq = X Or pp = Y Or qq = X + Br - 1 Or pp = Y + Br - 1 Then
        MColorR = iDATA(3, qq, pp) - ((Rnd * 10) - 5)
        If MColorR < 0 Then MColorR = 0
        MColorG = iDATA(2, qq, pp) - ((Rnd * 10) - 5)
        If MColorG < 0 Then MColorG = 0
        MColorB = iDATA(1, qq, pp) - ((Rnd * 10) - 5)
        If MColorB < 0 Then MColorB = 0
                iDATA(3, qq, pp) = MColorR
                iDATA(2, qq, pp) = MColorG
                iDATA(1, qq, pp) = MColorB
    Else
    iDATA(3, qq, pp) = R1
    iDATA(2, qq, pp) = G1
    iDATA(1, qq, pp) = B1
    End If
    Next pp, qq
Next Y, X
End Sub

Public Sub SetWaveLinesH(Dist%, Wave!, Ampl%, LW%, Eff%)
On Error Resume Next
Dim Degree As Single, k!, pp%
With CDC1
Wave = Wave / 10
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
.Pic3.Picture = Im
        For X = 0 To .Pic3.Width
    For Y = 0 To .Pic3.Height Step Dist
Degree = X * (Wave) / 180 * 3.14
If Eff = 0 Then k = Cos(Degree) * Ampl
If Eff = 1 Then k = Abs(Cos(Degree) * Ampl)
If Eff = 2 Then k = -Abs(Cos(Degree) * Ampl)
    For pp = 0 To LW - 1
SetPixel CDC1.Pic3.hDC, X, Y + k + pp, CDC5.Label66.BackColor
    Next pp
Next Y, X
.Pic3.Refresh
End With
End Sub

Public Sub SetWaveLinesV(Dist%, Wave!, Ampl%, LW%, Eff%)
On Error Resume Next
Dim Degree As Single, k!, pp%
With CDC1
Wave = Wave / 10
.Pic3.Width = .Pic1(PicIdx).Width
.Pic3.Height = .Pic1(PicIdx).Height
.Pic3.Picture = Im
        For X = 0 To .Pic3.Width Step Dist
    For Y = 0 To .Pic3.Height
Degree = Y * (Wave) / 180 * 3.14
If Eff = 0 Then k = Cos(Degree) * Ampl
If Eff = 1 Then k = Abs(Cos(Degree) * Ampl)
If Eff = 2 Then k = -Abs(Cos(Degree) * Ampl)
    For pp = 0 To LW - 1
SetPixel CDC1.Pic3.hDC, X + k + pp, Y, CDC5.Label66.BackColor
Next pp
Next Y, X
.Pic3.Refresh
End With
End Sub

Public Sub BrightenPic(FC%)
  sF = (FC + 100) / 100
  For X = 0 To 255
    Speed(X) = X * sF
    If Speed(X) > 255 Then Speed(X) = 255
    If Speed(X) < 0 Then Speed(X) = 0
  Next X
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      iDATA(1, X, Y) = Speed(bDATA(1, X, Y))
      iDATA(2, X, Y) = Speed(bDATA(2, X, Y))
      iDATA(3, X, Y) = Speed(bDATA(3, X, Y))
    Next X
  Next Y
End Sub
