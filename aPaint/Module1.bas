Attribute VB_Name = "API_Rotation"
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public PicBits() As Byte, PicInfo As BITMAP
Public Cnt As Long, BytesPerLine As Long

Public RDEG As Integer
Public PI As Double
Public Trans As Double
Public KSize As KanvasSize

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type KanvasSize
    Width As Variant
    Height As Variant
End Type

Public Type FileData
    Name As String
    BrushData As String
End Type

Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Double)

Dim c1x As Integer, c1y As Integer
Dim c2x As Integer, c2y As Integer
Dim a As Double
Dim p1x As Integer, p1y As Integer
Dim p2x As Integer, p2y As Integer
Dim n As Integer, r As Integer
Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long

c1x = pic1.ScaleWidth \ 2
c1y = pic1.ScaleHeight \ 2
c2x = pic2.ScaleWidth \ 2
c2y = pic2.ScaleHeight \ 2
If c2x < c2y Then n = c2y Else n = c2x
n = (n - 1) * 2
pic1hDc = pic1.hDC
pic2hDc = pic2.hDC

For p2x = 0 To n / 2
    For p2y = 0 To n / 2
        If p2x = 0 Then a = PI / 2 Else a = Atn(p2y / p2x)
        r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
        p1x = r * Cos(a + theta)
        p1y = r * Sin(a + theta)
        c0& = GetPixel(pic1hDc, c1x + p1x, c1y + p1y)
        c1& = GetPixel(pic1hDc, c1x - p1x, c1y - p1y)
        c2& = GetPixel(pic1hDc, c1x + p1y, c1y - p1x)
        c3& = GetPixel(pic1hDc, c1x - p1y, c1y + p1x)
        If c0& <> -1 Then SetPixel pic2hDc, c2x + p2x, c2y + p2y, c0&
        If c1& <> -1 Then SetPixel pic2hDc, c2x - p2x, c2y - p2y, c1&
        If c2& <> -1 Then SetPixel pic2hDc, c2x + p2y, c2y - p2x, c2&
        If c3& <> -1 Then SetPixel pic2hDc, c2x - p2y, c2y + p2x, c3&
    Next
Next
End Sub

