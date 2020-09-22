Attribute VB_Name = "FiltersMod"
Option Explicit

Public Sub greys()
Dim Picbox As PictureBox
Dim x, y, texr, texg, texb, greycolor As Integer

Waitt.Show 'vbModal
DoEvents

If Form1.CopyPic.Visible = True Then
    Set Picbox = Form1.CopyPic
Else
    Set Picbox = Form1.MainPic
End If
    
Waitt.Pg.Max = (Picbox.Height)
For y = 0 To (Picbox.Height)
    Waitt.Pg.Value = y
    For x = 0 To (Picbox.Width)
        texr = Picbox.Point(x, y) And 255
        texg = (Picbox.Point(x, y) And 65280) / 256
        texb = (Picbox.Point(x, y) And 16711680) / 65535
        
        greycolor = (texr + texg + texb) / 3
        
        SetPixel Picbox.hDC, x, y, RGB(greycolor, greycolor, greycolor)
    Next x
Next y
If Form1.CopyPic.Visible = True Then
    Set Form1.StorePic.Picture = Nothing
    Form1.StorePic.Picture = Picbox.Image
End If
Unload Waitt
Form1.StBar.SimpleText = "Ready"
End Sub

Public Sub Invertt()
Dim Picbox As PictureBox

If Form1.CopyPic.Visible = True Then
    Set Picbox = Form1.CopyPic
Else
    Set Picbox = Form1.MainPic
End If
    
GetObject Picbox.Image, Len(PicInfo), PicInfo
BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight * 3) As Byte
GetBitmapBits Picbox.Image, UBound(PicBits), PicBits(1)
For Cnt = 1 To UBound(PicBits)
    PicBits(Cnt) = 255 - PicBits(Cnt)
Next Cnt
SetBitmapBits Picbox.Image, UBound(PicBits), PicBits(1)
Picbox.Refresh

If Form1.CopyPic.Visible = True Then
    Set Form1.StorePic.Picture = Nothing
    Form1.StorePic.Picture = Picbox.Image
End If
Form1.StBar.SimpleText = "Ready"
End Sub

Public Sub RRED()
Dim Picbox As PictureBox
Dim x, y, texr, texg, texb
Waitt.Show 'vbModal
DoEvents

If Form1.CopyPic.Visible = True Then
    Set Picbox = Form1.CopyPic
Else
    Set Picbox = Form1.MainPic
End If
    
Waitt.Pg.Max = (Picbox.Height)
For y = 0 To (Picbox.Height)
    Waitt.Pg.Value = y
    For x = 0 To (Picbox.Width)
        texr = Picbox.Point(x, y) And 255
        texg = (Picbox.Point(x, y) And 65280) / 256
        texb = (Picbox.Point(x, y) And 16711680) / 65535
        
        SetPixel Picbox.hDC, x, y, RGB(0, texg, texb)
    Next x
Next y
If Form1.CopyPic.Visible = True Then
    Set Form1.StorePic.Picture = Nothing
    Form1.StorePic.Picture = Picbox.Image
End If
Unload Waitt
Form1.StBar.SimpleText = "Ready"
End Sub

Public Sub RGREEN()
Dim Picbox As PictureBox
Dim x, y, texr, texg, texb
Waitt.Show 'vbModal
DoEvents

If Form1.CopyPic.Visible = True Then
    Set Picbox = Form1.CopyPic
Else
    Set Picbox = Form1.MainPic
End If
    
Waitt.Pg.Max = (Picbox.Height)
For y = 0 To (Picbox.Height)
    Waitt.Pg.Value = y
    For x = 0 To (Picbox.Width)
        texr = Picbox.Point(x, y) And 255
        texg = (Picbox.Point(x, y) And 65280) / 256
        texb = (Picbox.Point(x, y) And 16711680) / 65535
        
        SetPixel Picbox.hDC, x, y, RGB(texr, 0, texb)
    Next x
Next y
If Form1.CopyPic.Visible = True Then
    Set Form1.StorePic.Picture = Nothing
    Form1.StorePic.Picture = Picbox.Image
End If
Unload Waitt
Form1.StBar.SimpleText = "Ready"
End Sub

Public Sub RBLUE()
Dim Picbox As PictureBox
Dim x, y, texr, texg, texb
Waitt.Show 'vbModal
DoEvents

If Form1.CopyPic.Visible = True Then
    Set Picbox = Form1.CopyPic
Else
    Set Picbox = Form1.MainPic
End If
    
Waitt.Pg.Max = (Picbox.Height)
For y = 0 To (Picbox.Height)
    Waitt.Pg.Value = y
    For x = 0 To (Picbox.Width)
        texr = Picbox.Point(x, y) And 255
        texg = (Picbox.Point(x, y) And 65280) / 256
        texb = (Picbox.Point(x, y) And 16711680) / 65535
        
        SetPixel Picbox.hDC, x, y, RGB(texr, texg, 0)
    Next x
Next y
If Form1.CopyPic.Visible = True Then
    Set Form1.StorePic.Picture = Nothing
    Form1.StorePic.Picture = Picbox.Image
End If
Unload Waitt
Form1.StBar.SimpleText = "Ready"
End Sub

