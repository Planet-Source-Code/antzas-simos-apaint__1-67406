VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResizeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Strin As String
Private MoveX As Single
Private MoveY As Single
Private StX As Single
Private StY As Single
Private Button1 As Integer
Private XX As Single
Private YY As Single

Public Property Let X(aX As Single)
XX = aX
End Property

Public Property Get X() As Single
X = XX
End Property

Public Property Let Y(aY As Single)
YY = aY
End Property

Public Property Get Y() As Single
Y = YY
End Property

Public Property Let Button(BUtt As Integer)
Button1 = BUtt
End Property

Public Property Get Button() As Integer
Button = Button1
End Property

Public Property Let SetStartUpPosX(StPos As Single)
StX = StPos
End Property

Public Property Get SetStartUpPosX() As Single
SetStartUpPosX = StX
End Property

Public Property Let SetStartUpPosY(StPos As Single)
StY = StPos
End Property

Public Property Get SetStartUpPosY() As Single
SetStartUpPosY = StY
End Property

Public Property Let MovePosX(StPos As Single)
MoveX = StPos
End Property

Public Property Get MovePosX() As Single
MovePosX = MoveX
End Property

Public Property Let MovePosY(StPos As Single)
MoveY = StPos
End Property

Public Property Get MovePosY() As Single
MovePosY = MoveY
End Property

Public Sub MoveTheObject(Obj As PictureBox)
'CopyPic.Left = CopyPic.Left + X - XX
'CopyPic.Top = CopyPic.Top + Y - YY
Obj.Move Obj.Left + MovePosX - SetStartUpPosX, Obj.Top + MovePosY - SetStartUpPosY
End Sub

Public Sub Resize(Object As PictureBox, PicBox As PictureBox)

If Object.Visible = False Then Exit Sub
PicBox.MousePointer = 99
'Form1.StBar.SimpleText = X & " - " & Y & " - " & Button
If X > Object.Left + Object.Width And X < Object.Left + Object.Width + 11 Then
    If Y < Object.Top + Object.Height And Y > Object.Top Then
        PicBox.MousePointer = 9
    End If
End If
If Y > Object.Top + Object.Height And Y < Object.Top + Object.Height + 11 Then
    If X < Object.Left + Object.Width And X > Object.Left Then
        PicBox.MousePointer = 7
    End If
End If
If X > Object.Left + Object.Width And X < Object.Left + Object.Width + 11 Then
    If Y > Object.Top + Object.Height And Y < Object.Top + Object.Height + 11 Then
        PicBox.MousePointer = 8
    End If
End If
If Button = 1 Then
    If X > Object.Left + Object.Width - 31 And X < Object.Left + Object.Width + 31 Then
        If Y < Object.Top + Object.Height And Y > Object.Top Then
            PicBox.MousePointer = 9
            SizeChange Object, PicBox, X, 0
            'Object.Width = Abs(Object.Left - X)
        End If
    End If
    If Y > Object.Top + Object.Height - 31 And Y < Object.Top + Object.Height + 31 Then
        If X < Object.Left + Object.Width And X > Object.Left Then
            PicBox.MousePointer = 7
            SizeChange Object, PicBox, 0, Y
        End If
    End If
    If X > Object.Left + Object.Width - 31 And X < Object.Left + Object.Width + 31 Then
        If Y > Object.Top + Object.Height - 31 And Y < Object.Top + Object.Height + 31 Then
            PicBox.MousePointer = 8
            SizeChange Object, PicBox, X, Y
        End If
    End If
End If
End Sub

Private Sub SizeChange(Optional Obj As PictureBox, Optional Pic As PictureBox, Optional ByVal aX As Long, Optional ByVal aY As Long)

If aX <> 0 Then
    If aX > 0 Then
        Obj.Width = Abs(Obj.Left - aX)
    End If
End If
If aY <> 0 Then
    If aY > 0 Then
        Obj.Height = Abs(Obj.Top - aY)
    End If
End If
Strin = "Picture Size " & Obj.Width & "x" & Obj.Height
'Form1.StBar.SimpleText = "Picture Size " & Obj.Width & "x" & Obj.Height
End Sub

Public Property Get TextStr() As String
TextStr = Strin
End Property
