VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "aPaint By SimoS"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BrusPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3000
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   36
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox ClGap 
      Height          =   375
      Left            =   7200
      MouseIcon       =   "Form1.frx":0614
      MousePointer    =   99  'Custom
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox StorePic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   33
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar HSc 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      MouseIcon       =   "Form1.frx":091E
      MousePointer    =   99  'Custom
      SmallChange     =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8880
      Width           =   2415
   End
   Begin VB.VScrollBar VSc 
      Height          =   1575
      LargeChange     =   100
      Left            =   7800
      MouseIcon       =   "Form1.frx":0C28
      MousePointer    =   99  'Custom
      SmallChange     =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   9135
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0F32
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox MainToolBox 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   8130
      MouseIcon       =   "Form1.frx":124C
      MousePointer    =   99  'Custom
      ScaleHeight     =   607
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   248
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      Begin VB.CheckBox TransC 
         Caption         =   "Transparent Drop"
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   2520
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   120
         Pattern         =   "*.bmp"
         TabIndex        =   35
         Top             =   4440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox Cop 
         Caption         =   "Copy"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   2520
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox PicSize 
         Height          =   735
         Left            =   480
         ScaleHeight     =   675
         ScaleWidth      =   2835
         TabIndex        =   7
         Top             =   4560
         Visible         =   0   'False
         Width           =   2895
         Begin VB.PictureBox Si 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   5
            Left            =   2040
            ScaleHeight     =   585
            ScaleWidth      =   585
            TabIndex        =   13
            Top             =   120
            Width           =   615
         End
         Begin VB.PictureBox Si 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   4
            Left            =   1440
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.PictureBox Si 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   960
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   11
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox Si 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   600
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   10
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox Si 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   1
            Left            =   360
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   9
            Top             =   600
            Width           =   135
         End
         Begin VB.PictureBox Si 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   0
            Left            =   960
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.PictureBox ZPIC 
         Height          =   3375
         Left            =   0
         ScaleHeight     =   3315
         ScaleWidth      =   3675
         TabIndex        =   25
         Top             =   5400
         Width           =   3735
         Begin VB.PictureBox Picture1 
            Height          =   2775
            Left            =   480
            ScaleHeight     =   181
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   181
            TabIndex        =   27
            Top             =   360
            Width           =   2775
            Begin VB.Shape MIDL 
               DrawMode        =   6  'Mask Pen Not
               Height          =   375
               Left            =   600
               Top             =   240
               Width           =   375
            End
            Begin VB.Image ZIm 
               BorderStyle     =   1  'Fixed Single
               Height          =   615
               Left            =   240
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.ComboBox ZZ 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Zoom   X"
            Height          =   255
            Left            =   480
            TabIndex        =   28
            Top             =   60
            Width           =   735
         End
      End
      Begin VB.PictureBox DegPic 
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   2955
         TabIndex        =   21
         Top             =   4560
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Text            =   "130"
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Degrees"
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox MainToolBox1 
         AutoSize        =   -1  'True
         Height          =   1755
         Left            =   360
         MousePointer    =   2  'Cross
         Picture         =   "Form1.frx":1556
         ScaleHeight     =   1695
         ScaleWidth      =   3000
         TabIndex        =   19
         Top             =   0
         Width           =   3060
      End
      Begin VB.PictureBox SprayPic 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   2955
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   3015
         Begin VB.ComboBox SprayRad 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox SprayThik 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Rad"
            Height          =   255
            Left            =   1320
            TabIndex        =   18
            Top             =   45
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Thik"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   45
            Width           =   375
         End
      End
      Begin VB.ComboBox CDr 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1965
         Width           =   855
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   1680
         Left            =   0
         TabIndex        =   4
         Top             =   2760
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2963
         ButtonWidth     =   1270
         ButtonHeight    =   953
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pencil"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Line"
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Rect"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Circle"
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Select"
               Object.ToolTipText     =   "Select area to copy or crop"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Erase"
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Fill"
               Object.ToolTipText     =   "Fill with color"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Oval"
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Spray"
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Picker"
               Object.ToolTipText     =   "Color picker"
               Style           =   2
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gradient"
               Object.ToolTipText     =   "Gradient fore color to back color"
               Style           =   2
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Brushes"
               Object.ToolTipText     =   "Custom Brushes"
               Style           =   2
            EndProperty
         EndProperty
         MousePointer    =   99
      End
      Begin VB.PictureBox FCP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   1
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox BCP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   2
         Top             =   2040
         Width           =   735
      End
      Begin VB.Image PrvIm 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   2880
         Top             =   4440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Pen Size"
         Height          =   240
         Left            =   1440
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.PictureBox MainPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   240
      MouseIcon       =   "Form1.frx":6836
      MousePointer    =   99  'Custom
      ScaleHeight     =   407
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   3
      Top             =   600
      Width           =   6615
      Begin VB.PictureBox GrPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3840
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer SpT 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   960
         Top             =   480
      End
      Begin VB.PictureBox CopyPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3720
         MousePointer    =   15  'Size All
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   31
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape Sq1 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   1095
         Left            =   1440
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Shape Oval 
         DrawMode        =   6  'Mask Pen Not
         Height          =   375
         Left            =   1680
         Shape           =   2  'Oval
         Top             =   4560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Circ 
         DrawMode        =   6  'Mask Pen Not
         Height          =   615
         Left            =   4680
         Shape           =   3  'Circle
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Line Line1 
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   144
         X2              =   168
         Y1              =   184
         Y2              =   144
      End
   End
   Begin VB.Image MemoryPic 
      Height          =   375
      Index           =   0
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu Fil 
      Caption         =   "File"
      Begin VB.Menu Loa 
         Caption         =   "Open..."
      End
      Begin VB.Menu Sav 
         Caption         =   "Save As..."
      End
      Begin VB.Menu exi 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edit"
      Begin VB.Menu UNd 
         Caption         =   "Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu REDo 
         Caption         =   "Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu SelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu cleA 
         Caption         =   "Clear"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu CopyI 
         Caption         =   "Copy"
      End
      Begin VB.Menu PasteI 
         Caption         =   "Paste"
      End
      Begin VB.Menu StBr 
         Caption         =   "Save to brushes"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu TTooL 
      Caption         =   "Tools"
      Begin VB.Menu zzz 
         Caption         =   "Zoom"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Kanva 
      Caption         =   "Canvas"
      Begin VB.Menu KanvaSize 
         Caption         =   "Size..."
      End
   End
   Begin VB.Menu Filtt 
      Caption         =   "Filters"
      Begin VB.Menu GrScale 
         Caption         =   "Grey Scale"
      End
      Begin VB.Menu INVER 
         Caption         =   "Invert Colors"
      End
      Begin VB.Menu Rvred 
         Caption         =   "Remove Red"
      End
      Begin VB.Menu Rvgreen 
         Caption         =   "Remove Green"
      End
      Begin VB.Menu Rvblue 
         Caption         =   "Remove Blue"
      End
   End
   Begin VB.Menu Hidd1 
      Caption         =   "HidM1"
      Visible         =   0   'False
      Begin VB.Menu CopyPicture 
         Caption         =   "Copy"
      End
      Begin VB.Menu FlV 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu FlH 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu RT180 
         Caption         =   "Rotate"
      End
      Begin VB.Menu RTBD 
         Caption         =   "Rotate By Deg"
      End
      Begin VB.Menu SavToBr 
         Caption         =   "Save to brushes"
      End
   End
   Begin VB.Menu hhidd 
      Caption         =   "Hidd"
      Visible         =   0   'False
      Begin VB.Menu Hidd2 
         Caption         =   "Delete Brush"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OkToDraw As Boolean
Dim OKk As Boolean
Dim Z As Double
Dim Rsz As New ResizeClass
Dim mGradient As New clsGradient
Dim CurPic As Integer
Dim MemI As Integer
Dim WORKING As Boolean
Dim DIS As Double
Dim aX As Long
Dim aY As Long
Dim Radd As Double
Dim SaveDr As Integer
Dim DelL As Integer
Dim ErSize As Integer
Dim XX As Single
Dim YY As Single
Dim Sh As Integer
Dim ForeC As Long
Dim BackC As Long

Private Sub BCP_DblClick()
CD1.ShowColor
BCP.BackColor = CD1.Color
BackC = BCP.BackColor
End Sub

Private Sub CopyI_Click()
If CopyPic.Visible Then
    Clipboard.Clear
    Clipboard.SetData CopyPic.Picture
Else
    Clipboard.Clear
    Clipboard.SetData MainPic.Picture
End If
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Screen.ActiveForm.PopupMenu hhidd
End If
End Sub

Private Sub GrScale_Click()
greys
End Sub

Private Sub Hidd2_Click()
Kill (File1.Path & "\" & File1)
File1.Refresh
Set PrvIm.Picture = Nothing
End Sub

Private Sub INVER_Click()
Invertt
End Sub

Private Sub PasteI_Click()
On Error GoTo bye
StorePic.Picture = Clipboard.GetData(vbCFBitmap)
TransparentBlt MainPic.hDC, 0, 0, StorePic.ScaleWidth, StorePic.ScaleHeight, StorePic.hDC, 0, 0, StorePic.ScaleWidth, StorePic.ScaleHeight, StorePic.BackColor
MainPic.Refresh

ZoomUpdate
Form_Resize
KSize.Width = MainPic.ScaleWidth
KSize.Height = MainPic.ScaleHeight
Exit Sub
bye:
MsgBox "Invalid Picture", vbExclamation + vbOKOnly, "aPaint By SimoS"
End Sub

Private Sub FCP_DblClick()
CD1.ShowColor
FCP.BackColor = CD1.Color
ForeC = FCP.BackColor
End Sub

Private Sub File1_Click()
Set BrusPic.Picture = Nothing
BrusPic.Picture = LoadPicture(File1.Path & "\" & File1)
If BrusPic.ScaleWidth > 57 Or BrusPic.ScaleHeight > 57 Then
    With PrvIm
        .Width = 57
        .Height = 57
        .Stretch = True
        .Picture = BrusPic.Image
    End With
Else
    With PrvIm
        .Stretch = False
        .Width = 57
        .Height = 57
        .Picture = BrusPic.Image
End With
End If
End Sub

Private Sub Rvblue_Click()
RBLUE
End Sub

Private Sub Rvgreen_Click()
RGREEN
End Sub

Private Sub Rvred_Click()
RRED
End Sub

Private Sub SavToBr_Click()
BruD.Show vbModal
End Sub

Private Sub CDr_Click()
MainPic.DrawWidth = CDr.List(CDr.ListIndex)
End Sub

Private Sub cleA_Click()
If MsgBox("Clear ?", vbYesNo + vbExclamation, "aPaint By SimoS") = vbYes Then
    MainPic.Line (0, 0)-(MainPic.ScaleWidth, MainPic.ScaleHeight), MainPic.BackColor, BF
    ZoomUpdate
Else
    Exit Sub
End If
End Sub

Private Sub CopyPic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And CopyPic.Visible = True Then
CopyPic.Visible = False
OKk = False
End If
End Sub

Private Sub exi_Click()
Unload Me
End
End Sub

Private Sub CopyPicture_Click()
Clipboard.Clear
Clipboard.SetData CopyPic.Picture
End Sub

Private Sub FlH_Click()
RDEG = 90
StorePic.Picture = CopyPic.Image
CopyPic.PaintPicture StorePic.Picture, CopyPic.ScaleWidth, 0, _
 -CopyPic.ScaleWidth, CopyPic.ScaleHeight
CopyPic.Picture = CopyPic.Image

StorePic.PaintPicture StorePic.Picture, StorePic.ScaleWidth, 0, _
 -StorePic.ScaleWidth, StorePic.ScaleHeight
StorePic.Picture = StorePic.Image
End Sub

Private Sub FlV_Click()
RDEG = 90
StorePic.Picture = CopyPic.Image
CopyPic.PaintPicture StorePic.Picture, 0, CopyPic.ScaleHeight, _
 CopyPic.ScaleWidth, -CopyPic.ScaleHeight
CopyPic.Picture = CopyPic.Image

StorePic.PaintPicture StorePic.Picture, 0, StorePic.ScaleHeight, _
 StorePic.ScaleWidth, -StorePic.ScaleHeight
StorePic.Picture = StorePic.Image
End Sub

Sub rtbd_Click()
CusD.Show vbModal
End Sub

Sub rt180_Click()
'CopyPic.PaintPicture StorePic.Picture, CopyPic.ScaleWidth, CopyPic.ScaleHeight, _
 -CopyPic.ScaleWidth, -CopyPic.ScaleHeight
'CopyPic.Picture = CopyPic.Image

'StorePic.PaintPicture StorePic.Picture, StorePic.ScaleWidth, StorePic.ScaleHeight, _
 -StorePic.ScaleWidth, -StorePic.ScaleHeight
'StorePic.Picture = StorePic.Image
Select Case RDEG
Case 0
CopyPic.Width = StorePic.Width
CopyPic.Height = StorePic.Height
Case 180
CopyPic.Width = StorePic.Width
CopyPic.Height = StorePic.Height
Case 90
CopyPic.Width = StorePic.Height
CopyPic.Height = StorePic.Width
Case 270
CopyPic.Width = StorePic.Height
CopyPic.Height = StorePic.Width
Case Else
Dim g As Double
g = LineLen(StorePic.Left, StorePic.Top, StorePic.Left + StorePic.Width, StorePic.Top + StorePic.Height)
CopyPic.Width = g
CopyPic.Height = g
CopyPic.Cls
Set CopyPic.Picture = Nothing
bmp_rotate StorePic, CopyPic, RDEG * Trans
RDEG = 270
StorePic.Picture = CopyPic.Image
Exit Sub
End Select
CopyPic.Cls
bmp_rotate StorePic, CopyPic, RDEG * Trans
RDEG = RDEG - 90
If RDEG = -90 Then RDEG = 270

End Sub


Private Sub Form_Load()
Dim o As Integer, I As Integer

File1.Path = App.Path & "\Brushes\"
File1.Selected(0) = True

MainPic.MouseIcon = LoadPicture(App.Path & "\pen_r.cur")
MainPic.Top = 0
MainPic.Left = 0

KSize.Width = 500
KSize.Height = 400

With MainPic
    .Width = KSize.Width
    .Height = KSize.Height
End With

ZIm.Stretch = True
RDEG = 270
PI = Atn(1) * 4
Trans = PI / 180
BackC = BCP.BackColor
SaveDr = 1
ZZ.AddItem 2
ZZ.AddItem 4
ZZ.AddItem 8
ZZ.AddItem 12
ZZ.ListIndex = 0
Z = ZZ.Text

For o = 1 To 5
    Si(o).Top = (PicSize.ScaleHeight / 2) - (Si(o).Height / 2)
Next o
ErSize = 1
Sh = 1

For I = 1 To 100
    CDr.AddItem I
Next I
CDr.ListIndex = 0

For I = 5 To 200 Step 5
    SprayThik.AddItem I
Next I
SprayThik.ListIndex = 15

For I = 5 To 50 Step 5
    SprayRad.AddItem I
Next I
SprayRad.ListIndex = 3
Form_Resize

For I = 1 To 30
    Load MemoryPic(I)
    MemoryPic(I).Left = MemoryPic(0).Left
    MemoryPic(I).Top = MemoryPic(0).Top
Next I
MemoryPic(1).Picture = MainPic.Image
MemI = 1
MainPic.Refresh

ZoomUpdate
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 99
If x > MainPic.Left + MainPic.Width And x < MainPic.Left + MainPic.Width + 11 Then
    If y < MainPic.ScaleHeight Then
        Me.MousePointer = 9
    End If
End If
If y > MainPic.Top + MainPic.Height And y < MainPic.Top + MainPic.Height + 11 Then
    If x < MainPic.ScaleWidth Then
        Me.MousePointer = 7
    End If
End If
If x > MainPic.Left + MainPic.Width And x < MainPic.Left + MainPic.Width + 11 Then
    If y > MainPic.Top + MainPic.Height And y < MainPic.Top + MainPic.Height + 11 Then
        Me.MousePointer = 8
    End If
End If
If Button = 1 Then
    If x > MainPic.Left + MainPic.Width - 31 And x < MainPic.Left + MainPic.Width + 31 Then
        If y < MainPic.ScaleHeight Then
            Me.MousePointer = 9
            KanvasChange x
        End If
    End If
    If y > MainPic.Top + MainPic.Height - 31 And y < MainPic.Top + MainPic.Height + 31 Then
        If x < MainPic.ScaleWidth Then
            Me.MousePointer = 7
            KanvasChange 0, y
        End If
    End If
    If x > MainPic.Left + MainPic.Width - 31 And x < MainPic.Left + MainPic.Width + 31 Then
        If y > MainPic.Top + MainPic.Height - 31 And y < MainPic.Top + MainPic.Height + 31 Then
            Me.MousePointer = 8
            KanvasChange x, y
        End If
    End If
End If
End Sub

Sub KanvasChange(Optional ByVal aX As Long = 0, Optional ByVal aY As Long = 0)
If aX <> 0 Then
    If aX > 0 Then
        KSize.Width = Abs(MainPic.Left) + aX
    End If
End If
If aY <> 0 Then
    If aY > 0 Then
        KSize.Height = Abs(MainPic.Top) + aY
    End If
End If
MainPic.Width = KSize.Width
MainPic.Height = KSize.Height
StBar.SimpleText = "Canvas Size " & KSize.Width & "x" & KSize.Height
'Form_Resize
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_Resize
ZoomUpdate
ZoomUpdate
End Sub

Private Sub Form_Resize()
On Error Resume Next
'MainPic.Left = 0
'MainPic.Top = 0
VSc.Top = 0
VSc.Left = Me.ScaleWidth - MainToolBox.ScaleWidth - VSc.Width - 1
VSc.Height = Me.ScaleHeight - StBar.Height - HSc.Height
HSc.Left = 0
HSc.Top = Me.ScaleHeight - StBar.Height - HSc.Height
HSc.Width = Me.ScaleWidth - MainToolBox.ScaleWidth - VSc.Width
VSc.Min = 0
'VSc.Max = MainPic.ScaleHeight + MainPic.Top
VSc.Max = MainPic.ScaleHeight - HSc.Top + IIf(HSc.Visible, HSc.Height, 0)
HSc.Min = 0
'HSc.Max = MainPic.ScaleWidth + MainPic.Left
HSc.Max = MainPic.ScaleWidth - VSc.Left + IIf(VSc.Visible, VSc.Width, 0)

If (MainPic.Left + MainPic.Width) + HSc.Value > VSc.Left - 5 Then
    HSc.Visible = True
Else
    HSc.Visible = False
    MainPic.Left = 0
    HSc.Value = 0
End If

If (MainPic.Top + MainPic.Height) + VSc.Value > HSc.Top - 5 Then
    VSc.Visible = True
Else
    VSc.Visible = False
    MainPic.Top = 0
    VSc.Value = 0
End If

If HSc.Visible = True And VSc.Visible = True Then
    ClGap.Move HSc.Left + HSc.Width, VSc.Top + VSc.Height, VSc.Width, HSc.Height
    ClGap.ToolTipText = "Size " & MainPic.ScaleWidth & "x" & MainPic.ScaleHeight
    ClGap.Visible = True
Else
    ClGap.Visible = False
End If

'ZoomUpdate
End Sub

Private Sub StBr_Click()
SavToBr_Click
End Sub

Private Sub VSc_SCROLL()
VSc_Change
End Sub

Private Sub HSc_SCROLL()
HSc_Change
End Sub

Private Sub VSc_Change()
MainPic.Top = -VSc.Value
Form_Resize
'ZoomUpdate
End Sub

Private Sub HSc_Change()
MainPic.Left = -HSc.Value
Form_Resize
'ZoomUpdate
End Sub

Private Sub KanvaSize_Click()
KanD.texW.Text = KSize.Width
KanD.texH.Text = KSize.Height
KanD.Show vbModal
End Sub

Private Sub Loa_Click()
Dim TempF As String

OkToDraw = False
MainPic.AutoSize = True
CD1.Filter = "Pictures (*.bmp;*.ico;*.jpeg;*.jpg)|*.bmp;*.ico;*.jpeg;*.jpg"
CD1.ShowOpen
TempF = CD1.FileName
If TempF = "" Then MainPic.AutoSize = False: Exit Sub

MainPic.Picture = LoadPicture(TempF)
MainPic.Refresh
UndRed
Form_Resize
ZoomUpdate
MainPic.AutoSize = False
KSize.Width = MainPic.ScaleWidth
KSize.Height = MainPic.ScaleHeight
End Sub

Private Sub Sav_Click()
Dim TempF As String
CD1.Filter = "Pictures (*.bmp)|*.bmp"
CD1.ShowSave
TempF = CD1.FileName
If TempF = "" Then Exit Sub

SavePicture MainPic.Image, TempF
End Sub

Private Sub MainPic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Sq1.Visible = True Then DelSelect
If KeyCode = 46 And CopyPic.Visible = True Then
    CopyPic.Visible = False
    OKk = False
End If
End Sub

Private Sub SelAll_Click()
If OKk = True Then
    CopyThePic
End If
Toolbar1.Buttons(Sh).Value = tbrUnpressed
Sh = 5
Toolbar1.Buttons(5).Value = tbrPressed
Toolbar1_ButtonClick Toolbar1.Buttons(5)
'MainPic.MousePointer = 99
'MainPic.MouseIcon = LoadPicture(App.Path & "\3dwarro.cur")


MainPic_MouseDown 1, 0, 0, 0
MainPic_MouseMove 1, 0, MainPic.ScaleWidth, MainPic.ScaleHeight
MainPic_MouseUp 1, 0, MainPic.ScaleWidth, MainPic.ScaleHeight

End Sub

Private Sub MainPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
OkToDraw = True
Sq1.Visible = False
If MainPic.MousePointer = 99 Then CopyPic.Visible = False
If Button = 1 Then
Select Case Sh
Case 1
MainPic.DrawWidth = CDr.List(CDr.ListIndex)
MainPic.PSet (x, y), ForeC
XX = x
YY = y
Case 2
XX = x
YY = y
Line1.X1 = x
Line1.Y1 = y
Line1.X2 = x
Line1.Y2 = y
Line1.BorderColor = ForeC
Line1.Visible = True
Case 3
Sq1.BorderColor = ForeC
XX = x
YY = y
Sq1.Top = y
Sq1.Left = x
Sq1.Width = 0
Sq1.Height = 0
Sq1.Visible = True
Case 4
Circ.BorderColor = ForeC
XX = x
YY = y
Circ.Visible = True
DrawCir x, y
Case 5
If OKk = True And MainPic.MousePointer = 99 Then
    StBr.Enabled = False
    CopyThePic
    RDEG = 270
    Set StorePic.Picture = Nothing
    StorePic.Picture = StorePic.Image
    Exit Sub
End If
StBr.Enabled = True
Sq1.BorderColor = 0
XX = x
YY = y
Sq1.Top = y
Sq1.Left = x
Sq1.Width = 0
Sq1.Height = 0
Sq1.Visible = True
Case 6
MainPic.Line (x, y)-((x + DelL) - 1, (y + DelL) - 1), MainPic.BackColor, BF
Case 7
FillArea x, y
Case 8
Oval.BorderColor = ForeC
XX = x
YY = y
Oval.Visible = True
DrawOval x, y
Case 9
XX = x
YY = y
SpT.Enabled = True
Case 10
FCP.BackColor = MainPic.Point(x, y)
ForeC = FCP.BackColor
Case 11
Sq1.BorderColor = 0
XX = x
YY = y
Sq1.Top = y
Sq1.Left = x
Sq1.Width = 0
Sq1.Height = 0
Sq1.Visible = True
Case 12
TransparentBlt MainPic.hDC, x - BrusPic.ScaleWidth / 2, y - BrusPic.ScaleHeight / 2, BrusPic.ScaleWidth, BrusPic.ScaleHeight, BrusPic.hDC, 0, 0, BrusPic.ScaleWidth, BrusPic.ScaleHeight, BrusPic.BackColor
MainPic.Refresh
End Select
End If
End Sub

Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 And OkToDraw = True Then
Select Case Sh
Case 1
Pencill x, y
StBar.SimpleText = "X " & x & ", Y " & y
Case 2
Line1.X2 = x
Line1.Y2 = y
StBar.SimpleText = "X " & x & ", Y " & y & "   Line Length " & Int(LineLen(XX, YY, x, y))
Case 3
DrawSelection x, y
StBar.SimpleText = "X " & x & ", Y " & y & "   Rect Size " & Sq1.Width & " x " & Sq1.Height
Case 4
DrawCir x, y
StBar.SimpleText = "X " & x & ", Y " & y & "   Rad Size " & Int(Radd)
Case 5
If CopyPic.Visible = False Then
    DrawSelection x, y
Else
    With Rsz
        .x = x
        .y = y
        .Button = Button
        .Resize CopyPic, MainPic
        StBar.SimpleText = .TextStr
    End With
On Error GoTo bye
CopyPic.PaintPicture StorePic.Picture, 0, 0, _
 CopyPic.ScaleWidth, CopyPic.ScaleHeight
'CopyPic.Picture = CopyPic.Image
End If
Case 6
MainPic.Line (x, y)-((x + DelL) - 1, (y + DelL) - 1), MainPic.BackColor, BF
Case 7
'
Case 8
DrawOval x, y
Case 9
XX = x
YY = y
Case 10
FCP.BackColor = MainPic.Point(x, y)
ForeC = FCP.BackColor
Case 11
DrawSelection x, y
StBar.SimpleText = "X " & x & ", Y " & y & "   Rect Size " & Sq1.Width & " x " & Sq1.Height
Case 12
TransparentBlt MainPic.hDC, x - BrusPic.ScaleWidth / 2, y - BrusPic.ScaleHeight / 2, BrusPic.ScaleWidth, BrusPic.ScaleHeight, BrusPic.hDC, 0, 0, BrusPic.ScaleWidth, BrusPic.ScaleHeight, BrusPic.BackColor
MainPic.Refresh
End Select
Exit Sub
End If
If Button = 2 And Sh = 10 Then
BCP.BackColor = MainPic.Point(x, y)
BackC = BCP.BackColor
End If
StBar.SimpleText = "X " & x & ", Y " & y
ZoomUpdate
MakeZoom x, y
If CopyPic.Visible Then
    With Rsz
        .x = x
        .y = y
        .Button = Button
        .Resize CopyPic, MainPic
    End With
End If
bye:
End Sub

Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And OkToDraw = True Then
Select Case Sh
Case 1
'Pencill X, Y
Case 2
MainPic.DrawWidth = CDr.Text
MainPic.Line (XX, YY)-(x, y), ForeC
Line1.Visible = False
Case 3
MainPic.DrawWidth = CDr.Text
Sq1.Visible = False
MainPic.Line (XX, YY)-(x - 1, y - 1), ForeC, B
Case 4
Circ.Visible = False
MainPic.DrawWidth = CDr.Text
MainPic.Circle (XX, YY), Radd, ForeC ', , , 2
Case 5
If MainPic.MousePointer = 99 Then
    DrawSelection x, y
    DrawCP XX, YY
    StorePic.Picture = CopyPic.Image
    'StorePic.Picture = StorePic.Image
End If
Case 6
'
Case 7
'
Case 8
MainPic.DrawWidth = CDr.Text
Oval.Visible = False
MainPic.Circle (XX, YY), IIf(Oval.Width > Oval.Height, (Oval.Width / 2), (Oval.Height / 2)), ForeC, , , (Oval.Height) / (Oval.Width)
Case 9
SpT.Enabled = False
Case 10
FCP.BackColor = MainPic.Point(x, y)
ForeC = FCP.BackColor
Case 11
DrawSelection x, y
DrawGradient
Case 12
'
End Select
End If
ZoomUpdate
UndRed
End Sub

Sub CopyThePic()
On Error GoTo bye
If TransC.Value = 0 Then
    MainPic.PaintPicture CopyPic.Picture, CopyPic.Left + 1, CopyPic.Top + 1, _
     CopyPic.ScaleWidth, CopyPic.ScaleHeight
    CopyPic.Visible = False
Else
    TransparentBlt MainPic.hDC, CopyPic.Left + 1, CopyPic.Top + 1, CopyPic.ScaleWidth, CopyPic.ScaleHeight, CopyPic.hDC, 0, 0, CopyPic.ScaleWidth, CopyPic.ScaleHeight, CopyPic.BackColor
End If

bye:
End Sub

Sub DrawCP(aX As Single, aY As Single)
On Error GoTo bye
If OKk = True Then OKk = False: Exit Sub
Set CopyPic.Picture = Nothing
CopyPic.Left = Sq1.Left
CopyPic.Top = Sq1.Top
CopyPic.Width = Sq1.Width
CopyPic.Height = Sq1.Height
If Cop.Value = 0 Then
    MainPic.DrawWidth = 1
    MainPic.Line (Sq1.Left + 1, Sq1.Top + 1)-(Sq1.Left + Sq1.Width - 2, Sq1.Top + Sq1.Height - 2), MainPic.BackColor, BF
    MainPic.DrawWidth = CDr.Text
End If

CopyPic.PaintPicture MainPic.Picture, -1, -1, CopyPic.ScaleWidth + 1, CopyPic.ScaleHeight + 1, _
CopyPic.Left, CopyPic.Top, CopyPic.ScaleWidth + 1, CopyPic.ScaleHeight + 1

CopyPic.Visible = True
Sq1.Visible = False
CopyPic.Picture = CopyPic.Image
OKk = True
Exit Sub
bye:
Sq1.Visible = False
OKk = False
End Sub

Private Sub CopyPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    With Rsz
        .SetStartUpPosX = x
        .SetStartUpPosY = y
    End With
End If

If Button = 2 Then
    Screen.ActiveForm.PopupMenu Hidd1
End If
End Sub

Private Sub CopyPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    With Rsz
        .MovePosX = x
        .MovePosY = y
        .MoveTheObject CopyPic
    End With
End If
End Sub

Public Sub ZoomUpdate()
MainPic.AutoRedraw = False

With ZIm
    .Width = Int(MainPic.ScaleWidth * Z)
    .Height = Int(MainPic.ScaleHeight * Z)
End With

MIDL.Move Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, Z, Z
MainPic.Picture = MainPic.Image
ZIm.Picture = MainPic.Image
MainPic.AutoRedraw = True
End Sub

Sub MakeZoom(ByVal aX, ByVal aY)
ZIm.Move (Picture1.ScaleWidth / 2) - (aX * Z), (Picture1.ScaleHeight / 2) - (aY * Z)
'x1 = (Picture1.ScaleWidth / 2) - (aX * Z)
'y1 = (Picture1.ScaleHeight / 2) - (aY * Z)
'ZIm.Move x1, y1
End Sub

Private Sub ZZ_Click()
Z = Val(ZZ.Text)
ZoomUpdate
End Sub

Sub UndRed()
If MemI > 29 Then
    Dim I
    For I = 1 To 29
        MemoryPic(I).Picture = MemoryPic(I + 1).Picture
    Next I
    MemoryPic(30).Picture = MainPic.Image
Else
    MemI = MemI + 1
    MemoryPic(MemI).Picture = MainPic.Image
End If
UNd.Enabled = True
REDo.Enabled = False
CurPic = MemI
End Sub

Private Sub REDo_Click()
If MemI < 30 Then
    MemI = MemI + 1
    MainPic.Picture = MemoryPic(MemI).Picture
    If MemI = 30 Or MemI = CurPic Then REDo.Enabled = False
    UNd.Enabled = True
    ZoomUpdate
End If
End Sub

Private Sub UNd_Click()
If MemI > 1 Then
    MemI = MemI - 1
    MainPic.Picture = MemoryPic(MemI).Picture
    If MemI = 1 Then UNd.Enabled = False
    REDo.Enabled = True
    ZoomUpdate
End If
End Sub

Sub DrawSelection(aX, aY)
Sq1.Left = IIf(aX > XX, XX, aX)
Sq1.Top = IIf(aY > YY, YY, aY)
Sq1.Width = Abs(aX - XX)
Sq1.Height = Abs(aY - YY)
End Sub

Sub DrawSelection1(aX, aY)
Circ.Left = IIf(aX > XX, XX, aX)
Circ.Top = IIf(aY > YY, YY, aY)
Circ.Width = Abs(aX - XX)
Circ.Height = Abs(aY - YY)
End Sub

Sub Pencill(ByVal aX As Single, ByVal aY As Single)
MainPic.Line (aX, aY)-(XX, YY), ForeC
XX = aX
YY = aY
End Sub

Private Sub MainToolBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bye
If Button = 1 Then
    FCP.BackColor = MainToolBox1.Point(x, y)
    ForeC = FCP.BackColor
End If
If Button = 2 Then
    BCP.BackColor = MainToolBox1.Point(x, y)
    BackC = BCP.BackColor
End If
bye:
End Sub

Private Sub MainToolBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo bye
If Button = 1 Then
    FCP.BackColor = MainToolBox1.Point(x, y)
    ForeC = FCP.BackColor
End If
If Button = 2 Then
    BCP.BackColor = MainToolBox1.Point(x, y)
    BackC = BCP.BackColor
End If
bye:
End Sub

Private Sub Si_Click(Index As Integer)
Dim o
For o = 1 To 5
    Si(o).BackColor = &H80000005
Next o
Si(Index).BackColor = &H80000002
ErSize = Index
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\er" & ErSize & ".cur")
Select Case ErSize
Case 1
DelL = 4
Case 2
DelL = 6
Case 3
DelL = 12
Case 4
DelL = 24
Case 5
DelL = 32
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
OKk = False
StBr.Enabled = False
PrvIm.Visible = False
File1.Visible = False
Cop.Visible = False
TransC.Visible = False
Sq1.Visible = False
CopyPic.Visible = False
Set CopyPic.Picture = Nothing
Select Case Button.Index
Case 1
Sh = 1
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\pen_r.cur")
Case 2
Sh = 2
MainPic.MousePointer = 2
Case 3
Sh = 3
MainPic.MousePointer = 2
Sq1.BorderStyle = 1
Sq1.BorderColor = ForeC
Case 4
Sh = 4
MainPic.MousePointer = 2
Circ.BorderColor = ForeC
'Circ.Visible = True
Case 5
Sh = 5
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\3dwarro.cur")
Sq1.BorderStyle = 3
Sq1.BorderColor = 0
Cop.Visible = True
TransC.Visible = True
Case 6
Sh = 6
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\er" & ErSize & ".cur")
PicSize.Visible = True
MainPic.DrawWidth = 1
SaveDr = CDr
CDr.Visible = False
Label1.Visible = False
DegPic.Visible = False
SprayPic.Visible = False
Exit Sub
Case 7
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\fillc.cur")
Sh = 7
Case 8
MainPic.MousePointer = 2
Sh = 8
Case 9
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\spray.cur")
SprayPic.Visible = True
MainPic.DrawWidth = 1
SaveDr = CDr
CDr.Visible = False
Label1.Visible = False
PicSize.Visible = False
DegPic.Visible = False
Sh = 9
Exit Sub
Case 10
MainPic.MousePointer = 99
MainPic.MouseIcon = LoadPicture(App.Path & "\pickcol.cur")
Sh = 10
Case 11
MainPic.MousePointer = 2
Sq1.BorderStyle = 1
Sq1.BorderColor = 0
Sh = 11
DegPic.Visible = True
CDr.Visible = False
Label1.Visible = False
PicSize.Visible = False
SprayPic.Visible = False
Exit Sub
Case 12
Sh = 12
MainPic.MousePointer = 2
PrvIm.Visible = True
File1.Visible = True
CDr.Visible = False
Exit Sub
End Select
DegPic.Visible = False
CDr.Visible = True
Label1.Visible = True
MainPic.DrawWidth = SaveDr
PicSize.Visible = False
SprayPic.Visible = False
End Sub

Sub FillArea(aX As Single, aY As Single)
MainPic.FillStyle = 0
Randomize
MainPic.FillColor = ForeC

'API call
ExtFloodFill MainPic.hDC, aX, aY, MainPic.Point(aX, aY), 1

MainPic.FillStyle = 1
End Sub

Sub DrawCir(aX, aY)
Circ.Visible = True
Radd = LineLen(XX, YY, aX, aY)
Circ.Width = Radd * 2
Circ.Height = Radd * 2
Circ.Left = XX - Circ.Width / 2
Circ.Top = YY - Circ.Height / 2
End Sub

Function LineLen(ByVal CpX, ByVal CpY, ByVal CurX, ByVal CurY) As Double
LineLen = Sqr(((Abs(CpY - CurY)) ^ 2) + ((Abs(CpX - CurX)) ^ 2))
End Function

Sub DrawOval(aX, aY)
On Error Resume Next
Oval.Visible = True
Radd = LineLen(XX, YY, aX, aY)
Oval.Width = (Abs(XX - aX) * 2)
Oval.Height = (Abs(YY - aY) * 2)
Oval.Left = XX - Oval.Width / 2
Oval.Top = YY - Oval.Height / 2
End Sub

Sub DrawSpray()
WORKING = True
DoEvents
Dim F As Integer
Dim DEG As Integer
Randomize
For F = 1 To SprayThik
    aX = XX
    aY = YY
    DIS = Int(Rnd * SprayRad)
    If DIS < 11 And DIS > 4 Then
        If Int(Rnd * 3) < 2 Then
            GoTo ff
        End If
    End If
        If DIS < 5 Then
        If Int(Rnd * 8) < 7 Then
            GoTo ff
        End If
    End If
    DEG = Int(Rnd * 361)
    aX = aX + DIS * Cos(DEG * PI / 180)
    aY = aY + DIS * Sin(DEG * PI / 180)
    MainPic.PSet (aX, aY), ForeC
ff:
Next F
WORKING = False
End Sub

Private Sub SpT_Timer()
If WORKING = False Then DrawSpray
End Sub

Sub DelSelect()
MainPic.Line (Sq1.Left, Sq1.Top)-(Sq1.Left + Sq1.Width, Sq1.Top + Sq1.Height), MainPic.BackColor, BF
End Sub

Private Sub DrawGradient()
If IsNumeric(Text1.Text) = False Then Exit Sub

GrPic.Left = Sq1.Left
GrPic.Top = Sq1.Top
GrPic.Width = Sq1.Width
GrPic.Height = Sq1.Height
Sq1.Visible = False
GrPic.Visible = True
GrPic.Refresh

With mGradient
    .Angle = Text1.Text
    .Color1 = ForeC 'mlColor1
    .Color2 = BackC 'mlColor2
    .Draw GrPic
End With
GrPic.Refresh

TransparentBlt MainPic.hDC, GrPic.Left, GrPic.Top, GrPic.ScaleWidth, GrPic.ScaleHeight, GrPic.hDC, 0, 0, GrPic.ScaleWidth, GrPic.ScaleHeight, MainPic.BackColor
MainPic.Refresh

GrPic.Visible = False
    
End Sub

Private Sub zzz_Click()
If ZPIC.Visible Then
    ZPIC.Visible = False
    zzz.Checked = False
Else
    ZPIC.Visible = True
    zzz.Checked = True
End If
End Sub
