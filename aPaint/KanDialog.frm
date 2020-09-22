VERSION 5.00
Begin VB.Form KanD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canvas Size"
   ClientHeight    =   2070
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox texH 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox texW 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Height:"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Width:"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "KanD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
If (IsNumeric(texW.Text) = True) And (IsNumeric(texH.Text) = True) Then
        KSize.Width = Val(texW.Text)
        KSize.Height = Val(texH.Text)
        Form1.MainPic.Width = KSize.Width
        Form1.MainPic.Height = KSize.Height
        Unload Me
        Form1.ZoomUpdate
Else
MsgBox "Only Numbers"
End If
End Sub
