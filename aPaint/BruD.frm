VERSION 5.00
Begin VB.Form BruD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Brush"
   ClientHeight    =   1155
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox BrName 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Brush Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "BruD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
Dim SB As FileData
If BrName.Text <> "" Then
    SavePicture Form1.CopyPic.Picture, App.Path & "\Brushes\" & Me.BrName.Text & ".BMP"
    Form1.File1.Refresh
    Unload Me
Else
    MsgBox "Enter a name...", vbExclamation
End If
End Sub
