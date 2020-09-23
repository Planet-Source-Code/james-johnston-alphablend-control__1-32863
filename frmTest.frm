VERSION 5.00
Object = "{9B8B7971-FA33-4584-83EC-9DC92D49BFCE}#23.0#0"; "AlphaBlend.ocx"
Begin VB.Form frmTest 
   Caption         =   "Transparency Control Test"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   1455
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Transparency"
      Height          =   375
      Left            =   1935
      TabIndex        =   2
      Top             =   405
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   450
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Text            =   "128"
      Top             =   405
      Width           =   960
   End
   Begin ctlAlphaBlend.AlphaBlend fo 
      Left            =   4140
      Top             =   90
      _ExtentX        =   767
      _ExtentY        =   767
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "If you set Transparency to 0 you can hit enter when the window has focus to disable the transparency."
      Height          =   465
      Left            =   90
      TabIndex        =   4
      Top             =   945
      Width           =   3885
   End
   Begin VB.Label Label1 
      Caption         =   "Transparency Setting (0 = Transparent, 255 = Opaque)"
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   3975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    fo.Enabled = Check1.Value
End Sub

Private Sub Command1_Click()
    fo.Opacity = CLng(Text1.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fo.Enabled = False
        Check1.Value = False
    End If
End Sub

