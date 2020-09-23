VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CompCommand, Loading..."
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   ControlBox      =   0   'False
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timLoad 
      Interval        =   10
      Left            =   3075
      Top             =   1950
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   3540
      Left            =   375
      Top             =   375
      Width           =   6615
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CompCommand v1.0 - Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1875
      TabIndex        =   0
      Top             =   3900
      Width           =   5340
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    Open App.Path & "\nointro" For Input As #2
        If Not Err.Number = 53 Then
            Hide
            Err.Clear
        End If
    Close #2
End Sub

Private Sub timLoad_Timer()
    Static StartTime As Double

    If StartTime = 0 Then
        StartTime = Timer
    End If

    nCount = nCount + 1

    PaintPicture Icon, Rnd * (Shape1.Width - Shape1.Left - 32 * 15) + Shape1.Left, Rnd * (Shape1.Height - Shape1.Top - 32 * 15) + Shape1.Top
    
    If Timer - StartTime >= 1.5 Then
        Load frmCommand
        Unload Me
    End If
End Sub
