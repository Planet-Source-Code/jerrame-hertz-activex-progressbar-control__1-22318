VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3510
   ClientLeft      =   660
   ClientTop       =   600
   ClientWidth     =   2190
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2190
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   735
      Begin VB.VScrollBar VScroll1 
         Height          =   2775
         Left            =   240
         Max             =   50
         TabIndex        =   3
         Top             =   360
         Value           =   50
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ProgressBar"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      Begin ProgressBar.Meter Meter1 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   4895
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then Call FormDrag(Me)
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    End
    
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then Call FormDrag(Me)

End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    End
    
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then Call FormDrag(Me)
    
End Sub

Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    End
    
End Sub

Private Sub VScroll1_Change()

    SetVal

End Sub

Private Sub VScroll1_Scroll()

    SetVal

End Sub

Function SetVal()
    
    Meter1.Value = VScroll1.Value
    
End Function
