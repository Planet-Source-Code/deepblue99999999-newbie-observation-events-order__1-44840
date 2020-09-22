VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
      MsgBox "Activate"
End Sub

Private Sub Form_Initialize()
      MsgBox "Initialize"
End Sub

Private Sub Form_Load()
      MsgBox "Load"
End Sub

Private Sub Form_Paint()
      MsgBox "Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      MsgBox "QueryUnload"
End Sub

Private Sub Form_Terminate()
      MsgBox "Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
      MsgBox "Unload"
End Sub
