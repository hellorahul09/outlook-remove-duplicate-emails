VERSION 1.0 CLASS
Begin VB.Form ProgressForm 
   Caption         =   "Progress"
   ClientHeight    =   1000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4000
   StartUpPosition =   2  ' CenterScreen
   Begin VB.Label ProgressLabel 
      Caption         =   "Processing: 0%"
      Height          =   375
      Left            =   150
      Top             =   300
      Width           =   3700
   End
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    ' Initialize the progress label
    Me.ProgressLabel.Caption = "Processing: 0%"
End Sub

Public Sub UpdateProgress(Current As Long, Total As Long)
    ' Update the progress percentage
    Dim Percentage As Double
    Percentage = (Current / Total) * 100
    Me.ProgressLabel.Caption = "Processing: " & Format(Percentage, "0") & "%"
    Me.Repaint
End Sub
