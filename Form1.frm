VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengganti SubString dalam Suatu String"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ReplaceAll(SourceString As String, _
ReplaceThis As String, WithThis As String)
    Dim Temp As Variant
    Temp = Split(SourceString, ReplaceThis)
    ReplaceAll = Join(Temp, WithThis)
End Function
  
Private Sub Command1_Click()
    MsgBox ReplaceAll("Rizky Khapidsyah", "iz", "Dessy")
End Sub



