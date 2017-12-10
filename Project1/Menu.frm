VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H000000FF&
   Caption         =   "Menu"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Inscripcion"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prestamos"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Menu.Hide
Prestamos.Show
End Sub

Private Sub Command2_Click()
Menu.Hide
Inscripcion.Show
End Sub

Private Sub Command3_Click()
End
End Sub
