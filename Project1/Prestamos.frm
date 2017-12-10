VERSION 5.00
Begin VB.Form Prestamos 
   BackColor       =   &H0000FF00&
   Caption         =   "Prestamos"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Capital mas ganancia interes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   15
      Top             =   5160
      Width           =   3030
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Pago mensual mas interes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   14
      Top             =   4680
      Width           =   2760
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total de interes a ganar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   13
      Top             =   4080
      Width           =   2505
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Interes a ganar cada mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   12
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   2880
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Meses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tasa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "Prestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim interes, pago As Double
Dim cantidad, tasa As Double
Dim meses As Double
Private Sub Command1_Click()
cantidad = Val(Text1.Text)
tasa = Val(Text2.Text)
meses = Val(Text3.Text)
interes = (cantidad * tasa * meses)
Text5.Text = interes
Text4.Text = Val((Text5.Text) / meses)
pago = (cantidad / meses) + Val(Text4.Text)
Text6.Text = pago
Text7.Text = (cantidad + Val(Text5.Text))
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

End Sub

Private Sub Command3_Click()
Dim acep As String
acep = MsgBox("Seguro que deseas Salir", vbOKCancel, vbInformation)
If acep = vbOK Then
Me.Hide
Else
Me.Refresh
End If
End Sub
