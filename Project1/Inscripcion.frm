VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Inscripcion 
   BackColor       =   &H00C00000&
   Caption         =   "Incripcion"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   6465
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\programa\Inscripcion.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\programa\Inscripcion.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Estudiantes"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar Matricula"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      DataField       =   "Periodo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "Seccion"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "Materia"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Matricula"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Matricula"
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
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
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
      Left            =   600
      TabIndex        =   4
      Top             =   4800
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Seccion"
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
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Carrera"
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
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
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
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
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
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   840
   End
End
Attribute VB_Name = "Inscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ciclo As Integer
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Text1.SetFocus
MsgBox ("GUARDADO EXITOSAMENTE")
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command3_Click()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Command6.Visible = True
Text7.Visible = True
End Sub

Private Sub Command4_Click()
Inscripcion.Hide
Menu.Show
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
ciclo = 0
While (ciclo <> 1)
    If (Text7.Text = Text1.Text) Then
    ciclo = 1
    MsgBox ("Busqueda Exitosa")
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Command6.Visible = False
    Text7.Visible = False
    Text7.Text = ""
    Else
     If Adodc1.Recordset.EOF Then
     ciclo = 1
     MsgBox ("Matricula No Encontrada")
    Adodc1.Recordset.MoveFirst
    Command6.Visible = True
    Text7.Visible = True
    Text7.SetFocus
     Else
     Adodc1.Recordset.MoveNext
     End If
    End If
    Wend
    

End Sub
