VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H8000000B&
   Caption         =   "REPORTES GENERALES"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
   LinkTopic       =   "Form9"
   ScaleHeight     =   7575
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "REGRESAR MENU"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GENERAR REPORTE"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   10815
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "REPORTES GENERALES"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim line As String
    Dim datos() As String

    List1.Clear

    Open App.Path & "\" & "Productos.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        datos = Split(line, ";")

        List1.AddItem "Código de Barras: " & datos(0)
        List1.AddItem "Nombre del Producto: " & datos(1)
        List1.AddItem "Existencias: " & datos(2)
        List1.AddItem "Precio Unitario: " & datos(3)
        List1.AddItem "------------------------------------"
    Loop
    Close #1

    MsgBox "REPORTE GENERADO CON ÉXITO."
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
