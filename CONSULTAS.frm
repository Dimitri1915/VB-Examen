VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000B&
   Caption         =   "CONSULTAS DE PRODUCTO"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form5"
   ScaleHeight     =   7110
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR AL MENU"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NUEVA CONSULTA"
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
      Left            =   360
      TabIndex        =   9
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000B&
      Caption         =   "PRECIO UNITARIO"
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
      Left            =   360
      TabIndex        =   12
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label8 
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
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label7 
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
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   4200
      Width           =   4095
   End
   Begin VB.Label Label6 
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
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label Label5 
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
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "EXISTENCIAS:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "PRODUCTO:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "CODIGO DE BARRAS:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "CONSULTA POR MEDIO DE CODIGO DE BARRAS."
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim line As String
    Dim codigoBuscado As String
    Dim encontrado As Boolean
    Dim datos() As String
    
    codigoBuscado = Text1.Text
    encontrado = False
    
    Open App.Path & "\" & "Productos.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        datos = Split(line, ";")

        If Trim(datos(0)) = Trim(codigoBuscado) Then
            Label5.Caption = datos(0)
            Label6.Caption = datos(1)
            Label7.Caption = datos(2)
            Label8.Caption = datos(3)
            encontrado = True
            Exit Do
        End If
    Loop
    Close #1

    If Not encontrado Then
        MsgBox "PRODUCTO NO ENCONTRADO."
    End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

