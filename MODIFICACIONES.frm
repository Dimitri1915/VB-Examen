VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H8000000B&
   Caption         =   "MODIFICACIONES DE PRODUCTO"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form6"
   ScaleHeight     =   7350
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "REGRESAR AL MENU"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      TabIndex        =   13
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LIMPIAR CAMPOS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   12
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MODIFICAR PRODUCTOS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   11
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUSCAR PRODUCTO"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   16
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000B&
      Caption         =   "PRECIO UNITARIO"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   14
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   9
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "EXISTENCIAS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "C�DIGO DE BARRAS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "MODIFICACI�N DE PRODUCTOS POR MEDIO DE C�DIGO DE BARRAS."
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form6"
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
            Label3.Caption = datos(0)
            Label5.Caption = datos(1)
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
    Dim line As String
    Dim productos() As String
    Dim codigoBuscado As String
    Dim encontrado As Boolean
    Dim i As Integer
    Dim datos() As String

    codigoBuscado = Text1.Text
    encontrado = False
    i = 0

    ReDim productos(0)

    Open App.Path & "\" & "Productos.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        datos = Split(line, ";")
        
        i = i + 1
 
        ReDim Preserve productos(i)
        
        If Trim(datos(0)) = Trim(codigoBuscado) Then
            productos(i) = Text2.Text & ";" & Text3.Text & ";" & Text4.Text & ";" & Text5.Text
            encontrado = True
        Else
            productos(i) = line
        End If
    Loop
    Close #1

    If encontrado Then
        Open App.Path & "\" & "Productos.dat" For Output As #1
        For i = 1 To UBound(productos)
            Print #1, productos(i)
        Next i
        Close #1
        MsgBox "PRODUCTO MODIFICADO CON �XITO."
    Else
        MsgBox "PRODUCTO NO ENCONTRADO."
    End If

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Label3.Caption = ""
    Label5.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Text1.SetFocus
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Label3.Caption = ""
Label5.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
