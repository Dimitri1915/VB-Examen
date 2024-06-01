VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000B&
   Caption         =   "COMPRAS DE PRODUCTOS"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form7"
   ScaleHeight     =   6990
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
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
      Height          =   975
      Left            =   1920
      TabIndex        =   12
      Top             =   4440
      Width           =   3135
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
      Height          =   975
      Left            =   4800
      TabIndex        =   11
      Top             =   5760
      Width           =   1935
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
      Height          =   975
      Left            =   2520
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REGISTRAR COMPRAS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR PRECIO TOTAL"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
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
      Height          =   975
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   4815
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
      Height          =   975
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   4815
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
      Height          =   975
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label5 
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
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "CANTIDAD COMPRADA"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "CODIGO DE BARRAS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "COMPRA DE PRODUCTO NUEVO"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim cantidadComprada As Integer
    Dim precioUnitario As Double
    Dim costoTotal As Double

    cantidadComprada = Val(Text3.Text)
    precioUnitario = Val(Text4.Text)

    costoTotal = cantidadComprada * precioUnitario

    MsgBox "El costo total es: " & FormatCurrency(costoTotal)
End Sub

Private Sub Command2_Click()
    Dim codigoBarras As String
    Dim nombreProducto As String
    Dim cantidadComprada As Integer
    Dim precioUnitario As Double
    Dim costoTotal As Double
    Dim line As String
    Dim productos() As String
    Dim encontrado As Boolean
    Dim i As Integer
    Dim datos() As String
    Dim existenciasActuales As Integer
    Dim precioUnitarioActual As Double

    codigoBarras = Text1.Text
    nombreProducto = Text2.Text
    cantidadComprada = Val(Text3.Text)
    precioUnitario = Val(Text4.Text)

    costoTotal = cantidadComprada * precioUnitario

    encontrado = False
    i = 0
    ReDim productos(1)

    Open App.Path & "\" & "Productos.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        datos = Split(line, ";")

        i = i + 1

        If i > 1 Then
            ReDim Preserve productos(1 To i)
        End If

        If Trim(datos(0)) = Trim(codigoBarras) Then
            existenciasActuales = Val(datos(2))
            precioUnitarioActual = Val(datos(3))
            existenciasActuales = existenciasActuales + cantidadComprada
            precioUnitarioActual = (precioUnitarioActual + precioUnitario) / 2
            productos(i) = datos(0) & ";" & datos(1) & ";" & CStr(existenciasActuales) & ";" & CStr(precioUnitarioActual)
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
        MsgBox "COMPRA REGISTRADA CON ÉXITO. EL COSTO TOTAL ES: " & FormatCurrency(costoTotal)
    Else

        Open App.Path & "\" & "Productos.dat" For Append As #1
        Print #1, codigoBarras & ";" & nombreProducto & ";" & CStr(cantidadComprada) & ";" & CStr(precioUnitario)
        Close #1
        MsgBox "PRODUCTO NO ENCONTRADO. NUEVO PRODUCTO AGREGADO. EL COSTO TOTAL ES: " & FormatCurrency(costoTotal)
    End If

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

