VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000B&
   Caption         =   "BAJAS DE PRODUCTOS"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form4"
   ScaleHeight     =   6975
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "LIMPIAR CAMPOS"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REVISAR PRODUCTO"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Height          =   1095
      Left            =   4560
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAR DE BAJA PRODUCTO"
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
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
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label3 
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
      Height          =   975
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   6015
   End
   Begin VB.Label Label2 
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
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "DIGITE EL CÓDIGO DE BARRAS"
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
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim line As String
    Dim productos() As String
    Dim codigoABorrar As String
    Dim encontrado As Boolean
    Dim i As Integer

    codigoABorrar = Text1.Text
    encontrado = False
    i = 0
    
    Open App.Path & "\" & "Productos.dat" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        If Trim(Split(line, ";")(0)) <> Trim(codigoABorrar) Then
            i = i + 1
            ReDim Preserve productos(1 To i)
            productos(i) = line
        Else
            encontrado = True
        End If
    Loop
    Close #1

    If encontrado Then
        Open App.Path & "\" & "Productos.dat" For Output As #1
        If i > 0 Then
            For i = 1 To UBound(productos)
                Print #1, productos(i)
            Next i
        End If
        Close #1
        MsgBox "PRODUCTO ELIMINADO CON ÉXITO."
    Else
        MsgBox "PRODUCTO NO ENCONTRADO."
    End If
    
    Text1.Text = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Text1.SetFocus
    
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
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

            Label2.Caption = datos(1)
            Label3.Caption = datos(2)
            encontrado = True
            Exit Do
        End If
    Loop
    Close #1
    
    If Not encontrado Then
        MsgBox "PRODUCTO NO ENCONTRADO."
    End If
    
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Label2.Caption = ""
Label3.Caption = ""
Text1.SetFocus
End Sub

