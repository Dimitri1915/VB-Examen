VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000B&
   Caption         =   "ALTAS DE PRODUCTOS"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   LinkTopic       =   "Form3"
   ScaleHeight     =   7590
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
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
      Left            =   3000
      TabIndex        =   11
      Top             =   4920
      Width           =   3495
   End
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
      Height          =   1095
      Left            =   4560
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox Text2 
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
      Left            =   3000
      TabIndex        =   6
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text3 
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
      Left            =   3000
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
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
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAR DE ALTA PRODUCTO"
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
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label5 
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
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "DAR DE ALTA UN PRODUCTO."
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "EXISTENCIAS"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "NOMBRE DEL PRODUCTO"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "CODIGO DE BARRA"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim codigoBarras As String
    Dim nombreProducto As String
    Dim existencias As Integer
    Dim precioUnitario As Double
    
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "POR FAVOR, COMPLETE TODOS LOS CAMPOS."
        Exit Sub
    End If

    codigoBarras = Text1.Text
    nombreProducto = Text2.Text
    existencias = Val(Text3.Text)
    precioUnitario = Val(Text4.Text)

    Open App.Path + "\" + "Productos.dat" For Append As #1
    Print #1, codigoBarras & ";" & nombreProducto & ";" & CStr(existencias) & ";" & CStr(precioUnitario)
    Close #1

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.SetFocus

    MsgBox "PRODUCTO DADO DE ALTA CORRECTAMENTE."
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

