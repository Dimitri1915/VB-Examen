VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000B&
   Caption         =   "ALTAS DE PRODUCTOS"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form3"
   ScaleHeight     =   6990
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
      Height          =   1095
      Left            =   4560
      TabIndex        =   8
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR TEXTOS"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   5400
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
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   3135
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
      Left            =   3360
      TabIndex        =   5
      Top             =   3960
      Width           =   3135
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
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
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
      Left            =   480
      TabIndex        =   0
      Top             =   5400
      Width           =   1935
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
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Width           =   6015
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
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "NOMBRE"
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
      TabIndex        =   2
      Top             =   2640
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
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
        MsgBox "POR FAVOR, COMPLETE TODOS LOS CAMPOS."
        Exit Sub
    End If
    Open App.Path + "\" + "Productos.dat" For Append As #1
    Print #1, Text1.Text & ";" & Text2.Text & ";" & Text3.Text
    Close #1
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
    MsgBox "PRODUCTO DADO DE ALTA CORRECTAMENTE."
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

