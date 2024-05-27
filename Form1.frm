VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Log in"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   6555
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Cascadia Code"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   720
         TabIndex        =   4
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ENTRAR"
         BeginProperty Font 
            Name            =   "Cascadia Code"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   3
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Cascadia Code"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000B&
         Caption         =   "CONTRASEÑA"
         BeginProperty Font 
            Name            =   "Cascadia Code"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Cascadia Code"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   " TIENDA DE ABARROTES "
         BeginProperty Font 
            Name            =   "Cascadia Code"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   3000
         Width           =   3495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim username As String
Dim password As String

    username = Text1.Text
    password = Text2.Text
        MsgBox "usuario ingresado: " & username
        MsgBox "contraseña ingresada: " & password
If username = "usuario" And password = "contraseña" Then
    MsgBox "¡Bienvenido" & ", " & username & "!"
Form2.Show
Me.Hide
    
Else
    MsgBox "DATOS INCORRECTOS"
Text1.Text = " "
Text2.Text = " "
Text1.SetFocus
End If
End Sub

