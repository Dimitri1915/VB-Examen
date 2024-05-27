VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   Caption         =   "SISTEMA PARA TIENDAS DE ABARROTES"
   ClientHeight    =   2970
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "BIENVENIDO A SU SISTEMA DE TIENDA DE ABARROTES, ELIGA LA OPCIÓN ARRIBA EN EL MENÚ."
      BeginProperty Font 
         Name            =   "Cascadia Code SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   2650
      Left            =   120
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2650
   End
   Begin VB.Menu menu 
      Caption         =   "MENU"
      Begin VB.Menu salir 
         Caption         =   "SALIR"
      End
   End
   Begin VB.Menu altas 
      Caption         =   "ALTAS"
   End
   Begin VB.Menu bajas 
      Caption         =   "BAJAS"
   End
   Begin VB.Menu consultas 
      Caption         =   "CONSULTAS"
   End
   Begin VB.Menu modificaciones 
      Caption         =   "MODIFICACIONES DE PRODUCTO"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub salir_Click()
MsgBox "Gracias por utilizar este sistema. Hasta la proxima."
End
End Sub
