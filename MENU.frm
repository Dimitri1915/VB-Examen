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
      Caption         =   "BIENVENIDO A SU SISTEMA DE TIENDA DE ABARROTES, ELIGA LA OPCI�N ARRIBA EN EL MEN�."
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
      Picture         =   "MENU.frx":0000
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
   Begin VB.Menu producto 
      Caption         =   "PRODUCTOS"
      Begin VB.Menu alta 
         Caption         =   "ALTAS"
      End
      Begin VB.Menu baja 
         Caption         =   "BAJAS"
      End
      Begin VB.Menu consulta 
         Caption         =   "CONSULTAS"
      End
      Begin VB.Menu modificacion 
         Caption         =   "MODIFICACIONES"
      End
   End
   Begin VB.Menu entrada 
      Caption         =   "ENTRADAS"
      Begin VB.Menu compras 
         Caption         =   "COMPRAS"
      End
   End
   Begin VB.Menu salida 
      Caption         =   "SALIDAS"
      Begin VB.Menu venta 
         Caption         =   "VENTAS"
      End
   End
   Begin VB.Menu reporte 
      Caption         =   "REPORTES"
      Begin VB.Menu reportegen 
         Caption         =   "REPORTE GENERAL"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
Form3.Show
End Sub

Private Sub baja_Click()
Form4.Show
End Sub

Private Sub compras_Click()
Form7.Show
End Sub

Private Sub consulta_Click()
Form5.Show
End Sub

Private Sub modificacion_Click()
Form6.Show
End Sub

Private Sub reportegen_Click()
Form9.Show
End Sub

Private Sub salir_Click()
MsgBox "GRACIAS POR UTILIZAR ESTE SISTEMA. HASTA LA PROXIMA."
End
End Sub

Private Sub venta_Click()
Form8.Show
End Sub
