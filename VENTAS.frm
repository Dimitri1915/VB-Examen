VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00404000&
   Caption         =   "frmVentas"
   ClientHeight    =   8370
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13500
   LinkTopic       =   "Form3"
   ScaleHeight     =   8370
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNombreProducto 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   7560
      TabIndex        =   11
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdCalcularTotal 
      Caption         =   "CALCULAR TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton btnObtenerDatos 
      BackColor       =   &H00000000&
      Caption         =   "OBTENER DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      MaskColor       =   &H00FF8080&
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar la venta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvVentas 
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo de barras"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "precio"
         Text            =   "precio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Nombre del producto"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdAgregarModificar 
      Caption         =   "agregar/modificar ventas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txtPrecio 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   8640
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtCantidad 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   8520
      TabIndex        =   1
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtCodigoBarras 
      BackColor       =   &H0080FF80&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   9840
      X2              =   9840
      Y1              =   4320
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   9840
      X2              =   9840
      Y1              =   2880
      Y2              =   3240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "NOMBRE DEL PRODUCTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   12
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "CODIGO DE BARRAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Menu regmenu 
      Caption         =   "REGRESAR AL MENU"
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnObtenerDatos_Click()
    ' código para acceder a los datos de la fila seleccionada
    Dim item As ListItem
    
    If lvVentas.SelectedItem Is Nothing Then
        MsgBox "Por favor seleccione una fila en el ListView", vbExclamation
        Exit Sub
    End If
    
    ' Acceder a los datos asociados con la fila seleccionada
    Set item = lvVentas.SelectedItem
    Dim datos() As String
    datos = Split(item.Tag, "|")
    
    Dim codigoBarras As String
    Dim nombreProducto As String
    Dim cantidad As Integer
    Dim precio As Double
    codigoBarras = datos(0)
    nombreProducto = datos(1)
    cantidad = Val(datos(2))
    precio = Val(datos(3))
    
    ' Calcular el total
    Dim total As Double
    total = cantidad * precio
    
    
    MsgBox "Código de Barras: " & codigoBarras & vbCrLf & _
           "Nombre del Producto: " & nombreProducto & vbCrLf & _
           "Cantidad: " & cantidad & vbCrLf & _
           "Precio: " & precio & vbCrLf & _
           "Total: " & total
End Sub

Private Sub cmdAgregarModificar_Click()
    Dim codigoBarras As String
    Dim nombreProducto As String
    Dim cantidad As Integer
    Dim precio As Double
    Dim item As ListItem
    Dim found As Boolean
    
    ' Obtener los valores ingresados
    codigoBarras = txtCodigoBarras.Text
    nombreProducto = txtNombreProducto.Text
    cantidad = Val(txtCantidad.Text)
    precio = Val(txtPrecio.Text)
    
    ' Validar que los campos no estén vacíos
    If codigoBarras = "" Or nombreProducto = "" Or cantidad <= 0 Or precio <= 0 Then
        MsgBox "Por favor ingrese todos los datos correctamente", vbExclamation
        Exit Sub
    End If
    
    ' Buscar el producto en la lista
    found = False
    For Each item In lvVentas.ListItems
        If item.Text = codigoBarras Then
            ' Si se encuentra, modificar los valores existentes
            item.SubItems(1) = nombreProducto
            item.SubItems(2) = cantidad
            item.SubItems(3) = precio
            item.Tag = codigoBarras & "|" & nombreProducto & "|" & cantidad & "|" & precio
            found = True
            Exit For
        End If
    Next item
    
    ' Si no se encuentra, agregar un nuevo producto
    If Not found Then
        Set item = lvVentas.ListItems.Add()
        item.Text = codigoBarras
        item.SubItems(1) = nombreProducto
        item.SubItems(2) = cantidad
        item.SubItems(3) = precio
        item.Tag = codigoBarras & "|" & nombreProducto & "|" & cantidad & "|" & precio
    End If
    
    ' Limpiar los campos
    txtCodigoBarras.Text = ""
    txtNombreProducto.Text = ""
    txtCantidad.Text = ""
    txtPrecio.Text = ""
End Sub

Private Sub cmdCalcularTotal_Click()
    ' Obtener la cantidad y el precio ingresados por el usuario
    Dim cantidad As Integer
    Dim precio As Double
    
    cantidad = Val(txtCantidad.Text) ' Suponiendo que txtCantidad contiene la cantidad ingresada
    precio = Val(txtPrecio.Text) ' Suponiendo que txtPrecio contiene el precio ingresado
    
    ' Calcular el total multiplicando la cantidad por el precio
    Dim total As Double
    total = cantidad * precio
    
    ' Mostrar el resultado al usuario
    MsgBox "El total es: " & total, vbInformation

End Sub

Private Sub Form_Load()
    ' Configurar columnas del ListView al cargar el formulario
    With lvVentas
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Add , , "Código de Barras", 150
        .ColumnHeaders.Add , , "Nombre del Producto", 150
        .ColumnHeaders.Add , , "Cantidad", 100
        .ColumnHeaders.Add , , "Precio", 100
    End With
End Sub


Private Sub regmenu_Click()
Unload Me
    
    ' Mostrar el formulario del menú principal
    Form2.Show
End Sub

Private Sub salir_Click()
' Confirmar si el usuario realmente quiere salir
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea salir de la aplicación?", vbYesNo + vbQuestion, "Confirmar salida")
    
    If respuesta = vbYes Then
        ' Cerrar la aplicación
        End
    End If
End Sub
