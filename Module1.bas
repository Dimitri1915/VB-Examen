Attribute VB_Name = "Module1"
Option Explicit

Private Const Productos As String = "productos.txt"
Private Const Entradas As String = "entradas.txt"
Private Const Salidas As String = "salidas.txt"

Public Type Producto
    CodigoBarras As String
    Nombre As String
    Existencias As Integer
End Type

Public Productos() As Producto

Public Sub LeerProductos()
    Dim linea As String
    Dim campos() As String
    Dim i As Integer
    Dim f As Integer
    
    On Error GoTo FileError
    f = FreeFile
    Open ProductosFile For Input As f
        i = 0
        Do While Not EOF(f)
            Line Input #f, linea
            campos = Split(linea, "|")
            ReDim Preserve Productos(i)
            With Productos(i)
                .CodigoBarras = campos(0)
                .Nombre = campos(1)
                .Existencias = Val(campos(2))
            End With
            i = i + 1
        Loop
    Close f
    Exit Sub
FileError:
    Close f
    MsgBox "Error al leer el archivo de productos.", vbCritical
End Sub

Public Sub GuardarProductos()
    Dim i As Integer
    Dim f As Integer
    
    On Error GoTo FileError
    f = FreeFile
    Open ProductosFile For Output As f
        For i = 0 To UBound(Productos)
            Print #f, Productos(i).CodigoBarras & "|" & Productos(i).Nombre & "|" & Productos(i).Existencias
        Next i
    Close f
    Exit Sub
FileError:
    Close f
    MsgBox "Error al guardar el archivo de productos.", vbCritical
End Sub

