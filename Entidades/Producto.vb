Imports Entidades

Public Class Producto
    Implements IEquatable(Of Producto)
    Public Property Descripcion As String
    Public Property Familia As String
    Public Property IdProducto As Integer

    Public Property Precio As Double
    Public Property Estado As String

    Public Property Talla As String
    'Public Property IdEntrega As Integer
    'Public Property CodProducto As String
    'Public Property CodigoTalla As String
    'Public Property NumeroProducto As Integer

    Public Sub New(ByVal desc As String, familia As String, id As Integer, precio As Double, Estado As String, talla As String)
        Me.Descripcion = desc
        Me.Familia = familia
        IdProducto = id
        Me.Precio = precio
        Me.Talla = talla
        Me.Estado = Estado

    End Sub

    Public Sub New()

    End Sub

    Public Function Equals(other As Producto) As Boolean Implements IEquatable(Of Producto).Equals
        Return other.IdProducto = Me.IdProducto
    End Function
End Class
