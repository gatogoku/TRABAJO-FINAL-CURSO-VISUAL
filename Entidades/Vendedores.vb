Imports Ventas
Public Class Vendedores : Inherits Empleados

    Public Property IdVendedor As Integer
    Public Property Nombre As String
    Public Property Foto As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal id As Integer, nombre As String, foto As String)
        Me.IdVendedor = id
        Me.Nombre = nombre
        Me.Foto = foto


    End Sub


End Class
