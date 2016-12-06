Imports Empleados

Public Class Administradores : Inherits Empleados
    Implements IEquatable(Of Administradores)


    Public Property Nombre As String
    'Private Apellido As String
    Public Property Contraseña As String


    Public Sub New()

    End Sub


    'Public Sub New(ByVal nombre As String, apellido As String, contraseña As String)
    '    Me.Nombre = nombre
    '    Me.Contraseña = contraseña
    'End Sub

    Public Sub New(ByVal nombre As String, contraseña As String)
        Me.Nombre = nombre
        Me.Contraseña = contraseña
    End Sub

    Public Function Equals(other As Administradores) As Boolean Implements IEquatable(Of Administradores).Equals
        Return other.Nombre.Equals(Me.Nombre)
    End Function
End Class
