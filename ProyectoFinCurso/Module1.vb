Imports GestionDB, Entidades

Public Module Module1


    Public EmpleadoActual As Empleados
    Public miGestionBD As New GestionDB.BaseDatos
    Public ListaVentaProductosDisponibles As New List(Of Producto)
    Public ListaProductosVenta As New List(Of Producto)
    Public ListaProductosReservados As New List(Of Producto)
    Public ListaProductosParaDevolucion As New List(Of Producto)
    Public listaDevolucionFinal As New List(Of Producto)

End Module
