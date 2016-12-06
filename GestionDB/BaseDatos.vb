Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Imports Entidades


Public Class BaseDatos
    Dim ListaProductosVenta As New List(Of Producto)

    'PARTE DE LA BASE DE DATOS

    Private Const cadConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=CUASHOP_MODA.accdb;
Persist Security Info=False;"

    Dim sqlPredicado As String = ""
    Dim ConexionSQLCuashop As New OleDbConnection(cadConexion)

    Public Sub New()

    End Sub


    Public Function DevolverProductosTicket(ByVal ListaDevolucionDefinitiva As List(Of Producto), ByVal idVENTA As Integer) ' TENGO QUE PENSAR QUE SE PUEDE HACER


        sqlPredicado = "INSERT  INTO Productos  SELECT * FROM Vendidos WHERE Vendidos.IdProducto = @ID " 'REINSERTAR EN PRODUCTOS EN LA TABLA PRODUCTOS
        Dim CMDInsertarEnVentas As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        ' CMDInsertarEnVentas.Parameters.AddWithValue("@IDVENTA", 0)
        CMDInsertarEnVentas.Parameters.AddWithValue("@ID", Nothing)
        Try
            ConexionSQLCuashop.Open()
            For Each prod As Producto In ListaDevolucionDefinitiva
                CMDInsertarEnVentas.Parameters(0).Value = prod.IdProducto

                CMDInsertarEnVentas.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try


        'sqlPredicado = "SELECT * FROM VentasProducto WHERE @IdVenta = @IDVENTA"
        sqlPredicado = "DELETE FROM VentasProducto WHERE ventasProducto.IdVenta = @IDVENTA AND  ventasProducto.IdProducto = @IDPRODUCTO" ' BORRAR REGISTROS RELACIONADOS DE VENTASPRODUCTO
        Dim ObtenerMax As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        ObtenerMax.Parameters.AddWithValue("@IDVENTA", idVENTA)
        ObtenerMax.Parameters.AddWithValue("@IPRODUCTO", Nothing)
        Try
            ConexionSQLCuashop.Open()
            For Each prod As Producto In ListaDevolucionDefinitiva
                ObtenerMax.Parameters(1).Value = prod.IdProducto
                ObtenerMax.ExecuteScalar()
            Next

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try


        sqlPredicado = "DELETE FROM  Vendidos WHERE Vendidos.IdProducto = @ID " 'BORRAR LOS PRODUCTOS DE LA TABLA VENDIDOS
        Dim CMDBorrar As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        ' CMDInsertarEnVentas.Parameters.AddWithValue("@IDVENTA", 0)
        CMDBorrar.Parameters.AddWithValue("@ID", Nothing)
        Try
            ConexionSQLCuashop.Open()
            For Each prod As Producto In ListaDevolucionDefinitiva
                CMDBorrar.Parameters(0).Value = prod.IdProducto

                CMDBorrar.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
    End Function

    Public Function RetornarArticulosReservados() As List(Of Producto)
        sqlPredicado = "SELECT * FROM Reservados"
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        'CMDRetornarCateegoria.Parameters.AddWithValue("@ESTADO", "RESERVADO")
        Dim ListaProductosReservados As New List(Of Producto)
        Try
            ConexionSQLCuashop.Open()
            Dim Productos As OleDbDataReader = CMDRetornarCateegoria.ExecuteReader
            While Productos.Read()

                Dim productoX As New Producto
                productoX.Descripcion = Productos("Descripcion")
                productoX.IdProducto = Productos("IdProducto")
                productoX.Precio = Productos("Precio")
                productoX.Familia = ""
                'productoX.NumeroProducto = Productos("NumeroProd")
                productoX.Talla = Productos("Talla")
                Try
                    productoX.Estado = Productos("EstadoProducto")
                Catch ex As Exception
                    productoX.Estado = "DESCONOCIDO"
                End Try
                ListaProductosReservados.Add(productoX)
            End While

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
        Return ListaProductosReservados
        'FUNCION ANTERIOR

        'sqlPredicado = "SELECT * FROM Productos WHERE Productos.EstadoProducto = @ESTADO"
        'Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLHOTEL)
        'CMDRetornarCateegoria.Parameters.AddWithValue("@ESTADO", "RESERVADO")
        'Dim ListaProductosReservados As New List(Of Producto)
        'Try
        '    ConexionSQLHOTEL.Open()
        '    Dim Productos As OleDbDataReader = CMDRetornarCateegoria.ExecuteReader
        '    While Productos.Read()

        '        Dim productoX As New Producto
        '        productoX.Descripcion = Productos("Descripcion")
        '        productoX.IdProducto = Productos("IdProducto")
        '        productoX.Precio = Productos("Precio")
        '        productoX.Familia = ""
        '        'productoX.NumeroProducto = Productos("NumeroProd")
        '        productoX.Talla = Productos("Talla")
        '        Try
        '            productoX.Estado = Productos("EstadoProducto")
        '        Catch ex As Exception
        '            productoX.Estado = "DESCONOCIDO"
        '        End Try
        '        ListaProductosReservados.Add(productoX)
        '    End While

        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    ConexionSQLHOTEL.Close()
        'End Try
        'Return ListaProductosReservados

    End Function

    Public Sub QuitarReservaProducto(ByVal product As Producto)

        sqlPredicado = "INSERT  INTO Productos SELECT * FROM  Reservados  WHERE  Reservados.IdProducto = @ID "
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDRetornarCateegoria.Parameters.AddWithValue("@ID", product.IdProducto)
        Try
            ConexionSQLCuashop.Open()
            CMDRetornarCateegoria.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try

        sqlPredicado = "DELETE FROM Reservados  WHERE  Reservados.IdProducto = @ID "
        Dim CMDQuitar As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDQuitar.Parameters.AddWithValue("@ID", product.IdProducto)
        Try
            ConexionSQLCuashop.Open()
            CMDQuitar.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
    End Sub


    Public Function RetornarListaProductosTicket(ByVal NumeroTicket As Integer) As List(Of Producto)

        sqlPredicado = "SELECT * FROM Vendidos INNER JOIN VentasProducto ON (Vendidos.IdProducto = VentasProducto.IdProducto)  WHERE VentasProducto.IdVenta = @ID"
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDRetornarCateegoria.Parameters.AddWithValue("@ID", NumeroTicket)
        Dim ListaProductosTicket As New List(Of Producto)
        Try
            ConexionSQLCuashop.Open()
            Dim Productos As OleDbDataReader = CMDRetornarCateegoria.ExecuteReader
            While Productos.Read()

                Dim productoX As New Producto
                productoX.Descripcion = Productos("Descripcion")
                'Try
                productoX.IdProducto = Productos("Vendidos.IdProducto")
                'Catch ex As Exception
                'productoX.IdProducto = 0
                'end try

                productoX.Precio = Productos("Vendidos.Precio")
                productoX.Familia = ""
                'productoX.NumeroProducto = Productos("NumeroProd")
                productoX.Talla = Productos("Talla")
                Try
                    productoX.Estado = Productos("EstadoProducto")
                Catch ex As Exception
                    productoX.Estado = "DESCONOCIDO"
                End Try
                ListaProductosTicket.Add(productoX)
            End While

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
        Return ListaProductosTicket
    End Function

    'operaciones con los productos
    Public Function RetornarArticulosPorCategoria(ByVal Categoria As String) As List(Of Producto) 'FUNCIONA BIEN

        Dim exist As Boolean = False

        'sqlPredicado = "SELECT * FROM Productos  WHERE  Productos.Idfamilia = @TIPO "
        sqlPredicado = "SELECT * FROM Productos INNER JOIN Familias ON (Productos.Idfamilia = Familias.IdFamilia)  WHERE  Familias.Idfamilia = @TIPO "
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDRetornarCateegoria.Parameters.AddWithValue("@TIPO", Categoria)
        Dim ListaProductosCategoria As New List(Of Producto)
        Try
            ConexionSQLCuashop.Open()
            Dim Productos As OleDbDataReader = CMDRetornarCateegoria.ExecuteReader
            While Productos.Read()

                Dim productoX As New Producto
                productoX.Descripcion = Productos("Descripcion")
                productoX.IdProducto = Productos("IdProducto")
                productoX.Precio = Productos("Precio")
                productoX.Familia = Categoria
                'productoX.NumeroProducto = Productos("NumeroProd")
                productoX.Talla = Productos("Talla")

                Try
                    productoX.Estado = Productos("EstadoProducto")
                Catch ex As Exception
                    productoX.Estado = "DESCONOCIDO"
                End Try

                ListaProductosCategoria.Add(productoX)

            End While

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
        Return ListaProductosCategoria
    End Function

    Public Function AñadirArticulosVenta(ByVal prod As Producto) 'BIEN CREO

        ListaProductosVenta.Add(prod)

    End Function

    Public Function QuitarArticulosVenta(ByVal prod As Producto) 'BIEN CREO

        ListaProductosVenta.Remove(prod)

    End Function



    Public Function DevolverListaVenta() As List(Of Producto) 'BIEN CREO 

        Return ListaProductosVenta

    End Function

    Public Function ReservarProductos(ByVal product As Producto)

        'FUNCION NUEVA
        'sqlPredicado = "SELECT * FROM  Productos  WHERE  Productos.IdProducto = @ID "
        sqlPredicado = "INSERT  INTO Reservados SELECT * FROM  Productos  WHERE  Productos.IdProducto = @ID "
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDRetornarCateegoria.Parameters.AddWithValue("@ID", product.IdProducto)
        Try
            ConexionSQLCuashop.Open()
            CMDRetornarCateegoria.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
        sqlPredicado = "DELETE FROM Productos  WHERE  Productos.IdProducto = @ID "
        Dim CMDQuitar As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDQuitar.Parameters.AddWithValue("@ID", product.IdProducto)
        Try
            ConexionSQLCuashop.Open()
            CMDQuitar.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try
    End Function

    Public Function RealizarVentaListaProductos(ByVal ListaProductosVenta As List(Of Producto), ByVal vendedor As Empleados)

        sqlPredicado = "INSERT  INTO Ventas (IdTrabajador, FechaVenta) VALUES(@IDTRABAJADOR, DATE()) " 'INSERTA DATOS EN VENTAS
        Dim CMDInsertarEnVentas As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        ' CMDInsertarEnVentas.Parameters.AddWithValue("@IDVENTA", 0)
        CMDInsertarEnVentas.Parameters.AddWithValue("@IDTRABAJADOR", vendedor.IdVendedor)
        Try
            ConexionSQLCuashop.Open()
            CMDInsertarEnVentas.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try

        Dim IdMaximo As Integer

        sqlPredicado = "SELECT MAX (Ventas.IdVenta) FROM Ventas" 'DEVUELVE EL NUMERO MAXIMO DE VENTAS.IDVENTAS
        Dim ObtenerMax As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        Try
            ConexionSQLCuashop.Open()

            IdMaximo = ObtenerMax.ExecuteScalar()

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try


        sqlPredicado = "INSERT  INTO Vendidos SELECT * FROM  Productos  WHERE  Productos.IdProducto = @ID " ' INSERTA PRODUCTOS EN VENDIDOS DESDE PRODUCTOS REGISTROS ENTEROS
        ' Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLHOTEL)
        Dim InsertarEnVendidos As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        InsertarEnVendidos.Parameters.AddWithValue("@IDENTIDAD", Nothing)
        Try
            ConexionSQLCuashop.Open()
            For Each prod As Producto In ListaProductosVenta
                InsertarEnVendidos.Parameters(0).Value = prod.IdProducto
                InsertarEnVendidos.ExecuteNonQuery()
            Next

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try


        'SEGUNDA PARTE DEL PROCESO
        sqlPredicado = "INSERT  INTO VentasProducto VALUES(@IDVENTA,@IDPRODUCTO,@PRECIO) " 'INSERTA DATOS EN VENTAS.PRODUCTO X CADA PRODUCTO
        Dim InsertarVentasProductos As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        InsertarVentasProductos.Parameters.AddWithValue("@IDVENTA", Nothing)
        InsertarVentasProductos.Parameters.AddWithValue("@IDPRODUCTO", Nothing)
        InsertarVentasProductos.Parameters.AddWithValue("@PRECIO", Nothing)

        Try
            ConexionSQLCuashop.Open()
            For Each prod As Producto In ListaProductosVenta
                InsertarVentasProductos.Parameters(0).Value = IdMaximo
                InsertarVentasProductos.Parameters(1).Value = prod.IdProducto
                InsertarVentasProductos.Parameters(2).Value = prod.Precio

                InsertarVentasProductos.ExecuteNonQuery()
            Next

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try



        sqlPredicado = "DELETE FROM  Productos WHERE Productos.IdProducto = @ID " 'BORRAR LOS PRODUCTOS DE LA TABLA PRODUCTOS
        Dim CMDBorrar As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        ' CMDInsertarEnVentas.Parameters.AddWithValue("@IDVENTA", 0)
        CMDBorrar.Parameters.AddWithValue("@ID", Nothing)
        Try
            ConexionSQLCuashop.Open()
            For Each prod As Producto In ListaProductosVenta
                CMDBorrar.Parameters(0).Value = prod.IdProducto

                CMDBorrar.ExecuteNonQuery()
            Next
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try

    End Function

    Public Function DevolverArticulos(ByVal listaDevolucion As List(Of Producto))
        sqlPredicado = "INSERT  INTO Productos SELECT * FROM  Vendidos  WHERE  Vendidos.IdProducto = @ID "
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)

        Try
            ConexionSQLCuashop.Open()
            For Each product As Producto In listaDevolucion
                CMDRetornarCateegoria.Parameters.AddWithValue("@TIPO", product.IdProducto)
                CMDRetornarCateegoria.ExecuteNonQuery()
            Next

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try


    End Function


    'ticket 
    Public Function CrearTicket()

    End Function

    'ventas
    Public Function TotalVentasDia() As Double 'OleDbDataReader  ' ANTIGUA FUNCTION PARA CONSULTAR EN SQL
        '    Dim fecha As Date = (Date.Today)

        '    Dim dia As String = fecha.Day & "/" & fecha.Month & "/" & fecha.Year

        '    'sqlPredicado = "SELECT * FROM  Vendidos INNER JOIN Ventas ON (Vendidos.Idventa = Ventas.Idventa)  WHERE  Ventas.FechaVenta = @FECHA"
        '    sqlPredicado = "SELECT * FROM  Vendidos WHERE Vendidos.Idventa =  (SELECT IdVenta FROM Ventas WHERE  Ventas.FechaVenta = Date())"
        '    Dim CMDRetornarCajaDia As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        '    'CMDRetornarCajaDia.Parameters.AddWithValue("@FECHA", "1/06/2016")
        '    Try
        '        ConexionSQLCuashop.Open()
        '        Dim dr As OleDbDataReader = CMDRetornarCajaDia.ExecuteReader
        '        Return dr
        '        'dr.Close()
        '    Catch ex As Exception
        '        Throw ex
        '    Finally
        '        ' ConexionSQLCuashop.Close()
        '    End Try

        'End Function


        'Public Function TotalCajaDia() As Double
        '    ''Try
        '    ''    Dim dr As IDataReader = TotalVentasDia()
        '    ''    Dim total As Double = 0

        '    ''    While dr.Read
        '    ''        total = total + dr("Vendidos.Precio")
        '    ''    End While
        '    ''    dr.Close()
        '    ''    Return total

        '    ''Catch ex As Exception
        '    ''    Throw ex
        '    ''Finally
        '    ''    ConexionSQLCuashop.Close()
        '    ''End Try


        '    'NUEVA FUNCION


        '    Dim fecha As Date = (Date.Today)

        '    Dim dia As String = fecha.Day & "/" & fecha.Month & "/" & fecha.Year

        '    sqlPredicado = "SELECT * FROM  Vendidos INNER JOIN Ventas ON (Vendidos.Idventa = Ventas.Idventa)  WHERE  Ventas.FechaVenta >= @FECHA"
        '    'sqlPredicado = "SELECT * FROM  Vendidos WHERE Vendidos.Idventa =  (SELECT IdVenta FROM Ventas WHERE  Ventas.FechaVenta = Date())"
        '    Dim CMDRetornarCajaDia As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        '    CMDRetornarCajaDia.Parameters.AddWithValue("@FECHA", "1/06/2016") 'Today.Date)
        '    Try
        '        ConexionSQLCuashop.Open()
        '        Dim dr As OleDbDataReader = CMDRetornarCajaDia.ExecuteReader
        '        Dim total As Double = 0
        '        While dr.Read
        '            total = total + dr("Vendidos.Precio")
        '        End While
        '        dr.Close()
        '        Return total

        '        dr.Close()
        '    Catch ex As Exception
        '        Throw ex
        '    Finally
        '        ConexionSQLCuashop.Close()
        '    End Try

        'End Function


        sqlPredicado = "SELECT * FROM  Ventas  INNER JOIN  VentasProducto ON (VentasProducto.Idventa = Ventas.Idventa)  WHERE  Ventas.FechaVenta = @FECHA " '& Date.Today.ToShortDateString
        Dim CMDTotalDia As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDTotalDia.Parameters.AddWithValue("@FECHA", Date.Today)
        Try
            ConexionSQLCuashop.Open()

            Dim dr As OleDbDataReader = CMDTotalDia.ExecuteReader
            Dim total As Double = 0
            While dr.Read
                total = total + dr("Precio")
            End While

            Return total
        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try

    End Function

    Public Function DocumentoTotalCajaDia() As List(Of Producto)
        Dim productos As IDataReader '= TotalVentasDia()
        Dim ListaProductosVendidosDia As New List(Of Producto)

        Try
            While productos.Read
                Dim productoX As New Producto
                productoX.Descripcion = productos("Descripcion")
                productoX.IdProducto = productos("IdProducto")
                productoX.Precio = productos("Precio")
                productoX.Familia = ""
                'productoX.NumeroProducto = Productos("NumeroProd")
                productoX.Talla = productos("Talla")
                Try
                    productoX.Estado = productos("EstadoProducto")
                Catch ex As Exception
                    productoX.Estado = "DESCONOCIDO"
                End Try
                ListaProductosVendidosDia.Add(productoX)
            End While
            Return ListaProductosVendidosDia
        Catch ex As Exception
            Throw ex
        Finally
            conexionSQLCuashop.Close()
        End Try


    End Function





    Public Function TotalVentasTotal() As Double

        sqlPredicado = "SELECT * FROM  Vendidos" 'INNER JOIN Ventas ON (Vendidos.Idventa = Ventas.Idventa)  WHERE  Ventas.FechaVenta = @FECHA "
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)

        Try
            ConexionSQLCuashop.Open()
            'CMDRetornarCateegoria.Parameters.AddWithValue("@FECHA", dia)
            Dim dr As OleDbDataReader = CMDRetornarCateegoria.ExecuteReader
            Dim total As Double = 0
            While dr.Read
                total = total + dr("Precio")
            End While
            Return total
            'ConexionSQLCuashop.Close()
        Catch ex As Exception
            Throw ex

        Finally
            ConexionSQLCuashop.Close()
        End Try

    End Function


    'empleados

    Public Function EncriptarContraseña(ByVal contraseña As String)
        Encriptacion.MD5EncryptPass(contraseña)
    End Function




    Public Function ComprobarUsuario(ByVal usuario As String, ByVal contraseña As String) As Empleados

        If contraseña = "" Then
            sqlPredicado = "SELECT  Trabajadores.* FROM Trabajadores WHERE Trabajadores.Nombre = @USUARIO "
        Else
            sqlPredicado = "SELECT  Administradores.* FROM Administradores WHERE Administradores.Nombre = @USUARIO "
        End If
        Dim CMDRetornarCateegoria As New OleDbCommand(sqlPredicado, ConexionSQLCuashop)
        CMDRetornarCateegoria.Parameters.AddWithValue("@USUARIO", usuario)
        Try
            ConexionSQLCuashop.Open()
            Dim dr As OleDbDataReader = CMDRetornarCateegoria.ExecuteReader
            If dr.Read Then
                Dim empleado As Empleados
                Try
                    empleado = New Administradores
                    empleado.Contraseña = dr("Contraseña")
                    empleado.Nombre = dr("Nombre")

                    If empleado.Nombre = usuario And empleado.Contraseña = contraseña Then 'EncriptarContraseña(contraseña) Then ESTA ES LA PARTE PARA ENCRIPTAR CONTRASEÑA
                        Return empleado
                    End If

                Catch ex As Exception
                    empleado = New Vendedores
                    empleado.IdVendedor = dr("IdTrabajador")
                    empleado.Nombre = dr("Nombre")
                    Try
                        empleado.Foto = dr("Foto")
                    Catch ex2 As Exception
                        empleado.Foto = ""
                    End Try

                    If empleado.Nombre = usuario Then
                        Return empleado
                    End If
                End Try
            End If

        Catch ex As Exception
            Throw ex
        Finally
            ConexionSQLCuashop.Close()
        End Try

    End Function

    'SEGURIDAD ADMINISTRADOR


    'Public Function RealizarCopiaDeSeguridad() 'REALIZAR COPIA DE SEGURIDAD AÑADIENDO AL NOMBRE LA FECHA 

    '    Dim fechaActual As Date = Date.Today

    '    ' Copy the file to a new location without overwriting existing file.
    '    My.Computer.FileSystem.CopyFile(
    '"C:\UserFiles\TestFiles\testFile.txt",
    '"C:\UserFiles\TestFiles2\testFile.txt")


    '    ' Copy the file to a new folder and rename it.
    '    My.Computer.FileSystem.CopyFile(
    '        "C:\UserFiles\TestFiles\testFile.txt",
    '        "C:\UserFiles\TestFiles2\NewFile.txt",
    '        Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
    '        Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)


    'End Function


    'Public Function RestaurarCopiaDeSeguridad() ' SE MOSTRARA UN MENU CON LA COPIAS DE SEGURIDAD REALIZADAS Y SE PERMITIRA RESTAURAR LA SELECCIONADA




    'End Function

End Class
