Imports GestionDB, Entidades
Public Class frmVentas



    Dim CategoriaSeleccionada As String
    Dim compraRealizada As Boolean = False
    Dim productoAuxiliar As Producto
    Dim imagenes() As String = {"ABRIGO-CAZADORA.png", "ACCESORIOS-BISUTERIA.png", "BOLSOS-MOCHILAS.png", "CAMISA.png", "corta.png", "larga.png", "falda.png", "jersey.png", "pcorto.png", "plargo.png", "vestidop.png", "zapas.png"}
    'Dim NombresCategoriasas() As String = {"ABRIGO/CAZADORA", "ACCESORIOS/BISUTERIA", "BOLSOS/MOCHILLAS", "CAMISA", "CAMISA M. CORTA", "CAMISA M. LARGA", "FALDAS", "JERSEY", "PANTALONES CORTOS", "PANTALONES LARGOS", "VESTIDOS", "ZAPATOS"}
    Dim NombresCategoriasas() As String = {"AB", "AC", "BO", "CA", "CC", "CL", "FA", "JE", "PC", "PL", "VE", "ZA"}
    Private Sub Ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None

        For c As Integer = GroupBox1.Controls.Count - 1 To 0 Step -1
            Try
                GroupBox1.Controls(c).BackgroundImage = Image.FromFile("Recursos\" & imagenes(c))
                GroupBox1.Controls(c).BackgroundImageLayout = ImageLayout.Stretch
                GroupBox1.Controls(c).Tag = NombresCategoriasas(c)
                GroupBox1.Controls(c).Text = ""
            Catch ex As Exception
            End Try
        Next
        Button18.BackgroundImage = Image.FromFile("Recursos\Excel.png")
        GroupBox2.Left = Me.Width / 2 - GroupBox2.Width / 2
        ListBox1.Left = Me.Width / 2 - ListBox1.Width / 2
        DataGridView1.DefaultCellStyle.BackColor = Color.Black
        ListBox1.BackColor = Color.Black


    End Sub


    Public Sub Colores()

        DataGridView1.GridColor = Color.Black
        DataGridView1.BackgroundColor = Color.Black
        DataGridView1.ForeColor = Color.White

        For Each fila As DataGridViewRow In DataGridView1.Rows

            Select Case fila.Cells("Estado").Value
                Case "DISPONIBLE"
                    fila.DefaultCellStyle.BackColor = Color.ForestGreen
                Case "RESERVADO"
                    fila.DefaultCellStyle.BackColor = Color.Blue
                Case "VENDIDO"
                    fila.DefaultCellStyle.BackColor = Color.Red
                Case "DESCONOCIDO"
                    fila.DefaultCellStyle.BackColor = Color.DarkSlateGray
            End Select
        Next
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'MessageBox.Show(DataGridView1.CurrentCell.RowIndex, "cwll!")
        'SeleccionarProducto()
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        'MessageBox.Show(DataGridView1.CurrentCell.RowIndex, "clik")
        SeleccionarProducto()
    End Sub
    Private Sub SeleccionarProducto()
        If DataGridView1.CurrentCell.RowIndex = -1 Then MessageBox.Show("no elegida") : Exit Sub
        ' DataGridView1.SelectedRows.co
        Dim IndiceFila As Integer = DataGridView1.CurrentCell.RowIndex 'sender.RowIndex
        Dim Campos As New List(Of String)
        For c As Integer = 0 To 5
            Campos.Add(DataGridView1.Item(c, IndiceFila).Value)
        Next
        Dim productaux As New Producto(Campos.Item(0), Campos.Item(1), Campos.Item(2), Campos.Item(3), Campos.Item(4), Campos.Item(5))
        productoAuxiliar = productaux
    End Sub
    'Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) 'ESTE AUN NO SE SI ES UTIL O NO 

    '    Dim IndiceFila As Integer = DataGridView1.CurrentCell.RowIndex 'sender.RowIndex
    '    Dim Campos As New List(Of String)
    '    For c As Integer = 0 To 5
    '        Campos.Add(DataGridView1.Item(c, IndiceFila).Value)
    '    Next

    '    Dim productaux As New Producto(Campos.Item(0), Campos.Item(1), Campos.Item(2), Campos.Item(3), Campos.Item(4), Campos.Item(5))
    '    productoAuxiliar = productaux
    '    ''Dim camposopcion As New ToolStrip

    ''camposopcion.Items.Add("AÑADIR")
    ''camposopcion.Items.Add("RESERVAR")
    ''camposopcion.Items.Add("QUITAR")

    ''camposopcion.Width = 500
    ''camposopcion.Dock = DockStyle.Left

    'Dim opciones As New MenuStrip

    'opciones.Dock = DockStyle.None

    'opciones.Items.Add("AÑADIR")
    'opciones.Items.Add("RESERVAR")
    'opciones.Items.Add("QUITAR")

    'Dim punto As Point
    ''punto = sender.location
    'punto = New Point(GroupBox2.Width / 2 - opciones.Width / 2, 40)

    'opciones.Location = punto
    ''GroupBox2.Controls.Add(camposopcion)
    'opciones.BackColor = Color.Blue
    'opciones.ForeColor = Color.White
    'opciones.BringToFront()
    'opciones.Visible = True
    'GroupBox2.Controls.Add(opciones)

    'AddHandler opciones.Click, AddressOf opcionesEvents

    'End Sub

    'Private Sub opcionesEvents(sender As Object, e As EventArgs)
    '    'Dim opcion As New Button
    '    sender = TryCast(sender, ToolStrip)
    '    Select Case sender.Name
    '        Case "AÑADIR"
    '            Module1.miGestionBD.AñadirArticulosVenta(productoAuxiliar)
    '        Case "RESERVAR"
    '            Module1.miGestionBD.ReservarProductos(productoAuxiliar)
    '            DataGridView1.Update()
    '        Case "QUITAR"
    '            Module1.miGestionBD.QuitarArticulosVenta(productoAuxiliar)
    '    End Select

    'End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Me.Close()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button4.Click, Button3.Click, Button2.Click, Button12.Click, Button11.Click, Button10.Click, Button1.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()

        Dim Btntipo As New Button

        Btntipo = TryCast(sender, Button)
        CategoriaSeleccionada = Btntipo.Tag
        Module1.ListaVentaProductosDisponibles = miGestionBD.RetornarArticulosPorCategoria(Btntipo.Tag)

        DataGridView1.DataSource = Module1.ListaVentaProductosDisponibles
        Colores()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If productoAuxiliar Is Nothing Then
            MessageBox.Show("No hay producto")
            Exit Sub
        End If
        If Module1.ListaProductosVenta.Contains(productoAuxiliar) Then
            MessageBox.Show("El producto ya esta añadido")
            Exit Sub
        End If

        Module1.miGestionBD.AñadirArticulosVenta(productoAuxiliar)

        Module1.ListaProductosVenta = Module1.miGestionBD.DevolverListaVenta
        ListBox1.Items.Clear()
        'ListBox1.Items.AddRange(Module1.ListaProductosVenta.ToArray)

        For c As Integer = 0 To Module1.ListaProductosVenta.Count - 1
            'ListBox1.Items.Add(Module1.ListaProductosVenta(c).Descripcion & "  -  " & ListaProductosVenta(c).Familia & "  -  " & ListaProductosVenta(c).IdProducto & "  -  " & Module1.ListaProductosVenta(c).Precio & "  -  " & Module1.ListaProductosVenta(c).Estado & "  -  " & Module1.ListaProductosVenta(c).Talla)
            ListBox1.Items.Add(Module1.ListaProductosVenta(c))
            ListBox1.DisplayMember = "Descripcion" ' & "Familia"
        Next
        productoAuxiliar = Nothing
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If ListBox1.SelectedItem Is Nothing Then
            MessageBox.Show("No hay producto")
            Exit Sub
        End If

        productoAuxiliar = ListBox1.SelectedItem

        Module1.miGestionBD.QuitarArticulosVenta(productoAuxiliar)

        Module1.ListaProductosVenta = Module1.miGestionBD.DevolverListaVenta
        ListBox1.Items.Clear()

        'ListBox1.Items.AddRange(Module1.ListaProductosVenta.ToArray)

        For c As Integer = 0 To Module1.ListaProductosVenta.Count - 1
            'ListBox1.Items.Add(Module1.ListaProductosVenta(c).Descripcion & "  -  " & ListaProductosVenta(c).Familia & "  -  " & ListaProductosVenta(c).IdProducto & "  -  " & Module1.ListaProductosVenta(c).Precio & "  -  " & Module1.ListaProductosVenta(c).Estado & "  -  " & Module1.ListaProductosVenta(c).Talla)
            ListBox1.Items.Add(Module1.ListaProductosVenta(c))
            ListBox1.DisplayMember = "Descripcion" ' & "Familia"
        Next
        'ListBox1.Items.Add(productoAuxiliar) 'Module1.ListaProductosVenta(c).Descripcion & "  -  " & Module1.ListaProductosVenta(c).Precio & "  -  " & Module1.ListaProductosVenta(c).Estado & "  -  " & Module1.ListaProductosVenta(c).Talla)
        'Next
        'productoAuxiliar = Nothing
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Try
            Module1.miGestionBD.ReservarProductos(productoAuxiliar)
            MsgBox("PRODUCTO RESERVADO")
            Module1.ListaVentaProductosDisponibles.Clear()
            Module1.ListaVentaProductosDisponibles = miGestionBD.RetornarArticulosPorCategoria(CategoriaSeleccionada)
            DataGridView1.DataSource = ""
            DataGridView1.DataSource = Module1.ListaVentaProductosDisponibles
            Colores()
        Catch ex As Exception
            MsgBox("NO HAY PRODUCTO SELECCIONADO PARA RESERVAR")
        End Try

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If compraRealizada Then
            MsgBox("REINICIE VENTAS ANTES DE UNA NUEVA VENTA")
        Else
            If Module1.ListaProductosVenta.Count = 0 Then
                MsgBox("NO HAY PRODUCTOS AGREGADOS A LA CESTA")
            Else
                Module1.miGestionBD.RealizarVentaListaProductos(Module1.ListaProductosVenta, Module1.EmpleadoActual)
                compraRealizada = True
                MsgBox("COMPRA REALIZADA CORRECTAMENTE")
                'Module1.ListaProductosVenta.Clear()
                ListBox1.Items.Clear()
            End If
        End If


    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If compraRealizada Then
            Try
                Dim oExcel As Object
                Dim oBook As Object
                Dim oSheet As Object
                Dim contador As Integer = 2
                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.workbooks.add


                oSheet = oBook.Worksheets(1)
                oSheet.range("A1").value = "Nombre del producto"
                oSheet.range("B1").value = "Id Producto"
                oSheet.range("C1").value = "Precio"
                oSheet.range("D1").value = "Familia"
                oSheet.range("E1").value = "Talla"
                oSheet.range("F1").value = "Nombre de vendedor"
                oSheet.range("G1").value = "Codigo de ticket"
                oSheet.range("A1:E1").Font.Bold = True

                Dim total As Double = 0
                For Each prod As Producto In Module1.ListaProductosVenta

                    oSheet.range("A" & contador).Value = prod.Descripcion
                    oSheet.range("B" & contador).Value = prod.IdProducto
                    oSheet.range("C" & contador).Value = prod.Precio
                    oSheet.range("D" & contador).Value = prod.Familia
                    oSheet.range("E" & contador).Value = prod.Talla

                    contador = contador + 1
                    total = total + prod.Precio
                Next

                oSheet.range("C" & contador + 1).Value = total
                oSheet.range("F" & contador + 1).value = Module1.EmpleadoActual.Nombre
                oSheet.range("G" & contador + 1).value = ""
                ' oBook.open

                ' oBook.SaveAs("Book1.xlsx")
                oSheet = Nothing
                oBook = Nothing
                oExcel.Quit()
                oExcel = Nothing

                GC.Collect()
            Catch ex As Exception
                MsgBox("NO SE HA PODIDO IMPRIMIR EL TICKET")
            End Try


        Else
            MsgBox("NO SE PUEDE IMPRIMIR EL TICKET DE UNA VENTA QUE NO HA SIDO REALIZADA")
        End If
    End Sub



End Class