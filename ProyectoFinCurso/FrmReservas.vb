Imports Entidades

Public Class FrmReservas

    Dim productoAuxiliar As Producto


    Public Sub Iniciar()

        Module1.ListaProductosReservados = Module1.miGestionBD.RetornarArticulosReservados

        DataGridView1.DataSource = Module1.ListaProductosReservados


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


    Private Sub FrmReservas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None
        Iniciar()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Me.Close()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Try
            Module1.miGestionBD.QuitarReservaProducto(productoAuxiliar)
            MsgBox("PRODUCTO DISPONIBLE")
            Iniciar()
        Catch ex As Exception
            MsgBox("NO HA SELECCIONADO NINGUN PRODUCTO PARA QUITAR RESERVA")
        End Try

        'Module1.ListaProductosReservados.Clear()
        'Module1.ListaProductosReservados = miGestionBD.RetornarArticulosReservados
        'DataGridView1.DataSource = ""
        'DataGridView1.DataSource = Module1.ListaProductosReservados
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
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
End Class