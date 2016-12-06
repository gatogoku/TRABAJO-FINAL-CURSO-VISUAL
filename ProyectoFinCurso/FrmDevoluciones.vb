Public Class FrmDevoluciones

    Private Sub FrmDevoluciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None

        ListBox2.DisplayMember = "Descripcion"
        ListBox1.DisplayMember = "Descripcion"
        Button16.BackgroundImage = Image.FromFile("Recursos\OK.png")

    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        If TextBox1.Text = "" Then
            MsgBox("DEBE INTRODUCIR UNA REFERENCIA PARA BUSCAR PRODUCTOS A DEVOLVER")
        Else

            If Module1.ListaProductosParaDevolucion.Count = 0 Then
                Try
                    Module1.ListaProductosParaDevolucion = Module1.miGestionBD.RetornarListaProductosTicket(TextBox1.Text)
                    ListBox1.Items.AddRange(Module1.ListaProductosParaDevolucion.ToArray)
                Catch ex As Exception
                    MsgBox("EL VALOR DEBE SER NUMERICO")
                    TextBox1.Clear()
                End Try

            Else
                MsgBox("YA TIENE UNA BUSQUEDA EN CURSO")
            End If
        End If
    End Sub

    Private Sub Button16_Click_1(sender As Object, e As EventArgs) Handles Button16.Click
        If Module1.listaDevolucionFinal.Count = 0 Then
            MsgBox("NO HAY NINGUN PRODUCTO PARA DEVOLVER")
        Else
            Module1.miGestionBD.DevolverProductosTicket(Module1.listaDevolucionFinal, TextBox1.Text)
            Module1.listaDevolucionFinal.Clear()
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            MsgBox("PRODUCTOS DEVUELTOS")
        End If

    End Sub

    Private Sub Button17_Click_1(sender As Object, e As EventArgs) Handles Button17.Click


        If Not ListBox2.Text = "" Then   'TENGO QUE PENSAR ESTO MEJOR
            listaDevolucionFinal.Remove(ListBox2.SelectedItem)
            ListBox1.Items.Add(ListBox2.SelectedItem)
            ListBox2.Items.Clear()
            ListBox2.Items.AddRange(listaDevolucionFinal.ToArray)
        Else
            MsgBox("DEBE SELECCIONAR UN PRODUCTO PARA QUITAR DE DEVOLUCION")
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If Not ListBox1.Text = "" Then 'TENGO QUE PENSAR ESTO MEJOR
            Module1.listaDevolucionFinal.Add(ListBox1.SelectedItem)
            ListBox2.Items.Clear()
            ListBox2.Items.AddRange(listaDevolucionFinal.ToArray)

            ListBox1.Items.Remove(ListBox1.SelectedItem)
        Else
            MsgBox("DEBE SELECCIONAR UN PRODUCTO PARA DEVOLVER")
        End If
    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        Module1.listaDevolucionFinal.Clear()
        Me.Close()
        Module1.ListaProductosParaDevolucion.Clear()
    End Sub
End Class