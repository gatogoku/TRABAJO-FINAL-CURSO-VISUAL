Imports Entidades

Public Class FrmLogin
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox2.Text = Nothing Then TextBox2.Text = ""
            If Not Module1.miGestionBD.ComprobarUsuario(TextBox1.Text, TextBox2.Text).Equals(Nothing) Then
                Module1.EmpleadoActual = Module1.miGestionBD.ComprobarUsuario(TextBox1.Text, TextBox2.Text)
                If EmpleadoActual.GetType.Name = "Vendedores" Then
                    MessageBox.Show("USTED SE HA LOGUEADO COMO VENDEDOR")
                    FrmInicio.Show()
                    Me.Visible = False

                Else
                    MessageBox.Show("USTED SE HA LOGUEADO COMO ADMINISTRADOR")
                    FrmAdministrador.Show()
                    Me.Visible = False
                End If

            End If
        Catch ex As Exception
            MsgBox("ERROR EN DE LOGGIN O BASE DE DATOS CONSULTE AL ADMINISTRADOR")
            TextBox1.Clear()
            TextBox2.Clear()
        End Try


    End Sub


    Public Function RestaurarCopiaDeSeguridad() ' SE MOSTRARA UN MENU CON LA COPIAS DE SEGURIDAD REALIZADAS Y SE PERMITIRA RESTAURAR LA SELECCIONADA


    End Function

    Private Sub FrmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None
        Me.ShowInTaskbar = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class