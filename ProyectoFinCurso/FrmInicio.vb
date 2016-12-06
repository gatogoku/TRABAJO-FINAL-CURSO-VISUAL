Imports GestionDB, Entidades
Public Class FrmInicio


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None

        Label1.Left = GroupBox1.Width / 2 - Label1.Width / 2
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        frmVentas.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FrmDevoluciones.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FrmReservas.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        FrmCaja.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FrmLogin.Close()
        Me.Close()
    End Sub


End Class
