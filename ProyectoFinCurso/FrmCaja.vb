Public Class FrmCaja
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text = "CAJA DEL DIA: " & Module1.miGestionBD.TotalVentasDia.ToString & " $"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label1.Text = "CAJA TOTAL; " & Module1.miGestionBD.TotalVentasTotal().ToString & " $"
    End Sub

    Private Sub FrmCaja_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None
    End Sub
End Class