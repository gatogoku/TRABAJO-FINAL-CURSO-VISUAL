Public Class FrmCopiasSeguridad

    Dim ListaCopiasDeSeguridad As New List(Of String)

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If ListBox1.SelectedItem = "" Then
            MsgBox("NO SE HA SELECCIONADO NINGUN ARCHIVO PARA RESTAURAR")
        Else

            My.Computer.FileSystem.CopyFile(
      ListBox1.SelectedItem.ToString, "CUASHOP_MODA.accdb",
       Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
       Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

        End If

    End Sub

    Private Sub FrmCopiasSeguridad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None
        'ListaCopiasDeSeguridad.AddRange(System.IO.Directory.GetFiles("COPIAS DE SEGURIDAD"))
        'ListBox1.Items.AddRange(ListaCopiasDeSeguridad.ToArray)
        ListBox1.Items.AddRange(System.IO.Directory.GetFiles("COPIAS DE SEGURIDAD"))


    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class