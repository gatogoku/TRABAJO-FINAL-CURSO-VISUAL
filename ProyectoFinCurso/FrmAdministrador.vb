Public Class FrmAdministrador



    Public Function RealizarCopiaDeSeguridad() 'REALIZAR COPIA DE SEGURIDAD AÑADIENDO AL NOMBRE LA FECHA 

        Dim fx As Date = Date.Today

        Dim fechaActual As String = fx.Year & "-" & fx.Month & "-" & fx.Day


        If My.Computer.FileSystem.FileExists(
            "COPIAS DE SEGURIDAD/CUASHOP_MODA_" & fechaActual & ".accdb") Then
            MsgBox("USTED YA HA REALIZADO UNA COPIA DE SEGURIDAD HOY")
        Else
            ' Copy the file to a new location without overwriting existing file.
            My.Computer.FileSystem.CopyFile(
        "CUASHOP_MODA.accdb",
        "COPIAS DE SEGURIDAD/CUASHOP_MODA_" & fechaActual & ".accdb",
            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
            MsgBox("COPIA DE SEGURIDAD REALIZADA CON EXITO")
        End If



        ' Copy the file to a new folder and rename it.
        'My.Computer.FileSystem.CopyFile(
        '    "C:\UserFiles\TestFiles\testFile.txt",
        '    "C:\UserFiles\TestFiles2\NewFile.txt",
        '    Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
        '    Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

    End Function

    Public Function RestaurarCopiaDeSeguridad() ' SE MOSTRARA UN MENU CON LA COPIAS DE SEGURIDAD REALIZADAS Y SE PERMITIRA RESTAURAR LA SELECCIONADA

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RealizarCopiaDeSeguridad()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FrmLogin.Close()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        FrmCopiasSeguridad.Show()
    End Sub

    Private Sub FrmAdministrador_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.None
    End Sub
End Class