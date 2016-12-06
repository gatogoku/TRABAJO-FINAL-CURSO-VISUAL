Imports System.Security.Cryptography
Public Class Encriptacion
    Public Shared Function MD5EncryptPass(ByVal datoEncriptar As String) As String
        ' Método compartido: Shared
        ' Por ser compartido no necesita crear un objeto de la clase Encriptacion para acceder a él, sino que 
        ' Vale con poner: Encriptacion.MD5EncryptPass(variableConDatoAEncriptar) 
        ' Devuelve el valor de la variable encriptado
        Dim PasConMd5 As String = ""
        Dim md5 As New MD5CryptoServiceProvider
        Dim bytValue() As Byte
        Dim bytHash() As Byte
        Dim i As Integer

        bytValue = System.Text.Encoding.UTF8.GetBytes(datoEncriptar)

        bytHash = md5.ComputeHash(bytValue)
        md5.Clear()

        For i = 0 To bytHash.Length - 1
            PasConMd5 &= bytHash(i).ToString("x").PadLeft(2, "0")
        Next

        Return PasConMd5

    End Function
End Class
