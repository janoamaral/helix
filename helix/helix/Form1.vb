Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "helix"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .RequireCredentials = False
            .ModelPath = "G:\Dev\helix\helix\script_test.txt"
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            .CreateTable()
        End With
    End Sub
End Class
