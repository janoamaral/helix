Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "helix"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .RequireCredentials = False
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            .CreateNewDataBase()
        End With
    End Sub
End Class
