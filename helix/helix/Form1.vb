Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "soccam"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .ModelPath = "G:\Dev\helix\helix\script_test.txt"
            .RequireCredentials = False
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            MsgBox(.GenerateConnectionString(True))
            .CreateTable()
        End With

    End Sub
End Class
