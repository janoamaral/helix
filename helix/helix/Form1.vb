Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim newLog As New Ermac
        newLog.LogLevel = 2
        newLog.ErrorLevel = 2
        newLog.Code = 1
        newLog.Description = "Esto es una prueba"
        newLog.isHidden = True
        newLog.ModuleName = "LogCreation"
        newLog.SubSystem = "UnitTest1"
        newLog.Timestamp = Now
        newLog.Save()
    End Sub
End Class
