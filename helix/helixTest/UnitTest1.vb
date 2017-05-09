Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports helix

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub DbConnection()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = SQLEngine.dataBaseType.MYSQL
        a.Path = "200.42.62.140"
        a.Username = "soccam"
        a.Password = "Camara2017!"
        a.Port = 3306
        Debug.Print(a.ConnectionString)
        Assert.IsTrue(a.Start)
    End Sub
End Class