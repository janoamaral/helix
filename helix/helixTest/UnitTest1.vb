Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports helix

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub DbConnection()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = 1
        a.Path = "ALPHACORE\SQLEXPRESS"
        a.RequireCredentials = False
        Assert.IsTrue(a.Start)
    End Sub

    <TestMethod()> Public Sub db()

    End Sub

End Class