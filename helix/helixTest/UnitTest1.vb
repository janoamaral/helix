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

    <TestMethod()> Public Sub DbHelixConnection()
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "helix"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .RequireCredentials = False
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            Assert.IsTrue(.TestConnection())
        End With
    End Sub

    <TestMethod()> Public Sub SQLEngineBuilderConnection()
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "master"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .RequireCredentials = False
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            Assert.IsTrue(.TestConnection())
        End With
    End Sub

    <TestMethod()> Public Sub SQLEngineBuilderCreateDB()
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "helix"
            .SQLDbProperties.dbFullPath = "G:\Dev\helix\helix\helix\bin\Debug\"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .RequireCredentials = False
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            Assert.IsTrue(.CreateNewDataBase)
        End With
    End Sub

End Class