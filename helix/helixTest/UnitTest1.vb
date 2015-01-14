Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports helix

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub DbConnection()
        Dim a As New SQLEngine
        a.DatabaseName = "master"
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

    <TestMethod()> Public Sub DbTableCreation()
        Dim a As New SQLEngineBuilder
        With a
            .DataBaseName = "helix"
            .DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
            .ModelPath = "G:\Dev\helix\helix\script_test.txt"
            .RequireCredentials = False
            .ServerName = My.Computer.Name & "\SQLEXPRESS"
            Assert.IsTrue(.CreateTable)
        End With
    End Sub

    <TestMethod()> Public Sub LogCreation()
        Dim newLog As New Ermac
        newLog.LogFilePath = "G:\Dev\helix\helix\syslog.log"
        newLog.ErrorLevel = 2
        newLog.Code = 1
        newLog.Description = "Esto es una prueba"
        newLog.isHidden = True
        newLog.ModuleName = "LogCreation"
        newLog.SubSystem = "UnitTest1"
        newLog.Timestamp = Now
        Assert.IsTrue(newLog.Save())
    End Sub

End Class