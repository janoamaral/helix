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

    <TestMethod()> Public Sub MySqlQuery()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = SQLEngine.dataBaseType.MYSQL
        a.Path = "200.42.62.140"
        a.Username = "soccam"
        a.Password = "Camara2017!"
        a.Port = 3306
        a.Start()
        With a.Query
            .Reset()
            .TableName = "SOCIO"
            .AddSelectColumn("*")
            .WHEREstring = "1=1"
            If .Query Then
                Assert.AreNotEqual(.RecordCount, 0)
            Else
                Assert.Fail()
            End If
        End With
    End Sub

    <TestMethod()> Public Sub MySqlInsert()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = SQLEngine.dataBaseType.MYSQL
        a.Path = "200.42.62.140"
        a.Username = "soccam"
        a.Password = "Camara2017!"
        a.Port = 3306
        a.Start()
        With a.Insert
            .Reset()
            .TableName = "SOCIO"
            .AddColumnValue("SOCIO_ID", 81118)
            .AddColumnValue("SOCIO_NOMBRE", "PRUEBA")
            .AddColumnValue("SOCIO_APELLIDO", "INSERT DESDE SOCCAM")
            .AddColumnValue("SOCIO_DELETED", 1)
            .AddColumnValue("SOCIO_ESTADO", 0)
            .AddColumnValue("SOCIO_MODIFICADO", Now())
            Debug.Print(.SqlQueryString)
            Assert.IsTrue(.Insert())
        End With
    End Sub

    <TestMethod()> Public Sub MySqlInsertWithReturn()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = SQLEngine.dataBaseType.MYSQL
        a.Path = "200.42.62.140"
        a.Username = "soccam"
        a.Password = "Camara2017!"
        a.Port = 3306
        a.Start()
        With a.Insert
            .Reset()
            .TableName = "SOCIO"
            .AddColumnValue("SOCIO_ID", 81127)
            .AddColumnValue("SOCIO_NOMBRE", "TEST INSERT 81127")
            .AddColumnValue("SOCIO_APELLIDO", "CON RETORNO ID")
            .AddColumnValue("SOCIO_DELETED", 1)
            .AddColumnValue("SOCIO_ESTADO", 0)
            .AddColumnValue("SOCIO_MODIFICADO", Now())
            Dim newId As Integer
            .Insert(newId)
            Assert.AreEqual(0, newId)
        End With
    End Sub

    <TestMethod()> Public Sub MySqlUpdate()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = SQLEngine.dataBaseType.MYSQL
        a.Path = "200.42.62.140"
        a.Username = "soccam"
        a.Password = "Camara2017!"
        a.Port = 3306
        a.Start()
        With a.Update
            .Reset()
            .TableName = "SOCIO"
            .AddColumnValue("SOCIO_NOMBRE", "TEST UPDATE")
            .AddColumnValue("SOCIO_APELLIDO", "DESDE SOCCAM")
            .AddColumnValue("SOCIO_DELETED", 1)
            .AddColumnValue("SOCIO_ESTADO", 0)
            .AddColumnValue("SOCIO_MODIFICADO", Now())
            .WHEREstring = "SOCIO_ID >= ?"
            .AddWHEREparam(81119)

            Debug.Print(.SqlQueryString)
            Assert.IsTrue(.Update)
        End With
    End Sub
End Class