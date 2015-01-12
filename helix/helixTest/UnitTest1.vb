Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports helix

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestMethod1()
        Dim a As New SQLEngine
        a.DatabaseName = "soccam"
        a.dbType = 1

    End Sub

End Class