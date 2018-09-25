Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient

Public Class SQLEngineUpdate
    Inherits SQLBase


    Public Enum OperatorCriteria As Byte
        Igual = 0
        Distinto = 1
        Menor = 2
        MenorIgual = 3
        Mayor = 4
        MayorIgual = 5
        LikeString = 6
        Between = 7
    End Enum


    ''' <summary>
    ''' Ruta completa y nombre de archivo donde se van a guardar los logs
    ''' </summary>
    ''' <value>Cadena con la ruta completa y el nombre de archivo del log</value>
    ''' <returns>La ruta y el nombre del archivo log</returns>
    ''' <remarks></remarks>
    Public Property LogFileFullName As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & "syslog.log"

    ''' <summary>
    ''' Lista de parametros Columna/Valor para ser insertados en la tabla
    ''' </summary>
    ''' <remarks></remarks>
    Private _listOfUpdate As New List(Of String)

    ''' <summary>
    ''' Guarda o retorna la cadena con la clausula WHERE
    ''' </summary>
    ''' <returns>La cadena con la clausula WHERE actual</returns>
    Public Property WHEREstring As String
        Get
            Return _WHEREstring
        End Get
        Set(value As String)
            _WHEREstring = value
        End Set
    End Property

    Public Sub AddColumnValue(ByVal column As String, ByVal value As Object)
        _listOfUpdate.Add(column)

        Select Case _dbType
            Case 0
                Dim oleparam As New OleDbParameter
                oleparam.Value = value
                oleparam.ParameterName = "@p" & _QueryParamOle.Count
                _QueryParamOle.Add(oleparam)
            Case 1
                Dim sqlparam As New SqlParameter
                sqlparam.Value = value
                sqlparam.ParameterName = "@p" & _QueryParamSql.Count
                _QueryParamSql.Add(sqlparam)
            Case 2
                Dim mySqlparam As New mySqlParameter
                mySqlparam.Value = value
                mySqlparam.ParameterName = "@p" & _QueryParamMySql.Count
                _QueryParamMySql.Add(mySqlparam)
        End Select

    End Sub

    ''' <summary>
    ''' Agrega un nuevo elemento a la lista de parametros WHERE
    ''' </summary>
    ''' <param name="param">Un objeto para ser usado posteriormente en un comando SQL</param>
    Public Sub AddWHEREparam(ByVal param As Object)
        Select Case _dbType
            Case 0
                Dim oleparam As New OleDbParameter
                oleparam.Value = param
                oleparam.ParameterName = "@p" & _QueryParamOle.Count
                _QueryParamOle.Add(oleparam)
            Case 1
                Dim sqlparam As New SqlParameter
                sqlparam.Value = param
                sqlparam.ParameterName = "@p" & _QueryParamSql.Count
                _QueryParamSql.Add(sqlparam)
            Case 2
                Dim mySqlparam As New MySqlParameter
                mySqlparam.Value = param
                mySqlparam.ParameterName = "@p" & _QueryParamMySql.Count
                _QueryParamMySql.Add(mySqlparam)
        End Select
    End Sub

    ''' <summary>
    ''' Reinicia la clase a sus valores originales
    ''' </summary>
    Public Overrides Sub Reset()
        _QueryParam.Clear()
        _queryString = ""
        _tableName = ""
        _WHEREstring = ""
        _listOfUpdate.Clear()
        _QueryParamSql.Clear()
        _QueryParamMySql.Clear()
        _QueryParamOle.Clear()
    End Sub

    ''' <summary>
    ''' Agrega un query simple del formato COLUMNA operador VALOR [AND VALOR]
    ''' </summary>
    ''' <param name="column">Columna a buscar</param>
    ''' <param name="searchOperator">Operador a utilizar: =, !=...</param>
    ''' <param name="value">Valor a buscar</param>
    ''' <param name="valueEnd">Opcional cuando se usa BETWEEN</param>
    ''' <remarks></remarks>
    Public Sub SimpleSearch(ByVal column As String, ByVal searchOperator As OperatorCriteria, ByVal value As Object, Optional ByVal valueEnd As Object = Nothing)
        _WHEREstring = column
        Select Case searchOperator
            Case OperatorCriteria.Igual
                _WHEREstring += " = ?"
            Case OperatorCriteria.Distinto
                _WHEREstring += " <> ?"
            Case OperatorCriteria.Menor
                _WHEREstring += " < ?"
            Case OperatorCriteria.MenorIgual
                _WHEREstring += " <= ?"
            Case OperatorCriteria.Mayor
                _WHEREstring += " > ?"
            Case OperatorCriteria.MayorIgual
                _WHEREstring += " >= ?"
            Case OperatorCriteria.LikeString
                _WHEREstring += " LIKE ?"
            Case OperatorCriteria.Between
                _WHEREstring += " BETWEEN ? AND ?"
                AddWHEREparam(value)
                AddWHEREparam(valueEnd)
                Exit Sub
        End Select
        AddWHEREparam(value)
    End Sub

    ''' <summary>
    ''' Actualiza los registros de una tabla
    ''' </summary>
    ''' <returns>El estado de la operacion. TRUE si la operacion fue un exito, FALSE si fallo</returns>
    Public Function Update() As Boolean
        Dim _sqlCore As New SQLCore
        With _sqlCore
            .LastError.LogFilePath = _LogFileFullName
            .ConnectionString = _connectionString
            .dbType = _dbType
            .QueryString = GenerateQuery(True)
            Select Case _dbType
                Case 0
                    Return .ExecuteNonQuery(True, _QueryParamOle)
                Case 1
                    Return .ExecuteNonQuery(True, _QueryParamSql)
                Case 2
                    Return .ExecuteNonQuery(True, _QueryParamMySql)
                Case Else
                    Return False
            End Select
        End With


    End Function

    ''' <summary>
    ''' Genera una consulta SQL "INSERT" segun los parametros de la clase. Segun se escoja puede devolver una cadena para ser procesada por el SQLCore
    ''' o una cadena para depuracion
    ''' </summary>
    ''' <param name="toProcess">TRUE si la cadena va a ser procesada por SQLCore, FALSE si se usa como depuracion</param>
    ''' <returns>Una consulta SQL "UPDATE"</returns>
    Protected Overrides Function GenerateQuery(toProcess As Boolean) As String
        Dim tmpQuery As String = ""
        Dim tmpUpdateValues As String
        Dim tmpSET As String = ""

        tmpQuery = "UPDATE " & _tableName & " SET "

        Dim i As Integer = 0

        For Each tmpUpdateValues In _listOfUpdate
            If toProcess = False Then
                Select Case _dbType
                    Case 0
                        tmpSET &= tmpUpdateValues & "=" & _QueryParamOle(i).Value.ToString & ", "
                    Case 1
                        tmpSET &= tmpUpdateValues & "=" & _QueryParamSql(i).Value.ToString & ", "
                    Case 2
                        tmpSET &= tmpUpdateValues & "=" & _QueryParamMySql(i).Value.ToString & ", "
                End Select

            Else
                tmpSET &= tmpUpdateValues & "=?, "
            End If
            i += 1
        Next

        tmpQuery &= tmpSET.Remove(tmpSET.Length - 2, 2) ' quita el ultimo , y espacio

        If _WHEREstring.Length <> 0 Then
            Select Case _dbType
                Case 0
                    tmpQuery &= " WHERE " & _WHEREstring
                Case 1
                    tmpQuery &= " WHERE (" & _WHEREstring & ")" ' Agrega la seccion WHERE
                Case 2
                    tmpQuery &= " WHERE " & _WHEREstring
            End Select

        End If

        If toProcess = False Then
            Dim obj As Object
            For Each obj In _QueryParam                                             ' Por cada parametro
                tmpQuery = tmpQuery.Insert(tmpQuery.IndexOf("?"), obj.ToString)     ' Se reemplaza en la consulta final
                tmpQuery = tmpQuery.Remove(tmpQuery.IndexOf("?"), 1)                ' Dando la consulta lista para depuracion
            Next
        End If

        Return tmpQuery
    End Function
End Class
