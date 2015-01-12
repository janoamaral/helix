Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SQLEngineDelete
    Inherits SQLBase

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
        End Select
    End Sub

    ''' <summary>
    ''' Elimina todos los registros de una tabla
    ''' </summary>
    ''' <returns>TRUE si la operacion se realizo con exito. FALSE si fallo</returns>
    Public Function DeleteAll() As Boolean
        Dim core As New SQLCore
        With core
            .ConnectionString = _connectionString
            .dbType = _dbType
            Select Case _dbType
                Case 0
                    Dim dummy As New List(Of OleDbParameter)
                    Return .ExecuteNonQuery(False, dummy)
                Case 1
                    Dim dummy As New List(Of SqlParameter)
                    Return .ExecuteNonQuery(False, dummy)
                Case Else
                    Return False
            End Select
        End With
    End Function

    ''' <summary>
    ''' Elimina registros segun se especifique en la clausula WHERE
    ''' </summary>
    ''' <returns>TRUE Si la operacion fue un exito, FALSE si se omitio la clausula WHERE o fallo la operacion de borrado</returns>
    Public Function Delete() As Boolean
        If _WHEREstring.Length <> 0 Then
            Dim core As New SQLCore
            With core
                .ConnectionString = _connectionString
                .dbType = _dbType
                .QueryString = GenerateQuery(True)
                Select Case _dbType
                    Case 0
                        Return .ExecuteNonQuery(True, _QueryParamOle)
                    Case 1
                        Return .ExecuteNonQuery(True, _QueryParamSql)
                    Case Else
                        Return False
                End Select
            End With
        Else
            Return False    ' Regresa FALSE en el caso que se haya olvidado de poner la clausula WHERE
        End If
    End Function

    ''' <summary>
    ''' Reinicia la clase a sus valores originales
    ''' </summary>
    Public Overrides Sub Reset()
        _QueryParam.Clear()
        _QueryParamOle.Clear()
        _QueryParamSql.Clear()
        _queryString = ""
        _tableName = ""
        _WHEREstring = ""
    End Sub

    ''' <summary>
    ''' Genera una consulta SQL "DELETE" segun los parametros de la clase. Segun se escoja puede devolver una cadena para ser procesada por el SQLCore
    ''' o una cadena para depuracion
    ''' </summary>
    ''' <param name="toProcess">TRUE si la cadena va a ser procesada por SQLCore, FALSE si se usa como depuracion</param>
    ''' <returns>Una consulta SQL "DELETE"</returns>
    Protected Overrides Function GenerateQuery(ByVal toProcess As Boolean) As String
        Dim tmpQuery As String

        tmpQuery = "DELETE FROM " & _tableName
        If _WHEREstring.Length <> 0 Then
            tmpQuery &= " WHERE (" & _WHEREstring & ")"
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
