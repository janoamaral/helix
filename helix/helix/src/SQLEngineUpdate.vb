Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SQLEngineUpdate
    Inherits SQLBase

    ''' <summary>
    ''' Ruta completa y nombre de archivo donde se van a guardar los logs
    ''' </summary>
    ''' <value>Cadena con la ruta completa y el nombre de archivo del log</value>
    ''' <returns>La ruta y el nombre del archivo log</returns>
    ''' <remarks></remarks>
    Public Property LogFileFullName As String = Application.StartupPath & "\syslog.log"

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
        _QueryParamOle.Clear()
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
                        tmpSET &= tmpUpdateValues & "=" & _QueryParamOle(i).Value.ToString & ", "
                End Select

            Else
                tmpSET &= tmpUpdateValues & "= ?, "
            End If
            i += 1
        Next

        tmpQuery &= tmpSET.Remove(tmpSET.Length - 2, 2) ' quita el ultimo , y espacio

        If _WHEREstring.Length <> 0 Then
            tmpQuery &= " WHERE (" & _WHEREstring & ")" ' Agrega la seccion WHERE
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
