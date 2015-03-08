Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SQLEngineInsert
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
    Private _listOfInsert As New List(Of String)

    ''' <summary>
    ''' Ingresa un nuevo par Column/Value a la lista de parametros para usar en la consulta INSERT
    ''' </summary>
    ''' <param name="column">Nombre de la columna</param>
    ''' <param name="value">Valor a ser insertado</param>
    Public Sub AddColumnValue(ByVal column As String, ByVal value As Object)

        _listOfInsert.Add(column)

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
    ''' Crea una nuevo registro en la tabla
    ''' </summary>
    ''' <returns>El estado de la operacion. TRUE si la operacion fue un exito, FALSE si fallo</returns>
    Public Overloads Function Insert() As Boolean
        If _listOfInsert.Count <> 0 Then

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

        Else
            Return False
        End If

    End Function

    ''' <summary>
    ''' Inserta un nuevo registro en la base de datos y devuelve el ultimo ID creado
    ''' </summary>
    ''' <param name="lastID">Contenedor donde se devuelve el ultimo ID creado</param>
    ''' <returns>El estado de la operacion. TRUE si la operacion fue un exito, FALSE si fallo</returns>
    Public Overloads Function Insert(ByRef lastID As Long) As Boolean
        If _listOfInsert.Count <> 0 Then

            Dim _sqlCore As New SQLCore
            With _sqlCore
                .LastError.LogFilePath = _LogFileFullName
                .ConnectionString = _connectionString
                .dbType = _dbType
                .QueryString = GenerateQuery(True)
                Select Case _dbType
                    Case 0
                        Return .ExecuteNonQuery(True, _QueryParamOle, lastID)
                    Case 1
                        Return .ExecuteNonQuery(True, _QueryParamSql, lastID)
                    Case Else
                        Return False
                End Select
            End With

        Else
            Return False
        End If

    End Function

    ''' <summary>
    ''' Reinicia la clase a sus valores originales
    ''' </summary>
    Public Overrides Sub Reset()
        _listOfInsert.Clear()
        _QueryParamOle.Clear()
        _QueryParamSql.Clear()
        _QueryParam.Clear()
        _queryString = ""
        _WHEREstring = ""
    End Sub

    ''' <summary>
    ''' Genera una consulta SQL "INSERT" segun los parametros de la clase. Segun se escoja puede devolver una cadena para ser procesada por el SQLCore
    ''' o una cadena para depuracion
    ''' </summary>
    ''' <param name="toProcess">TRUE si la cadena va a ser procesada por SQLCore, FALSE si se usa como depuracion</param>
    ''' <returns>Una consulta SQL "INSERT"</returns>
    Protected Overrides Function GenerateQuery(ByVal toProcess As Boolean) As String
        Dim tmpQuery As String

        tmpQuery = "INSERT INTO " & _tableName & " "
        Dim columns As String = "("
        Dim values As String = "("
        Dim tmpFieldValue As String
        Dim i As Integer = 0

        For Each tmpFieldValue In _listOfInsert             ' Por cada parametro
            columns &= tmpFieldValue & ", "  ' Los campos sea para proceso o para depuracion siempre van a ser en texto plano
            If toProcess = False Then
                values &= tmpFieldValue & ", "
            Else
                values &= "?, "
            End If
            i += 1
        Next

        columns = columns.Remove(columns.Length - 2, 2) & ")"
        values = values.Remove(values.Length - 2, 2) & ")"
        Return tmpQuery & columns & " VALUES " & values

    End Function

End Class
