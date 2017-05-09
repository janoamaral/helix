Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

Public Class SQLCore

    Private _connectionString As String = ""

    Private _tableName As String = ""

    Private _dbType As Integer = 0

    Private _queryString As String = ""

    ''' <summary>
    ''' Controlador de error
    ''' </summary>
    ''' <remarks></remarks>
    Public LastError As New Ermac

    Public Property ConnectionString As String
        Get
            Return _connectionString
        End Get
        Set(value As String)
            _connectionString = value
        End Set
    End Property


    ''' <summary>
    ''' Setea o devuelve el tipo de base de datos
    ''' </summary>
    ''' <value>0 si es MS Access / 1 si es SQL Server / 2 mysql</value>
    ''' <returns>El tipo de base de datos actual</returns>
    Public Property dbType As Integer
        Get
            Return _dbType
        End Get
        Set(value As Integer)
            _dbType = value
        End Set
    End Property


    ''' <summary>
    ''' Testea la conexion a la base de datos
    ''' </summary>
    ''' <returns>TRUE si la conexion tuvo exito. FALSE si falla</returns>
    Public Function TestConnection() As Boolean
        Select Case _dbType
            Case 0                                                              ' Modo OleDB
                Using connection As New OleDbConnection(_connectionString)
                    Try
                        connection.Open()
                    Catch ex As Exception
                        Debug.Print(ex.Message)
                        LastError.SetError(ex, "SQLCore", "TestConnectionMSA")
                        Return False
                    End Try
                    Return True
                End Using
            Case 1                                                              ' Modo SQL Server
                Using connection As New SqlConnection(_connectionString)
                    Try
                        connection.Open()
                    Catch ex As Exception
                        LastError.SetError(ex, "SQLCore", "TestConnectionSQL")
                        Return False
                    Finally
                        connection.Close()
                    End Try
                    Return True
                End Using
            Case 2
                Using connection As New MySqlConnection(_connectionString)
                    Try
                        connection.Open()
                    Catch ex As Exception
                        LastError.SetError(ex, "SQLCore", "TestConnectionMySql")
                        Return False
                    Finally
                        connection.Close()
                        connection.Dispose()
                    End Try
                    Return True
                End Using

            Case Else                                                           ' Cualquier otro modo
                Return False
        End Select
    End Function


    ''' <summary>
    ''' Reinicia todas las variables a sus valores originales
    ''' </summary>
    Public Sub Reset()
        _connectionString = ""
        _tableName = ""
        _dbType = 0
        _queryString = ""
    End Sub


    ''' <summary>
    ''' Setea o devuelve la cadena con la consulta a ejecutar
    ''' </summary>
    ''' <value>Cadena con consulta SQL</value>
    ''' <returns>La consulta SQL</returns>
    Public Property QueryString As String
        Get
            Return _queryString
        End Get
        Set(value As String)
            _queryString = value
        End Set
    End Property


    ''' <summary>
    ''' Ejecuta una consulta que no requiere devolver datos (ej: dataReader, SQLSchema, etc)
    ''' </summary>
    ''' <param name="processParam">Flag indicando si hay parametros para ser procesados</param>
    ''' <param name="Param">Lista de parametros</param>
    ''' <returns>El resultado de la operacion. TRUE se ejecuto con exito, FALSE fallo</returns>
    Public Overloads Function ExecuteNonQuery(ByVal processParam As Boolean, Optional ByVal Param As List(Of OleDbParameter) = Nothing) As Boolean
        Using connection As New OleDbConnection(_connectionString)
            Dim command As New OleDbCommand(_queryString, connection)

            If processParam = True Then
                Dim tmpParam As OleDbParameter
                Dim tmpPos As Integer
                For Each tmpParam In Param
                    tmpPos = command.CommandText.IndexOf("?")
                    command.CommandText = command.CommandText.Remove(tmpPos, 1)
                    command.CommandText = command.CommandText.Insert(tmpPos, tmpParam.ParameterName)
                    command.Parameters.AddWithValue(tmpParam.ParameterName, tmpParam.Value)
                Next
            End If

            Try
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
                Return True
            Catch ex As Exception
                Console.Write(ex.Message)
                LastError.SetError(ex, "SQLCore", "ExecuteNonQuery")
                Return False
            End Try
        End Using
    End Function


    ''' <summary>
    ''' Ejecuta una consulta que no requiere devolver datos (ej: dataReader, SQLSchema, etc)
    ''' </summary>
    ''' <param name="processParam">Flag indicando si hay parametros para ser procesados</param>
    ''' <param name="Param">Lista de parametros</param>
    ''' <returns>El resultado de la operacion. TRUE se ejecuto con exito, FALSE fallo</returns>
    Public Overloads Function ExecuteNonQuery(ByVal processParam As Boolean, Optional ByVal Param As List(Of SqlParameter) = Nothing) As Boolean
        Using connection As New SqlConnection(_connectionString)
            Dim command As New SqlCommand(_queryString, connection)

            If processParam = True Then
                Dim tmpParam As SqlParameter
                Dim tmpPos As Integer
                For Each tmpParam In Param
                    tmpPos = command.CommandText.IndexOf("?")
                    If tmpPos >= 0 Then
                        command.CommandText = command.CommandText.Remove(tmpPos, 1)
                        command.CommandText = command.CommandText.Insert(tmpPos, tmpParam.ParameterName)
                        command.Parameters.AddWithValue(tmpParam.ParameterName, tmpParam.Value)
                    Else
                        Return False
                    End If
                Next
            End If

            Try
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
                Return True
            Catch ex As Exception
                Console.Write(ex.Message)
                LastError.SetError(ex, "SQLCore", "ExecuteNonQuery2")
                Return False
            End Try
        End Using
    End Function


    ''' <summary>
    ''' Ejecuta una consulta en MS Access que requiere devolver datos (ultimo ID creado)
    ''' </summary>
    ''' <param name="processParam">Flag indicando si hay parametros para ser procesados</param>
    ''' <param name="Param">Lista de parametros</param>
    ''' <param name="lastID">Contenedor donde va a ser devuelto el ultimo ID creado</param>
    ''' <returns>El resultado de la operacion. TRUE se ejecuto con exito, FALSE fallo</returns>
    Public Overloads Function ExecuteNonQuery(ByVal processParam As Boolean, ByVal Param As List(Of OleDbParameter), ByRef lastID As Long) As Boolean
        Using connection As New OleDbConnection(_connectionString)
            Dim command As New OleDbCommand(_queryString, connection)

            ' Llamada a la funcion en la base LAST_INSERT_ID() que devuelve el ultimo ID
            Dim tmpQueryString As String = _queryString

            If processParam = True Then
                Dim tmpParam As OleDbParameter
                Dim tmpPos As Integer
                For Each tmpParam In Param
                    tmpPos = command.CommandText.IndexOf("?")
                    command.CommandText = command.CommandText.Remove(tmpPos, 1)
                    command.CommandText = command.CommandText.Insert(tmpPos, tmpParam.ParameterName)
                    command.Parameters.AddWithValue(tmpParam.ParameterName, tmpParam.Value)
                Next
            End If

            Try

                ' TODO: Testear en ms access

                connection.Open()
                If command.ExecuteNonQuery() > 0 Then
                    ' Cambio de ejecutar ExecuteNonQuery a ExecuteScalar para que devuelva el ID
                    command.CommandText = "SELECT @@IDENTITY AS 'Identity'"
                    lastID = Convert.ToInt32(command.ExecuteScalar())
                Else
                    lastID = 0
                End If

                connection.Close()
            Catch ex As Exception
                Console.Write(ex.Message)
                LastError.SetError(ex, "SQLCore", "ExecuteNonQuery3")
                Return False
            End Try

            Return True
        End Using
    End Function


    ''' <summary>
    ''' Ejecuta una consulta en SQL Server que requiere devolver datos (ultimo ID creado)
    ''' </summary>
    ''' <param name="processParam">Flag indicando si hay parametros para ser procesados</param>
    ''' <param name="Param">Lista de parametros</param>
    ''' <param name="lastID">Contenedor donde va a ser devuelto el ultimo ID creado</param>
    ''' <returns>El resultado de la operacion. TRUE se ejecuto con exito, FALSE fallo</returns>
    Public Overloads Function ExecuteNonQuery(ByVal processParam As Boolean, ByVal Param As List(Of SqlParameter), ByRef lastID As Long) As Boolean
        Using connection As New SqlConnection(_connectionString)

            ' Llamada a la funcion en la base LAST_INSERT_ID() que devuelve el ultimo ID
            Dim tmpQueryString As String = _queryString

            Dim command As New SqlCommand(_queryString, connection)

            If processParam = True Then
                Dim tmpParam As SqlParameter
                Dim tmpPos As Integer
                For Each tmpParam In Param
                    tmpPos = command.CommandText.IndexOf("?")
                    command.CommandText = command.CommandText.Remove(tmpPos, 1)
                    command.CommandText = command.CommandText.Insert(tmpPos, tmpParam.ParameterName)
                    command.Parameters.AddWithValue(tmpParam.ParameterName, tmpParam.Value)
                Next
            End If

            Try
                connection.Open()
                ' Cambio de ejecutar ExecuteNonQuery a ExecuteScalar para que devuelva el ID
                command.ExecuteNonQuery()
                command.CommandText = "SELECT @@IDENTITY AS 'Identity'"
                lastID = Convert.ToInt32(command.ExecuteScalar())
                connection.Close()
                Return True
            Catch ex As Exception
                Console.Write(ex.Message)
                lastID = 0
                LastError.SetError(ex, "SQLCore", "ExecuteNonQuery4")
                Return False
            End Try
        End Using
    End Function

    Public Overloads Function ExecuteNonQuery(ByVal commandString As String) As Boolean
        Using connection As New SqlConnection(_connectionString)
            Dim command As New SqlCommand(commandString, connection)

            Try
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
                Return True
            Catch ex As Exception
                Console.Write(ex.Message)
                LastError.SetError(ex, "SQLCore", "ExecuteNonQuery5")
                Return False
            End Try
        End Using
    End Function


    ''' <summary>
    ''' Ejecuta una consulta contra una base de datos
    ''' </summary>
    ''' <param name="processParam">Flag indicando si hay parametros a procesar</param>
    ''' <param name="Param">Lista de parametros a procesar (Clausula WHERE)</param>
    ''' <param name="dbReader">DataReader donde se van a almacenar los registros recuperados</param>
    ''' <returns>El estado de la operacion. TRUE si la operacion fue un exito, FALSE si fallo</returns>
    Public Overloads Function ExecuteQuery(ByVal processParam As Boolean, ByVal Param As List(Of OleDbParameter), ByRef dbReader As DataTable) As Boolean

        Select Case _dbType
            Case 0
                Dim Accessconnection As New OleDbConnection(_connectionString)
                Dim Accesscommand As New OleDbCommand(_queryString, Accessconnection)
                Dim AccessDataReader As OleDbDataReader

                If processParam = True Then
                    Dim tmpParam As OleDbParameter
                    Dim tmpPos As Integer
                    For Each tmpParam In Param
                        tmpPos = Accesscommand.CommandText.IndexOf("?")
                        Accesscommand.CommandText = Accesscommand.CommandText.Remove(tmpPos, 1)
                        Accesscommand.CommandText = Accesscommand.CommandText.Insert(tmpPos, tmpParam.ParameterName)
                        Accesscommand.Parameters.AddWithValue(tmpParam.ParameterName, tmpParam.Value)
                    Next
                End If

                Try
                    Accessconnection.Open()
                    'Accesscommand.Prepare()
                    AccessDataReader = Accesscommand.ExecuteReader()
                    dbReader.Load(AccessDataReader)

                    Accessconnection.Close()
                    Return True
                Catch ex As Exception
                    Console.Write(ex.Message)
                    LastError.SetError(ex, "SQLCore", "ExecuteQuery")
                    Return False
                End Try
            Case Else
                Dim ex As New Exception
                LastError.SetError(ex, "SQLCore", "ExecuteQuery", "No existe tipo DB")
                Return False
        End Select
    End Function


    ''' <summary>
    ''' Ejecuta una consulta contra una base de datos
    ''' </summary>
    ''' <param name="processParam">Flag indicando si hay parametros a procesar</param>
    ''' <param name="Param">Lista de parametros a procesar (Clausula WHERE)</param>
    ''' <param name="dbReader">DataReader donde se van a almacenar los registros recuperados</param>
    ''' <returns>El estado de la operacion. TRUE si la operacion fue un exito, FALSE si fallo</returns>
    Public Overloads Function ExecuteQuery(ByVal processParam As Boolean, ByVal Param As List(Of SqlParameter), ByRef dbReader As DataTable) As Boolean
        Dim SQLconnection As New SqlConnection(_connectionString)
        Dim SQLcommand As New SqlCommand(_queryString, SQLconnection)
        Dim SQLReader As SqlDataReader

        If processParam = True Then
            Dim tmpParam As SqlParameter
            Dim tmpPos As Integer
            For Each tmpParam In Param
                tmpPos = SQLcommand.CommandText.IndexOf("?")
                SQLcommand.CommandText = SQLcommand.CommandText.Remove(tmpPos, 1)
                SQLcommand.CommandText = SQLcommand.CommandText.Insert(tmpPos, tmpParam.ParameterName)
                SQLcommand.Parameters.AddWithValue(tmpParam.ParameterName, tmpParam.Value)
            Next
        End If

        Try
            SQLconnection.Open()
            SQLReader = SQLcommand.ExecuteReader()
            dbReader.Load(SQLReader)                ' Pasa el resultado del datareader al data table para ser procesado y que no se pierda al cerrar la conexion

            SQLconnection.Close()
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            LastError.SetError(ex, "SQLCore", "ExecuteQuery2")
            Return False
        End Try
    End Function

End Class
