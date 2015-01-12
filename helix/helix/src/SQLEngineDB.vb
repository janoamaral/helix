Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SQLEngineDB

    ''' <summary>
    ''' Cadena de conexion a la base de datos
    ''' </summary>
    Private _connectionString As String = ""

    ''' <summary>
    ''' Tipo de base de datos
    ''' </summary>
    Private _dbType As Integer = 0

    ''' <summary>
    ''' La informacion de la base de datos
    ''' </summary>
    ''' <remarks>Almacena un schema de la base de datos para extraccion de informacion de la propia base (cantidad de tablas, etc)</remarks>
    Private _dbInfo As DataTable

    ''' <summary>
    ''' Guarda o retorna la cadena de conexion a la base de datos segun los parametros ingresados
    ''' </summary>
    ''' <returns>La cadena de conexion a la base de datos</returns>
    Public Property ConnectionString() As String
        Get
            Return _connectionString
        End Get
        Set(value As String)
            _connectionString = value
        End Set
    End Property


    ''' <summary>
    ''' Guarda o retorna el tipo de base de datos con la que se va a trabajar
    ''' </summary>
    ''' <value>Entero mayor que 0 indicando el tipo de base de datos. MS Access = 0 / SQL Server = 1</value>
    ''' <returns>El tipo de base de datos a trabajar</returns>
    Public Property DbType As Integer
        Get
            Return _dbType
        End Get
        Set(value As Integer)
            _dbType = value
        End Set
    End Property


    ''' <summary>
    ''' Prueba que la conexion a la base de datos sea correcta
    ''' </summary>
    ''' <returns>TRUE si la conexion se realizo con exito, FALSE si no se puedo conectar</returns>
    Public Function TestDbConnection() As Boolean
        Dim _core As New SQLCore
        With _core
            .ConnectionString = _connectionString
            .dbType = _dbType
            Return .TestConnection()
        End With
    End Function


    ''' <summary>
    ''' Reinicia el objeto a sus valores originales
    ''' </summary>
    Public Sub Reset()
        _connectionString = ""
        _dbType = 0
        _dbInfo.Clear()
    End Sub


    ''' <summary>
    ''' Adquiere el esquema de la base de datos para poder ser procesada
    ''' </summary>
    ''' <returns>True si la transaccion es correcta. False si fallo</returns>
    Public Function GetDbInfo() As Boolean
        Select Case _dbType
            Case 0
                Dim oleDbconnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection()
                oleDbconnection.ConnectionString = _connectionString

                Try
                    oleDbconnection.Open()
                    ' Get list of user tables
                    ' Para saber mas del uso de restricciones: http://msdn.microsoft.com/es-es/library/cc716722.aspx
                    _dbInfo = oleDbconnection.GetSchema("Tables", New String() {Nothing, Nothing, Nothing, "TABLE"})
                    oleDbconnection.Close()
                Catch ex As Exception
                    Return False
                End Try
            Case 1

                Dim sqlDbconnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection()
                sqlDbconnection.ConnectionString = _connectionString
                Try
                    sqlDbconnection.Open()
                    ' Get list of user tables
                    ' Para saber mas del uso de restricciones: http://msdn.microsoft.com/es-es/library/cc716722.aspx
                    _dbInfo = sqlDbconnection.GetSchema("Tables", New String() {Nothing, Nothing, Nothing, "TABLE"})
                    sqlDbconnection.Close()
                Catch ex As Exception
                    Return False
                End Try

            Case Else
                Return False
        End Select
        Return True
    End Function


    ''' <summary>
    ''' Traduce de un indice a una cadena de caracteres con el nombre de la tabla
    ''' </summary>
    ''' <param name="index">Indice de la tabla a ser traducida</param>
    ''' <returns>El nombre de la tabla si el indice es correcto, cadena nula en caso de fallar</returns>
    Public Function GetTableName(ByVal index As Integer) As String
        If (index >= 0) And (index < _dbInfo.Rows.Count) Then
            Return _dbInfo.Rows(index)(2).ToString
        Else
            Return ""
        End If
    End Function


    ''' <summary>
    ''' Indica la cantidad de tablas de usuario contenida en la base
    ''' </summary>
    ''' <returns>La cantidad de tablas de usuarios en la base de datos</returns>
    Public Function TablesCount() As Integer
        Try
            Return _dbInfo.Rows.Count
        Catch ex As Exception
            Return 0
        End Try
    End Function
End Class
