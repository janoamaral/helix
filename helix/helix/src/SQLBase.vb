Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public MustInherit Class SQLBase

    ''' <summary>
    ''' Estructura que contiene el par (Columna,Valor)
    ''' </summary>
    Protected Structure ColumnValue
        Dim column As String
        Dim value As Object
    End Structure

    ''' <summary>
    ''' Cadena de conexion a la base de datos
    ''' </summary>
    Protected _connectionString As String = ""

    ''' <summary>
    ''' Nombre de la tabla a trabajar
    ''' </summary>
    Protected _tableName As String = ""

    ''' <summary>
    ''' Tipo de base de datos
    ''' </summary>
    Protected _dbType As Integer = 0

    ''' <summary>
    ''' Cadena conteniendo la clausula WHERE.
    ''' </summary>
    ''' <remarks>El formato es "(campo1 = ?) AND (campo2 = ?)"</remarks>
    Protected _WHEREstring As String = ""

    ''' <summary>
    ''' Lista de objetos para ser pasados a la consulta y reemplazados en la clausula WHERE
    ''' </summary>
    ''' <remarks></remarks>
    Protected _QueryParam As New List(Of Object)


    Protected _QueryParamOle As New List(Of OleDbParameter)

    Protected _QueryParamSql As New List(Of SqlParameter)

    ''' <summary>
    ''' La consulta que se va a relizar contra la base de datos
    ''' </summary>
    Protected _queryString As String = ""

    ''' <summary>
    ''' Guarda o retorna la cadena de conexion a la base de datos segun los parametros ingresados
    ''' </summary>
    ''' <returns>La cadena de conexion a la base de datos</returns>
    Public Property ConnectionString As String
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
    ''' Retorna la cadena con la consulta SQL que se va a hacer contra la base de datos
    ''' </summary>
    ''' <returns>La consulta SQL</returns>
    Public ReadOnly Property SqlQueryString As String
        Get
            Return GenerateQuery(False)
        End Get
    End Property

    ''' <summary>
    ''' Guarda o retorna el nombre de la tabla a trabajar
    ''' </summary>
    ''' <returns>El nombre de la tabla utilizada actualmente</returns>
    Public Property TableName As String
        Get
            Return _tableName
        End Get
        Set(value As String)
            _tableName = value
        End Set
    End Property

  

    ''' <summary>
    ''' Reinicia la clase a sus valores originales
    ''' </summary>
    Public MustOverride Sub Reset()

    ''' <summary>
    ''' Genera una consulta SQL segun los parametros de la clase. Segun se escoja puede devolver una cadena para ser procesada por el SQLCore
    ''' o una cadena para depuracion
    ''' </summary>
    ''' <param name="toProcess">TRUE si la cadena va a ser procesada por SQLCore, FALSE si se usa como depuracion</param>
    ''' <returns>Una consulta SQL</returns>
    Protected MustOverride Function GenerateQuery(ByVal toProcess As Boolean) As String


End Class
