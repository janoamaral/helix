Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SQLEngineQuery
    Inherits SQLBase

    ''' <summary>
    ''' Ruta completa y nombre de archivo donde se van a guardar los logs
    ''' </summary>
    ''' <value>Cadena con la ruta completa y el nombre de archivo del log</value>
    ''' <returns>La ruta y el nombre del archivo log</returns>
    ''' <remarks></remarks>
    Public Property LogFileFullName As String = Application.StartupPath & "\syslog.log"

    ''' <summary>
    ''' Almacena las partes de la consulta JOIN
    ''' </summary>
    Private _joinQuery As String = ""

    ''' <summary>
    ''' Flag indicando si el DataReader esta disponible para ser leido
    ''' </summary>
    Private _flagReaderReady As Boolean = False

    ''' <summary>
    ''' Indica la cantidad de JOIN anidados
    ''' </summary>
    Private _joinCount As Integer = 0

    ''' <summary>
    ''' Abstraccion del tipo de base de datos utilizado
    ''' </summary>
    ''' <remarks>Se utiliza esta tabla temporal para almacenar los resultados de las consultas una vez cerrada la conexion a la base de datos</remarks>
    Private _queryResult As New DataTable

    ''' <summary>
    ''' Lector de la consulta
    ''' </summary>
    ''' <remarks></remarks>
    Private _queryResultReader As DataTableReader

    ''' <summary>
    ''' Cantidad de columnas devueltas en la consulta
    ''' </summary>
    ''' <remarks></remarks>
    Private _columnCount As Integer = 0

    ''' <summary>
    ''' Cantidad de registros devueltos en la consulta
    ''' </summary>
    ''' <remarks></remarks>
    Private _recordCount As Integer = 0

    ''' <summary>
    ''' Almacena las columnas de las que se quieren recuperar los registros
    ''' </summary>
    Private _selectColumn As New List(Of String)

    ''' <summary>
    ''' Almacena las columnas que se quieren ordenar
    ''' </summary>
    ''' <remarks></remarks>
    Private _orderColumn As New List(Of String)

    ''' <summary>
    ''' Modo de ordenacion del resultado el query
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum sortOrder
        ascending = 0
        descending = 1
    End Enum

    ''' <summary>
    ''' Agrega la primera clausula JOIN
    ''' </summary>
    ''' <param name="table1">Primera tabla de comparacion</param>
    ''' <param name="table2">Segunda tabla de comparacion</param>
    ''' <param name="commonColumnTable1">Primera columna en comun entre tabla</param>
    ''' <param name="commonColumnTable2">Segunda columna en comun entre tablas</param>
    ''' <remarks>Estructura de un JOIN simple (SELECT * FROM (table1 INNER JOIN table2 ON commonColumnTable1 = commonColumnTable2)
    ''' El @ sera reemplazado por la cantidad de "(" necesarios en la consulta final para que los JOIN anidados cierren sus respectivos parentesis
    ''' </remarks>
    Public Sub AddFirstJoin(ByVal table1 As String, ByVal table2 As String, ByVal commonColumnTable1 As String, ByVal commonColumnTable2 As String)
        _joinQuery = "@(" & table1 & " INNER JOIN " & table2 & " ON " & commonColumnTable1 & " = " & commonColumnTable2 & ")"
    End Sub

    ''' <summary>
    ''' Agrega una columna con el metodo de ordenacion
    ''' </summary>
    ''' <param name="column">Nombre de la columna a ordenar</param>
    ''' <param name="sortingOrder">Metodo de ordenacion</param>
    ''' <remarks></remarks>
    Public Sub AddOrderColumn(ByVal column As String, ByVal sortingOrder As sortOrder)
        Select Case sortingOrder
            Case sortOrder.ascending
                _orderColumn.Add(column & " ASC")
            Case sortOrder.descending
                _orderColumn.Add(column & " DESC")
        End Select
    End Sub


    ''' <summary>
    ''' Agrega una nueva clausula JOIN anidada a la consulta
    ''' </summary>
    ''' <param name="table">Tabla que se agrega a la comparacion</param>
    ''' <param name="commonColumnTable1">Columna en comun con las primeras 2 tablas (agregadas en AddFirstJoin)</param>
    ''' <param name="commonColumnTable2">Columna en comun de la tabla que se agrega a la comparacion</param>
    ''' <remarks>Estructura de una JOIN query compuesta
    ''' SELECT * FROM (tabla1 INNER JOIN tabla2 ON tabla1.columnaComun = tabla2.columnaComun) INNER JOIN tabla3 ON tabla1.columnaComun = tabla3.columnaComun
    ''' Lo que se agrega en esta funcion es el segundo (y sucesivos) INNER JOIN a la consulta
    ''' </remarks>
    Public Sub AddNestedJoin(ByVal table As String, ByVal commonColumnTable1 As String, ByVal commonColumnTable2 As String)
        If _joinQuery.Length <> 0 Then
            _joinQuery &= " INNER JOIN " & table & " ON " & commonColumnTable1 & " = " & commonColumnTable2 & ")"
            _joinCount += 1
        End If
    End Sub


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
    ''' Cantidad de columnas devueltas en la consulta
    ''' </summary>
    ''' <returns>La cantidad de columnas devueltas</returns>
    Public ReadOnly Property ColumnCount() As Integer
        Get
            Return _columnCount
        End Get
    End Property


    ''' <summary>
    ''' Cantidad de registros devueltos por la consulta
    ''' </summary>
    ''' <returns>La cantidad de registros devueltos</returns>
    Public ReadOnly Property RecordCount() As Integer
        Get
            Return _recordCount
        End Get
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
    ''' Agrega una nueva columna a la clausula SELECT
    ''' </summary>
    ''' <param name="column">Nombre de la columna</param>
    ''' <param name="tableName">Nombre de la tabla</param>
    Public Sub AddSelectColumn(ByVal column As String, Optional ByVal tableName As String = "")
        If tableName.Length <> 0 Then column = tableName & "." & column
        _selectColumn.Add(column)
    End Sub


    ''' <summary>
    ''' Lee los datos de la consulta y avanza una posicion en la lista de registros de la consulta actual
    ''' </summary>
    ''' <returns>El resultado de la operacion. TRUE si se leyo correctamente, FALSE si el DataReader no esta preparado o falla la lectura</returns>
    Public Function QueryRead() As Boolean
        If _flagReaderReady = True Then
            Return _queryResultReader.Read()
        Else
            _flagReaderReady = False
            Return False
        End If
    End Function


    ''' <summary>
    ''' Devuelve el valor de la columna en el registro actual
    ''' </summary>
    ''' <param name="index">Numero de columna</param>
    ''' <returns>El valor de la columna en el registro actual. Si falla devuelve FALSE</returns>
    Public Function GetQueryData(ByVal index As Integer) As Object
        If _flagReaderReady = True Then
            If IsDBNull(_queryResultReader.Item(index)) Then
                Return ""
            Else
                Return _queryResultReader.Item(index)
            End If

        Else
            Return False
        End If
    End Function


    ''' <summary>
    ''' Genera una consulta SQL "SELECT" segun los parametros de la clase. Segun se escoja puede devolver una cadena para ser procesada por el SQLCore
    ''' o una cadena para depuracion
    ''' </summary>
    ''' <param name="toProcess">TRUE si la cadena va a ser procesada por SQLCore, FALSE si se usa como depuracion</param>
    ''' <returns>Una consulta SQL "SELECT"</returns>
    Protected Overrides Function GenerateQuery(toProcess As Boolean) As String
        Dim tmpQuery As String = "SELECT "
        Dim tmpStr As String
        If _selectColumn.Count <> 0 Then                                    ' Detecta que se eligen las columnas
            For Each tmpStr In _selectColumn                                ' Agrega las columnas que se van a extraer los registros
                tmpQuery &= tmpStr & ", "
            Next
            tmpQuery = tmpQuery.Remove(tmpQuery.Length - 2, 2)              ' Quitar la coma y el espacio sobrante
        Else
            tmpQuery &= "*"                                                 ' Seleccionar todas las columnas
        End If

        tmpQuery &= " FROM "

        Dim tmpJoin As String = _joinQuery
        If _joinQuery.Length <> 0 Then                                  ' Si la consulta contiene una clausula JOIN
            tmpStr = ""
            If _joinCount <> 0 Then
                Dim i As Integer = 0
                For i = 1 To _joinCount
                    tmpStr &= "("                                           ' Agregar los parentesis necesarios para completar la clausula JOIN
                Next
                tmpJoin = tmpJoin.Insert(tmpJoin.IndexOf("@"), tmpStr)   ' Inserta los parentesis
            End If
            tmpJoin = tmpJoin.Replace("@", "")      ' Eliminar el @
            tmpQuery &= tmpJoin
        Else                                        ' Si la consulta es con solo una tabla
            tmpQuery &= _tableName
        End If


        If _WHEREstring.Length <> 0 Then                    ' Agregar la consulta WHERE
            tmpQuery &= " WHERE " & _WHEREstring
        End If

        If _orderColumn.Count > 0 Then
            tmpQuery &= " ORDER BY "
            For Each tmpStr In _orderColumn
                tmpQuery &= tmpStr & ", "
            Next
            ' Remover la ultima coma y espacio
            tmpQuery = tmpQuery.Remove(tmpQuery.LastIndexOf(","), 2)
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


    ''' <summary>
    ''' Reinicia la clase a sus valores originales
    ''' </summary>
    Public Overrides Sub Reset()
        _flagReaderReady = False
        _joinCount = 0
        _joinQuery = ""
        _QueryParam.Clear()
        _selectColumn.Clear()
        _queryString = ""
        _WHEREstring = ""

        If _queryResult.IsInitialized Then
            _queryResult.Reset()
            If Not IsNothing(_queryResultReader) Then
                If Not (_queryResultReader.IsClosed) Then
                    _queryResultReader.Close()
                End If
            End If
        End If

        _QueryParamOle.Clear()
        _QueryParamSql.Clear()
        _orderColumn.Clear()
        _columnCount = 0
        _recordCount = 0
    End Sub


    ''' <summary>
    ''' Ejecuta la consulta contra la base de datos
    ''' </summary>
    ''' <returns>El resultado de la consulta. TRUE si la consulta se realizo con exito, FALSE si fallo</returns>
    Public Function Query() As Boolean
        Dim core As New SQLCore

        With core
            .LastError.LogFilePath = _LogFileFullName
            .ConnectionString = _connectionString
            .dbType = _dbType
            .QueryString = GenerateQuery(True)

            Debug.Print(.QueryString)

            Select Case _dbType
                Case 0
                    If .ExecuteQuery(True, _QueryParamOle, _queryResult) Then
                        _flagReaderReady = True
                        _columnCount = _queryResult.Columns.Count
                        _recordCount = _queryResult.Rows.Count
                        _queryResultReader = _queryResult.CreateDataReader()
                        Return True
                    Else
                        _flagReaderReady = False
                        Return False
                    End If
                Case 1
                    If .ExecuteQuery(True, _QueryParamSql, _queryResult) Then
                        _flagReaderReady = True
                        _columnCount = _queryResult.Columns.Count
                        _recordCount = _queryResult.Rows.Count

                        _queryResultReader = _queryResult.CreateDataReader()

                        Return True
                    Else
                        _flagReaderReady = False
                        Return False
                    End If
                Case Else
                    Return False
            End Select
        End With
    End Function

End Class
