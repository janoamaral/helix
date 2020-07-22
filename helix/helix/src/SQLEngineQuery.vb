
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

Public Class SQLEngineQuery
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

    Public Const OPERATOR_IGUAL As String = " = ? "
    Public Const OPERATOR_DISTINTO As String = " <> ? "
    Public Const OPERATOR_MENOR As String = " < ?"
    Public Const OPERATOR_MENORIGUAl As String = " <= ? "
    Public Const OPERATOR_MAYOR As String = " > ? "
    Public Const OPERATOR_MAYORIGUAl As String = " >= ? "
    Public Const OPERATOR_LIKE As String = " LIKE ? "
    Public Const OPERATOR_BETWEEN As String = " BETWEEN ? AND ? "

    Public Shared Function OperatorToString(ByVal op As OperatorCriteria) As String
        Select Case op
            Case OperatorCriteria.Igual
                Return OPERATOR_IGUAL
            Case OperatorCriteria.Distinto
                Return OPERATOR_DISTINTO
            Case OperatorCriteria.Menor
                Return OPERATOR_MENOR
            Case OperatorCriteria.MenorIgual
                Return OPERATOR_MENORIGUAl
            Case OperatorCriteria.Mayor
                Return OPERATOR_MAYOR
            Case OperatorCriteria.MayorIgual
                Return OPERATOR_MAYORIGUAl
            Case OperatorCriteria.LikeString
                Return OPERATOR_LIKE
            Case OPERATOR_BETWEEN
                Return OPERATOR_BETWEEN
            Case Else
                Return ""
        End Select
    End Function

    Private Structure CountSum
        Dim sqlFunction As String
        Dim rowName As String
        Dim asRowName As String
    End Structure

    Public Function OperatorString(ByVal operador As OperatorCriteria) As String
        Select Case operador
            Case OperatorCriteria.Igual
                Return OPERATOR_IGUAL
            Case OperatorCriteria.Distinto
                Return Operator_Distinto
            Case OperatorCriteria.Menor
                Return Operator_Menor
            Case OperatorCriteria.MenorIgual
                Return Operator_MenorIgual
            Case OperatorCriteria.Mayor
                Return Operator_Mayor
            Case OperatorCriteria.MayorIgual
                Return Operator_MayorIgual
            Case OperatorCriteria.LikeString
                Return Operator_Like
            Case OperatorCriteria.Between
                Return Operator_Between
            Case Else
                Return ""
        End Select
    End Function

    Private _lstFunctions As New List(Of CountSum)

    ''' <summary>
    ''' Ruta completa y nombre de archivo donde se van a guardar los logs
    ''' </summary>
    ''' <value>Cadena con la ruta completa y el nombre de archivo del log</value>
    ''' <returns>La ruta y el nombre del archivo log</returns>
    ''' <remarks></remarks>
    Public Property LogFileFullName As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & "syslog.log"

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

    Private _columnNames As New List(Of String)

    Private Structure ExcludeField
        Dim column As String
        Dim eOperator As Byte
        Dim value As Object
    End Structure

    Private _excludeFields As New List(Of ExcludeField)


    Public ReadOnly Property ColumnNames As List(Of String)
        Get
            Return _columnNames
        End Get
    End Property


    ' TODO: Cambiarle el nombre de excluir 
    Public Sub _AND(ByVal fieldName As String, ByVal eOperator As OperatorCriteria, ByVal value As Object)
        Dim tmpXclude As ExcludeField
        tmpXclude.column = fieldName
        tmpXclude.eOperator = eOperator
        tmpXclude.value = value

        _excludeFields.Add(tmpXclude)
    End Sub




    ''' <summary>
    ''' Modo de ordenacion del resultado el query
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum sortOrder
        ascending = 0
        descending = 1
    End Enum

    Public Property RowLimit As Integer = 0

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
        Select Case DbType
            Case 0
                _joinQuery = "@(" & table1 & " INNER JOIN " & table2 & " ON (" & commonColumnTable1 & " = " & commonColumnTable2 & "))"
            Case 1
                _joinQuery = "@(" & table1 & " INNER JOIN " & table2 & " ON " & commonColumnTable1 & " = " & commonColumnTable2 & ")"
        End Select

    End Sub

    ''' <summary>
    ''' Agrega una columna con el metodo de ordenacion
    ''' </summary>
    ''' <param name="column">Nombre de la columna a ordenar</param>
    ''' <param name="sortingOrder">Metodo de ordenacion</param>
    ''' <remarks></remarks>
    Public Sub AddOrderColumn(ByVal column As String, ByVal sortingOrder As SortOrder)
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
            Case 2
                Dim mySqlparam As New MySqlParameter
                mySqlparam.Value = param
                mySqlparam.ParameterName = "@p" & _QueryParamMySql.Count
                _QueryParamMySql.Add(mySqlparam)
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
    '''  Devuelve el valor de la columna en el registro actual
    ''' </summary>
    ''' <param name="columnName">El nombre de la columna</param>
    ''' <returns>El valor de la columna en el registro actual. Si falla devuelve FALSE</returns>
    ''' <remarks></remarks>
    Public Function GetQueryData(ByVal columnName As String) As Object
        If _flagReaderReady = True Then
            If IsDBNull(_queryResultReader.Item(columnName)) Then
                Return ""
            Else
                Return _queryResultReader.Item(columnName)
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Devuelve el DataTableReader de la consulta
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ResultReader As DataTableReader
        Get
            Return _queryResultReader
        End Get
    End Property

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
                If IsNothing(valueEnd) Then
                    AddWHEREparam(value)
                Else
                    AddWHEREparam(valueEnd)
                End If

                Exit Sub
        End Select
        AddWHEREparam(value)
    End Sub


    ''' <summary>
    ''' Genera una consulta SQL "SELECT" segun los parametros de la clase. Segun se escoja puede devolver una cadena para ser procesada por el SQLCore
    ''' o una cadena para depuracion
    ''' </summary>
    ''' <param name="toProcess">TRUE si la cadena va a ser procesada por SQLCore, FALSE si se usa como depuracion</param>
    ''' <returns>Una consulta SQL "SELECT"</returns>
    Protected Overrides Function GenerateQuery(toProcess As Boolean) As String
        Dim tmpQuery As String = "SELECT "
        Dim tmpStr As String

        If RowLimit <> 0 Then
            tmpQuery &= "TOP " & RowLimit & " "
        End If

        If _selectColumn.Count <> 0 Then                                    ' Detecta que se eligen las columnas
            For Each tmpStr In _selectColumn                                ' Agrega las columnas que se van a extraer los registros
                tmpQuery &= tmpStr & ", "
            Next
            tmpQuery = tmpQuery.Remove(tmpQuery.Length - 2, 2)              ' Quitar la coma y el espacio sobrante
        Else
            tmpQuery &= "*"                                                 ' Seleccionar todas las columnas
        End If

        If _lstFunctions.Count > 0 Then
            tmpQuery &= ", "
            For Each a As CountSum In _lstFunctions
                tmpQuery &= a.sqlFunction & "(" & a.rowName & ") AS " & a.asRowName & ", "
            Next
            tmpQuery = tmpQuery.Remove(tmpQuery.Length - 2, 2)              ' Quitar la coma y el espacio sobrante
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
            Select Case _dbType
                Case 0
                    Dim obj As Object
                    For Each obj In _QueryParamOle                                             ' Por cada parametro
                        tmpQuery = tmpQuery.Insert(tmpQuery.IndexOf("?"), obj.value.ToString)     ' Se reemplaza en la consulta final
                        tmpQuery = tmpQuery.Remove(tmpQuery.IndexOf("?"), 1)                ' Dando la consulta lista para depuracion
                    Next
                Case 1
                    Dim obj As Object
                    For Each obj In _QueryParamSql                                             ' Por cada parametro
                        tmpQuery = tmpQuery.Insert(tmpQuery.IndexOf("?"), obj.value.ToString)     ' Se reemplaza en la consulta final
                        tmpQuery = tmpQuery.Remove(tmpQuery.IndexOf("?"), 1)                ' Dando la consulta lista para depuracion
                    Next
                Case 2
                    Dim obj As Object
                    For Each obj In _QueryParamMySql                                             ' Por cada parametro
                        tmpQuery = tmpQuery.Insert(tmpQuery.IndexOf("?"), obj.value.ToString)     ' Se reemplaza en la consulta final
                        tmpQuery = tmpQuery.Remove(tmpQuery.IndexOf("?"), 1)                ' Dando la consulta lista para depuracion
                    Next
            End Select

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
        _lstFunctions.Clear()
        _columnNames.Clear()
        _excludeFields.Clear()
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
        _QueryParamMySql.Clear()
        _orderColumn.Clear()
        _columnCount = 0
        _recordCount = 0
        RowLimit = 0
    End Sub


    ''' <summary>
    ''' Ejecuta la consulta contra la base de datos
    ''' </summary>
    ''' <returns>El resultado de la consulta. TRUE si la consulta se realizo con exito, FALSE si fallo</returns>
    Public Function Query(Optional ByVal useCustomDataReader As Boolean = False, Optional ByRef dt As DataTable = Nothing) As Boolean
        Dim core As New SQLCore

        With core
            .LastError.LogFilePath = _LogFileFullName
            .ConnectionString = _connectionString
            .dbType = _dbType
            .QueryString = GenerateQuery(True)

            Select Case _dbType
                Case 0
                    If .ExecuteQuery(True, _QueryParamOle, _queryResult) Then
                        _flagReaderReady = True
                        _columnCount = _queryResult.Columns.Count
                        _recordCount = _queryResult.Rows.Count
                        _queryResultReader = _queryResult.CreateDataReader()

                        If useCustomDataReader Then
                            dt = _queryResult.Copy
                        End If


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

                        Dim i As Integer = 0
                        While i < _queryResult.Columns.Count
                            _columnNames.Add(_queryResult.Columns.Item(i).ColumnName)
                            i += 1
                        End While

                        _queryResultReader = _queryResult.CreateDataReader()

                        If useCustomDataReader Then
                            dt = _queryResult.Copy()
                        End If

                        Return True
                    Else
                        _flagReaderReady = False
                        Return False

                    End If
                Case 2
                    If .ExecuteQuery(True, _QueryParamMySql, _queryResult) Then
                        _flagReaderReady = True
                        _columnCount = _queryResult.Columns.Count
                        _recordCount = _queryResult.Rows.Count

                        Dim i As Integer = 0
                        While i < _queryResult.Columns.Count
                            _columnNames.Add(_queryResult.Columns.Item(i).ColumnName)
                            i += 1
                        End While

                        _queryResultReader = _queryResult.CreateDataReader()

                        If useCustomDataReader Then
                            dt = _queryResult.Copy()
                        End If

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

    Public Function GetFullQuery() As String
        Return GenerateQuery(False)
    End Function


End Class
