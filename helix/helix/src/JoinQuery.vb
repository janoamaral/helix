Option Explicit On
Option Strict On

Public Class JoinQuery
    Private Structure ColumnOperationValue
        Dim column As String
        Dim operation As String
        Dim value As String
    End Structure

    Private Structure ColumnValue
        Dim column As String
        Dim value As String
    End Structure

    Dim _joinQuery As String = ""
    Dim _ComparationColumns As New List(Of ColumnOperationValue)
    Dim _logicalOperators As New List(Of String)
    Dim _tables As ColumnValue
    Dim _isFirstChain As Boolean = True
    Dim _columns As New List(Of String)


    Public Property IsFirstChain() As Boolean
        Get
            Return _isFirstChain
        End Get
        Set(value As Boolean)
            _isFirstChain = value
        End Set
    End Property


    Public Sub SetTables(ByVal table1 As String, ByVal table2 As String)
        _tables.column = table1
        _tables.value = table2
    End Sub

    Public Sub AddComparationColumn(ByVal column1 As String, ByVal comparator As String, ByVal column2 As String)
        Dim tmp As ColumnOperationValue
        tmp.column = column1
        tmp.operation = comparator
        tmp.value = column2
        _ComparationColumns.Add(tmp)
    End Sub

    ''' <summary>
    ''' Union entre 2 o mas JoinQueries, por ejemplo AND, OR, etc
    ''' </summary>
    ''' <param name="logicalOperator">Cadena con el operador logico entre 2 JoinQueries</param>
    ''' <remarks></remarks>
    Public Sub AddLogicalOperator(ByVal logicalOperator As String)
        _logicalOperators.Add(logicalOperator)
    End Sub

    Public Sub ClearAllParameters()
        _joinQuery = ""
        _ComparationColumns.Clear()
        _logicalOperators.Clear()
        _tables = Nothing
        _isFirstChain = True
        _columns.Clear()
    End Sub

    Public Function GetJoinQuery() As String
        Dim tmpQuery As String = ""

        If _isFirstChain = True Then
            tmpQuery = "SELECT # FROM @" & _tables.column & " INNER JOIN " & _tables.value & " ON ("
        Else
            tmpQuery = ") INNER JOIN " & _tables.column & " ON ("
        End If
        '|<-                              Query join simple                                 ->|
        '                                                                                     |<-                  Query join compuesta                   ->|
        'SELECT * FROM (tabla1 INNER JOIN tabla2 ON tabla1.columnaComun = tabla2.columnaComun) INNER JOIN tabla3 ON tabla1.columnaComun = tabla3.columnaComun
        Dim parametersString As String = ""
        If ((_ComparationColumns.Count = 1) And (_ComparationColumns.Count <> 0)) Then 'Genera un join simple (tabla1.columnaComun = tabla2.columnaComun)
            tmpQuery &= _ComparationColumns(0).column & " " & _ComparationColumns(0).operation & " " & _ComparationColumns(0).value
        Else
            'Genera un join compuesto (tabla1.columnaComun = tabla3.columnaComun AND tablaN.columnaComun = tablaM.columnaComun)
            Dim tmpCompColumns As ColumnOperationValue
            Dim j As Integer = 1
            For Each tmpCompColumns In _ComparationColumns
                If ((j Mod 2) = 0) And (_logicalOperators.Count <> 0) Then
                    parametersString &= " " & _logicalOperators(j - 1)
                End If
                parametersString &= " " & tmpCompColumns.column & tmpCompColumns.operation & tmpCompColumns.value
                j += 1
            Next
        End If
        Return tmpQuery & parametersString
    End Function
End Class
