Imports helix.SQLEngine
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SQLEngineBuilder

    ''' <summary>
    ''' Opciones de tipos de cursores
    ''' </summary>
    Public Enum cursorType As Byte
        GLOBAL_CURSOR = 0
        LOCAL_CURSOR = 1
    End Enum

    ''' <summary>
    ''' Tipos de parametizaciones
    ''' </summary>
    Public Enum parametizationType As Byte
        SIMPLE = 0
        FORCED = 1
    End Enum

    ''' <summary>
    ''' Tipos de accesos
    ''' </summary>
    Public Enum accessType As Byte
        MULTI_USER = 0
        SINGLE_USER = 1
        RESTRICTED_USER = 2
    End Enum

    ''' <summary>
    ''' Tipos de recuperacion
    ''' </summary>
    Public Enum recoveryType As Byte
        FULL = 0
        BULK_LOGGED = 1
        SIMPLE = 2
    End Enum

    ''' <summary>
    ''' Tipos de verificacion de pagina
    ''' </summary>
    Public Enum pageVerifyType As Byte
        CHECKSUM = 0
        TORN_PAGE_DETECTION = 1
        NONE = 1
    End Enum

    ''' <summary>
    ''' Tipos de durabilidad retrasada
    ''' </summary>
    Public Enum delayedDurabilityType As Byte
        DISABLED = 0
        ALLOWED = 1
        FORCED = 2
    End Enum


    Public Structure dbFileGroup
        Dim name As String
        Dim files As Integer
        Dim isReadOnly As Boolean
        Dim isDefault As Boolean
    End Structure


    ''' <summary>
    ''' Nombre de la base de datos a crearse
    ''' </summary>
    ''' <value>Cadena con el nombre de la base a crearse</value>
    ''' <returns>El nombre de la base a crearse</returns>
    ''' <remarks>Si en el archivo de script no se encuentra el nombre de la base de datos se utiliza esta</remarks>
    Public Property DataBaseName As String = "helix"

    ''' <summary>
    ''' Nombre del servidor
    ''' </summary>
    ''' <value>Puede ser una direccion de IP, servidor/instancia o </value>
    ''' <returns>El nombre del servidor de base de datos</returns>
    ''' <remarks></remarks>
    Public Property ServerName As String = ".\SQLEXPRESS"

    ''' <summary>
    ''' Indica si debe crearse un usuario especifico para el manejo de la base de datos
    ''' </summary>
    ''' <value>Un booleano indicando si debe o no crearse un usuario nuevo</value>
    ''' <returns>Si debe crearse un usuario nuevo</returns>
    ''' <remarks></remarks>
    Public Property CreateDbUser As Boolean = False

    ''' <summary>
    ''' Modo de autenticacion a la base de datos
    ''' </summary>
    ''' <value>Booleano que indica el tipo de autenticacion: True = Windows, False = Mixta</value>
    ''' <returns>El tipo de autenticacion</returns>
    ''' <remarks>Si es true se usa autenticacion mixta, False autenticacion de Windows</remarks>
    Public Property RequireCredentials As Boolean = False

    ''' <summary>
    ''' El nombre de usuario en modo de autenticacion mixta
    ''' </summary>
    ''' <value>Cadena con el nombre del usuario de la base de datos</value>
    ''' <returns>El nombre de usuario</returns>
    ''' <remarks></remarks>
    Public Property Username As String = ""

    ''' <summary>
    ''' El password en modo de autenticacion mixta
    ''' </summary>
    ''' <value>Cadena con contraseña del usuario de la base de datos</value>
    ''' <returns>El password del usuario</returns>
    ''' <remarks></remarks>
    Public Property Password As String = ""

    ''' <summary>
    ''' Ubicacion en el sistema de archivos donde se encuentra el archivo con indicaciones de creacion de tablas
    ''' </summary>
    ''' <value>Cadena con el path completo y el nombre de archivo con script de creacion de tablas</value>
    ''' <returns>El path completo del archivo de script para la creacion de tablas</returns>
    ''' <remarks></remarks>
    Public Property ModelPath As String = ""

    ''' <summary>
    ''' Tipo de base de datos
    ''' </summary>
    ''' <value>El tipo de base datos</value>
    ''' <returns></returns>
    ''' <remarks>0 = Ms Access, 1 = SQL Server</remarks>
    Public Property DatabaseType As dataBaseType

    Public Structure SQLServerDBProperties
        Dim dbOwner As String
        Dim dbFilesGroup As List(Of dbFileGroup)
        Dim dbName As String
        Dim dbFullPath As String
        Dim dbInitialSizeKb As Integer
        Dim dbFileGrowth As String
        Dim logSizeKb As Integer
        Dim logFileGrowth As String
        Dim compatibilityLevel As Integer
        Dim ansiNullDefault As Boolean
        Dim ansiNulls As Boolean
        Dim ansiWarnings As Boolean
        Dim ansiPadding As Boolean
        Dim arithmeticAbort As Boolean
        Dim autoClose As Boolean
        Dim autoShrink As Boolean
        Dim autoCreateStatistics As Boolean
        Dim autoUpdateStatistics As Boolean
        Dim cursorCloseOnCommit As Boolean
        Dim cursorDefault As cursorType
        Dim concatenateNullYieldsNull As Boolean
        Dim numericRoundAbort As Boolean
        Dim quotedIdentifier As Boolean
        Dim recursiveTriggers As Boolean
        Dim autoUpdateStatisticsAsync As Boolean
        Dim dateCorrelationOptimization As Boolean
        Dim parameterization As parametizationType
        Dim readCommittedSnapshot As Boolean
        Dim readWrite As Boolean
        Dim recovery As recoveryType
        Dim restrictAccess As accessType
        Dim pageVerify As pageVerifyType
        Dim targetRecoveryTime As Integer
        Dim delayedDurability As delayedDurabilityType
        Dim isWindowsAuthenticated As Boolean
    End Structure


    Public SQLDbProperties As SQLServerDBProperties


    Public Sub New()
        MyBase.New()
        With SQLDbProperties
            .dbOwner = ""
            .dbName = "master"
            .dbFullPath = Application.StartupPath & "\"
            .dbInitialSizeKb = 5120
            .dbFileGrowth = "1024KB"
            .logSizeKb = 1024
            .logFileGrowth = "10%"
            .compatibilityLevel = 120
            .ansiNullDefault = False
            .ansiNulls = False
            .ansiWarnings = False
            .ansiPadding = False
            .arithmeticAbort = False
            .autoClose = False
            .autoShrink = False
            .autoCreateStatistics = True
            .autoUpdateStatistics = True
            .cursorCloseOnCommit = False
            .cursorDefault = cursorType.GLOBAL_CURSOR
            .concatenateNullYieldsNull = False
            .numericRoundAbort = False
            .quotedIdentifier = False
            .recursiveTriggers = False
            .autoUpdateStatisticsAsync = False
            .dateCorrelationOptimization = False
            .parameterization = parametizationType.SIMPLE
            .readCommittedSnapshot = False
            .readWrite = True
            .recovery = recoveryType.SIMPLE
            .restrictAccess = accessType.MULTI_USER
            .pageVerify = pageVerifyType.CHECKSUM
            .targetRecoveryTime = 0
            .delayedDurability = delayedDurabilityType.DISABLED
            .isWindowsAuthenticated = True
        End With
    End Sub


    ''' <summary>
    ''' Genera una cadena de conexion
    ''' </summary>
    ''' <returns>La cadena de conexion</returns>
    ''' <remarks></remarks>
    Public Function GenerateConnectionString() As String

        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS

            Case SQLEngine.dataBaseType.SQL_SERVER
                Dim tmpStr As String = ""
                tmpStr &= "Data Source=" & ServerName & ";"
                If _RequireCredentials = True Then
                    tmpStr &= "Integrated Security=False;"
                    tmpStr &= "uid=" & _Username & ";"
                    tmpStr &= "Password=" & _Password & ";"
                Else
                    tmpStr &= "Integrated Security=True;"
                End If
                Return tmpStr & "Connect Timeout=15;Encrypt=False;TrustServerCertificate=False"
        End Select
        Return ""
    End Function

    ''' <summary>
    ''' Prueba la conexion a la base de datos para verificar que todo esta correcto
    ''' </summary>
    ''' <returns>True si se pudo conectar con exito, False si no pudo</returns>
    ''' <remarks></remarks>
    Public Function TestConnection() As Boolean
        Dim tmpCore As New SQLCore

        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS

            Case SQLEngine.dataBaseType.SQL_SERVER
                With tmpCore
                    .dbType = DatabaseType.SQL_SERVER
                    .ConnectionString = GenerateConnectionString()
                    Return .TestConnection()
                End With
        End Select
        Return False
    End Function

    Public Function CreateNewDataBase() As Boolean
        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS
            Case SQLEngine.dataBaseType.SQL_SERVER
                Dim tmpStr As String

                Dim strAlterPrefix As String = "ALTER DATABASE [" & _DataBaseName & "] SET "

                tmpStr = "CREATE DATABASE [" & _DataBaseName & "] CONTAINMENT = NONE ON  PRIMARY " & _
                        "( NAME = N'" & _DataBaseName & "', FILENAME = N'" & SQLDbProperties.dbFullPath & _DataBaseName & ".mdf', "
                tmpStr += "SIZE = " & SQLDbProperties.dbInitialSizeKb & "KB , "
                tmpStr += "FILEGROWTH = " & SQLDbProperties.dbFileGrowth & " ) "
                tmpStr += "LOG ON ( NAME = N'" & _DataBaseName & "_log', FILENAME = N'" & SQLDbProperties.dbFullPath & _DataBaseName & "_log.ldf' , SIZE = " & _
                          SQLDbProperties.logSizeKb.ToString & "KB , FILEGROWTH = " & SQLDbProperties.logFileGrowth & ");"


                tmpStr += strAlterPrefix & "COMPATIBILITY_LEVEL = " & SQLDbProperties.compatibilityLevel & ";"


                tmpStr += strAlterPrefix & "ANSI_NULL_DEFAULT "
                If SQLDbProperties.ansiNullDefault = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "ANSI_NULLS "
                If SQLDbProperties.ansiNulls = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "ANSI_PADDING "
                If SQLDbProperties.ansiPadding = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                'tmpStr += strAlterPrefix & "ANSI_WARNINGS "
                'If SQLDbProperties.ansiWarnings = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "ARITHABORT "
                'If SQLDbProperties.arithmeticAbort = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "AUTO_CLOSE "
                'If SQLDbProperties.autoClose = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "AUTO_SHRINK "
                'If SQLDbProperties.autoClose = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "AUTO_CREATE_STATISTICS "
                'If SQLDbProperties.autoCreateStatistics = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "AUTO_UPDATE_STATISTICS "
                'If SQLDbProperties.autoUpdateStatistics = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "CURSOR_CLOSE_ON_COMMIT "
                'If SQLDbProperties.cursorCloseOnCommit = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "CURSOR_DEFAULT "
                'Select Case SQLDbProperties.cursorDefault
                '    Case cursorType.GLOBAL_CURSOR
                '        tmpStr += "GLOBAL;"
                '    Case cursorType.LOCAL_CURSOR
                '        tmpStr += "LOCAL;"
                'End Select

                'tmpStr += strAlterPrefix & "CONCAT_NULL_YIELDS_NULL "
                'If SQLDbProperties.concatenateNullYieldsNull = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "NUMERIC_ROUNDABORT "
                'If SQLDbProperties.numericRoundAbort = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "QUOTED_IDENTIFIER "
                'If SQLDbProperties.quotedIdentifier = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "RECURSIVE_TRIGGERS "
                'If SQLDbProperties.recursiveTriggers = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "AUTO_UPDATE_STATISTICS_ASYNC "
                'If SQLDbProperties.autoUpdateStatisticsAsync = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "DATE_CORRELATION_OPTIMIZATION "
                'If SQLDbProperties.dateCorrelationOptimization = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                'tmpStr += strAlterPrefix & "PARAMETERIZATION "
                'Select Case SQLDbProperties.parameterization
                '    Case parametizationType.SIMPLE
                '        tmpStr += "SIMPLE;"
                '    Case parametizationType.FORCED
                '        tmpStr += "FORCED;"
                'End Select


                'tmpStr += strAlterPrefix & "READ_COMMITTED_SNAPSHOT "
                'If SQLDbProperties.readCommittedSnapshot = True Then
                '    tmpStr += "ON;"
                'Else
                '    tmpStr += "OFF;"
                'End If


                ''tmpStr += strAlterPrefix & "READ_WRITE "
                ''If SQLDbProperties.readWrite = True Then
                ''tmpStr += "ON;"
                ''Else
                ''tmpStr += "OFF;"
                ''End If

                'tmpStr += strAlterPrefix & "RECOVERY "
                'Select Case SQLDbProperties.recovery
                '    Case recoveryType.FULL
                '        tmpStr += "FULL;"
                '    Case recoveryType.BULK_LOGGED
                '        tmpStr += "BULK_LOGGED;"
                '    Case recoveryType.SIMPLE
                '        tmpStr += "SIMPLE;"
                'End Select

                'tmpStr += strAlterPrefix & " "
                'Select Case SQLDbProperties.restrictAccess
                '    Case accessType.MULTI_USER
                '        tmpStr += "MULTI_USER;"
                '    Case accessType.SINGLE_USER
                '        tmpStr += "SINGLE;"
                '    Case accessType.RESTRICTED_USER
                '        tmpStr += "RESTRICTED_USER;"
                'End Select

                'tmpStr += strAlterPrefix & "PAGE_VERIFY "
                'Select Case SQLDbProperties.pageVerify
                '    Case pageVerifyType.CHECKSUM
                '        tmpStr += "CHECKSUM;"
                '    Case pageVerifyType.NONE
                '        tmpStr += "NONE;"
                '    Case pageVerifyType.TORN_PAGE_DETECTION
                '        tmpStr += "TORN_PAGE_DETECTION;"
                'End Select


                'tmpStr += strAlterPrefix & "TARGET_RECOVERY_TIME = " & SQLDbProperties.targetRecoveryTime.ToString & "SECONDS;"


                'tmpStr += strAlterPrefix & "DELAYED_DURABILITY = "
                'Select Case SQLDbProperties.delayedDurability
                '    Case delayedDurabilityType.DISABLED
                '        tmpStr += "DISABLED;"
                '    Case delayedDurabilityType.ALLOWED
                '        tmpStr += "ALLOWED;"
                '    Case delayedDurabilityType.FORCED
                '        tmpStr += "FORCED;"
                'End Select

                'If SQLDbProperties.isWindowsAuthenticated = True Then
                '    tmpStr += "USE [" & _DataBaseName & "];" & "IF NOT EXISTS (SELECT name FROM sys.filegroups WHERE is_default=1 AND name = N'PRIMARY') ALTER DATABASE [" & _
                '        _DataBaseName & "] MODIFY FILEGROUP [PRIMARY] DEFAULT"
                'End If

                Dim tmpCore As New SQLCore
                tmpCore.dbType = DatabaseType.SQL_SERVER
                Dim tmpdbname = _DataBaseName
                ' Guarda el nombre de la base de datos del usuario en una variable temporaria porque la cadena de conexion
                ' no se puede conectar a la base de datos que todavia no existe
                _DataBaseName = "master"
                tmpCore.ConnectionString = GenerateConnectionString()
                _DataBaseName = tmpdbname
                Return tmpCore.ExecuteNonQuery(tmpStr)
        End Select
        Return False
    End Function

End Class
