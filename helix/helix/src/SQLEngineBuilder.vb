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

    Public Enum dbUserRoles As Byte
        AccessAdmin = 0
        BackupOperator = 1
        DataReader = 2
        DataWriter = 3
        DdlAdmin = 4
        DenyDataReader = 5
        DenyDataWriter = 6
        Owner = 7
        SecurityAdmin = 8
    End Enum

    ''' <summary>
    ''' Lista de roles del usuario de 
    ''' </summary>
    ''' <remarks></remarks>
    Private _dbRoles As New List(Of dbUserRoles)

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
    ''' Nombre del usuario con privilegios minimos para conectarse a la base de datos 
    ''' </summary>
    ''' <value>Cadena con el nombre de usuario para conectarse a la base de datos</value>
    ''' <returns>El nombre del usuario normal</returns>
    ''' <remarks></remarks>
    Public Property DbUsername As String = "helix_user"

    ''' <summary>
    ''' Contraseña del usuario con privilegios minimos para conectarse a la base de datos 
    ''' </summary>
    ''' <value>Cadena con la contraseña para conectarse a la base de datos</value>
    ''' <returns>La contraseña del usuario normal</returns>
    ''' <remarks></remarks>
    Public Property DbPassword As String = "v1ttqpEE3sTd"

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

    ''' <summary>
    ''' Estructura con configuracion de base de datos
    ''' </summary>
    ''' <remarks>Normalmente con default es suficiente</remarks>
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

    Public LastError As New Ermac

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
    Public Function GenerateConnectionString(ByVal createDb As Boolean) As String

        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS

            Case SQLEngine.dataBaseType.SQL_SERVER
                Dim tmpStr As String = ""
                tmpStr &= "Data Source=" & ServerName & ";"

                If createDb Then tmpStr &= "Initial Catalog=master;" Else tmpStr &= "Initial Catalog=" & _DataBaseName & ";"

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

                    ' Se devuelve la conexion con creacion de base de datos porque se asegura que la base existe
                    .ConnectionString = GenerateConnectionString(True)
                    Return .TestConnection()
                End With
        End Select
        Return False
    End Function


    ''' <summary>
    ''' Crea una base de datos en el destino seleccionado
    ''' </summary>
    ''' <returns>True si se creo con exito, False si fallo</returns>
    ''' <remarks></remarks>
    Public Function CreateNewDataBase() As Boolean
        Dim tmpStr As String
        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS
                ' TODO: Crear base de datos en ms access
                Dim bAns As Boolean
                Dim cat As New ADOX.Catalog
                Try

                Catch ex As Exception

                End Try
            Case SQLEngine.dataBaseType.SQL_SERVER

                Dim strAlterPrefix As String = "ALTER DATABASE [" & _DataBaseName & "] SET "

                tmpStr = "CREATE DATABASE [" & _DataBaseName & "] CONTAINMENT = NONE ON  PRIMARY " &
                        "( NAME = N'" & _DataBaseName & "', FILENAME = N'" & SQLDbProperties.dbFullPath & _DataBaseName & ".mdf', "
                tmpStr += "SIZE = " & SQLDbProperties.dbInitialSizeKb & "KB , "
                tmpStr += "FILEGROWTH = " & SQLDbProperties.dbFileGrowth & " ) "
                tmpStr += "LOG ON ( NAME = N'" & _DataBaseName & "_log', FILENAME = N'" & SQLDbProperties.dbFullPath & _DataBaseName & "_log.ldf' , SIZE = " &
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


                tmpStr += strAlterPrefix & "ANSI_WARNINGS "
                If SQLDbProperties.ansiWarnings = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "ARITHABORT "
                If SQLDbProperties.arithmeticAbort = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "AUTO_CLOSE "
                If SQLDbProperties.autoClose = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "AUTO_SHRINK "
                If SQLDbProperties.autoClose = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "AUTO_CREATE_STATISTICS "
                If SQLDbProperties.autoCreateStatistics = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "AUTO_UPDATE_STATISTICS "
                If SQLDbProperties.autoUpdateStatistics = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "CURSOR_CLOSE_ON_COMMIT "
                If SQLDbProperties.cursorCloseOnCommit = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "CURSOR_DEFAULT "
                Select Case SQLDbProperties.cursorDefault
                    Case cursorType.GLOBAL_CURSOR
                        tmpStr += "GLOBAL;"
                    Case cursorType.LOCAL_CURSOR
                        tmpStr += "LOCAL;"
                End Select

                tmpStr += strAlterPrefix & "CONCAT_NULL_YIELDS_NULL "
                If SQLDbProperties.concatenateNullYieldsNull = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "NUMERIC_ROUNDABORT "
                If SQLDbProperties.numericRoundAbort = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "QUOTED_IDENTIFIER "
                If SQLDbProperties.quotedIdentifier = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "RECURSIVE_TRIGGERS "
                If SQLDbProperties.recursiveTriggers = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "AUTO_UPDATE_STATISTICS_ASYNC "
                If SQLDbProperties.autoUpdateStatisticsAsync = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "DATE_CORRELATION_OPTIMIZATION "
                If SQLDbProperties.dateCorrelationOptimization = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "PARAMETERIZATION "
                Select Case SQLDbProperties.parameterization
                    Case parametizationType.SIMPLE
                        tmpStr += "SIMPLE;"
                    Case parametizationType.FORCED
                        tmpStr += "FORCED;"
                End Select


                tmpStr += strAlterPrefix & "READ_COMMITTED_SNAPSHOT "
                If SQLDbProperties.readCommittedSnapshot = True Then
                    tmpStr += "ON;"
                Else
                    tmpStr += "OFF;"
                End If


                tmpStr += strAlterPrefix & "RECOVERY "
                Select Case SQLDbProperties.recovery
                    Case recoveryType.FULL
                        tmpStr += "FULL;"
                    Case recoveryType.BULK_LOGGED
                        tmpStr += "BULK_LOGGED;"
                    Case recoveryType.SIMPLE
                        tmpStr += "SIMPLE;"
                End Select

                tmpStr += strAlterPrefix & " "
                Select Case SQLDbProperties.restrictAccess
                    Case accessType.MULTI_USER
                        tmpStr += "MULTI_USER;"
                    Case accessType.SINGLE_USER
                        tmpStr += "SINGLE;"
                    Case accessType.RESTRICTED_USER
                        tmpStr += "RESTRICTED_USER;"
                End Select

                tmpStr += strAlterPrefix & "PAGE_VERIFY "
                Select Case SQLDbProperties.pageVerify
                    Case pageVerifyType.CHECKSUM
                        tmpStr += "CHECKSUM;"
                    Case pageVerifyType.NONE
                        tmpStr += "NONE;"
                    Case pageVerifyType.TORN_PAGE_DETECTION
                        tmpStr += "TORN_PAGE_DETECTION;"
                End Select


                tmpStr += strAlterPrefix & "TARGET_RECOVERY_TIME = " & SQLDbProperties.targetRecoveryTime.ToString & "SECONDS;"


                tmpStr += strAlterPrefix & "DELAYED_DURABILITY = "
                Select Case SQLDbProperties.delayedDurability
                    Case delayedDurabilityType.DISABLED
                        tmpStr += "DISABLED;"
                    Case delayedDurabilityType.ALLOWED
                        tmpStr += "ALLOWED;"
                    Case delayedDurabilityType.FORCED
                        tmpStr += "FORCED;"
                End Select

                'If SQLDbProperties.isWindowsAuthenticated = True Then
                '    tmpStr += "USE [" & _DataBaseName & "];" & "IF NOT EXISTS (SELECT name FROM sys.filegroups WHERE is_default=1 AND name = N'PRIMARY') ALTER DATABASE [" & _
                '        _DataBaseName & "] MODIFY FILEGROUP [PRIMARY] DEFAULT"
                'End If

                Dim tmpCore As New SQLCore
                tmpCore.dbType = dataBaseType.SQL_SERVER
                tmpCore.ConnectionString = GenerateConnectionString(True)
                Return tmpCore.ExecuteNonQuery(tmpStr)
        End Select
        Return False
    End Function

    Public Sub AddUserRole(ByVal userRole As dbUserRoles)
        _dbRoles.Add(userRole)
    End Sub

    ''' <summary>
    ''' Ejecuta el script de creacion de usuario de db
    ''' </summary>
    ''' <returns>TRUE si se creo correctamente, FALSE si no</returns>
    ''' <remarks></remarks>
    Public Function CreateLoginUser() As Boolean
        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS
                ' TODO: Crear tablas en MS Access
            Case SQLEngine.dataBaseType.SQL_SERVER
                Dim tmpCore As New SQLCore
                tmpCore.dbType = DatabaseType.SQL_SERVER
                tmpCore.ConnectionString = GenerateConnectionString(False)
                Dim comm As String = GenerateUserScript()
                ' Crear el usuario
                If tmpCore.ExecuteNonQuery(comm) Then
                    ' Crear los permisos
                    comm = GenerateUserPermissionScript()
                    ' Si todo funciono deberia ejecutar correctamente
                    Return tmpCore.ExecuteNonQuery(comm)
                End If
        End Select
        Return False
    End Function


    Private Function GenerateUserScript() As String
        Dim tmpStr As String = ""

        If CreateDbUser = True Then
            tmpStr = "USE [master];"
            tmpStr &= "CREATE LOGIN [" & DbUsername & "] WITH PASSWORD=N'" & DbPassword & "', DEFAULT_DATABASE=[" & DataBaseName & "], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF;"
            tmpStr &= "USE [" & DataBaseName & "];"
            tmpStr &= "CREATE USER [" & DbUsername & "] FOR LOGIN [" & DbUsername & "];"
        End If

        Return tmpStr
    End Function

    ''' <summary>
    ''' Genera el script para crear el usuario normal del sistema
    ''' </summary>
    ''' <returns>Cadena con el script de creacion de usuario</returns>
    ''' <remarks></remarks>
    Private Function GenerateUserPermissionScript() As String
        Dim tmpStr As String = ""

        If CreateDbUser = True Then

            For Each role In _dbRoles
                tmpStr &= "USE [" & DataBaseName & "];"

                Dim tmpStrRole As String = ""
                Select Case role
                    Case dbUserRoles.AccessAdmin
                        tmpStrRole = "db_accessadmin"
                    Case dbUserRoles.BackupOperator
                        tmpStrRole = "db_backupoperator"
                    Case dbUserRoles.DataReader
                        tmpStrRole = "db_datareader"
                    Case dbUserRoles.DataWriter
                        tmpStrRole = "db_datawriter"
                    Case dbUserRoles.DdlAdmin
                        tmpStrRole = "db_ddladmin"
                    Case dbUserRoles.DenyDataReader
                        tmpStrRole = "db_denydatareader"
                    Case dbUserRoles.DenyDataWriter
                        tmpStrRole = "db_denydatawriter"
                    Case dbUserRoles.Owner
                        tmpStrRole = "db_owner"
                    Case dbUserRoles.SecurityAdmin
                        tmpStrRole = "db_securityadmin"
                End Select

                tmpStr &= "ALTER ROLE [" & tmpStrRole & "] ADD MEMBER [" & DbUsername & "];"
            Next
        End If

        'tmpStr &= "ALTER ROLE [public] ADD MEMBER [" & DbUsername & "];"
        Return tmpStr

    End Function


    ''' <summary>
    ''' Ejecuta el script de creacion de tabla en la base de datos 
    ''' </summary>
    ''' <returns>True si se ejecuto el script con exito, False si fallo</returns>
    ''' <remarks></remarks>
    Public Function CreateTable() As Boolean
        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS
                ' TODO: Crear tablas en MS Access
            Case SQLEngine.dataBaseType.SQL_SERVER
                Dim tmpCore As New SQLCore
                tmpCore.dbType = DatabaseType.SQL_SERVER
                tmpCore.ConnectionString = GenerateConnectionString(False)
                Dim comm As String = generateTableScript()
                Return tmpCore.ExecuteNonQuery(comm)
        End Select
        Return False
    End Function

    ''' <summary>
    ''' Genera el comando para crear las tablas
    ''' </summary>
    ''' <returns>Cadena con el comando para crear la/s tablas</returns>
    ''' <remarks></remarks>
    Private Function GenerateTableScript() As String

        Dim scriptLine As String = ""
        Dim tmpScript As String = ""
        tmpScript = "SET QUOTED_IDENTIFIER ON;SET ARITHABORT ON;SET NUMERIC_ROUNDABORT OFF;SET CONCAT_NULL_YIELDS_NULL ON;SET ANSI_NULLS ON;SET ANSI_PADDING ON;SET ANSI_WARNINGS ON;"

        If _ModelPath.Length = 0 Or My.Computer.FileSystem.FileExists(_ModelPath) = False Then
            Return False
        End If

        Dim lineReader As New System.IO.StreamReader(_ModelPath)
        Dim splitLine As String()

        ' Comienzo de lectura del archivo modelo
        Do While lineReader.Peek() <> -1
            scriptLine = lineReader.ReadLine()

            If scriptLine.StartsWith("DATABASE_NAME") Then
                splitLine = scriptLine.Split("=")
                If splitLine.GetLength(0) >= 2 Then
                    _DataBaseName = splitLine(1).Trim(" ")
                Else
                    Return False
                End If
            End If

            If scriptLine.StartsWith("DATABASE_TYPE") Then
                splitLine = scriptLine.Split("=")
                If splitLine.GetLength(0) >= 2 Then
                    Select Case splitLine(1).Trim(" ").ToUpper
                        Case "MS_ACCESS"
                            _DatabaseType = SQLEngine.dataBaseType.MS_ACCESS
                        Case "SQL_SERVER"
                            _DatabaseType = SQLEngine.dataBaseType.SQL_SERVER
                    End Select
                Else
                    Return False
                End If
            End If

            If scriptLine.StartsWith("TABLE") Then
                Dim tableName As String = ""
                splitLine = scriptLine.Split(" ")
                If splitLine.GetLength(0) >= 2 Then
                    tableName = splitLine(1).Trim(" ")
                    tmpScript &= "CREATE TABLE dbo." & tableName & "("
                Else
                    Return False
                End If

                Dim pkField As String = ""
                Do While (lineReader.Peek() <> -1) And (scriptLine.ToUpper.Trim(" ") <> "END TABLE")
                    scriptLine = lineReader.ReadLine()
                    Dim processLine As String = ParseTableField(scriptLine, splitLine(1).Trim(" "))

                    If processLine.StartsWith("ISPK_") Then
                        processLine = processLine.Replace("ISPK_", "")
                        pkField = processLine.Split(" ")(0).Trim(" ")
                    End If
                    tmpScript &= processLine
                Loop

                tmpScript = tmpScript.Trim(",") & ") ON [PRIMARY];"
                tmpScript &= "ALTER TABLE dbo." & tableName & "  ADD CONSTRAINT PK_" & tableName & " PRIMARY KEY CLUSTERED (" & _
                             pkField & ") WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];"
                tmpScript &= "ALTER TABLE dbo." & tableName & " SET (LOCK_ESCALATION = TABLE);"

            End If
        Loop
        Debug.Print(tmpScript)
        Return tmpScript
    End Function


    Private Function ParseTableField(ByVal line As String, Optional tableName As String = "") As String

        ' Estructura del archivo script
        ' DATABASE_NAME = helix
        ' DATABASE_TYPE = SQL_SERVER

        ' TABLE table1
        ' pk_field = pk, auto:1/1
        ' string_std = string
        ' string_field = string, max
        ' string_acotada = string, 32
        ' date_field = date
        ' time_field = time
        ' timestamp_field = timestamp
        ' money_field = money
        ' remote_ref_field = pkfield
        ' long_text_field = Text
        ' null_allowed_field = anytype NULL
        ' END TABLE

        Dim fieldName As String = ""
        Dim tmpString As String = ""
        Dim splitLine As String()
        Dim splitOptions As String()


        Select Case _DatabaseType
            Case SQLEngine.dataBaseType.MS_ACCESS
                ' TODO: Generar codigo para ms access
            Case SQLEngine.dataBaseType.SQL_SERVER
                If line.Length > 0 And Not line.Trim(" ").StartsWith("#") Then

                    ' Primer split
                    ' (pk_field) = (pk, auto:1/1)
                    splitLine = line.Split("=")
                    If splitLine.Length > 1 Then
                        tmpString = tableName & "_" & splitLine(0).Trim(" ") & " "

                        ' Segundo split
                        ' splitOptions 0  , 1
                        '             (pk), (auto:1/1)
                        splitOptions = splitLine(1).Split(",")
                        Dim typeField As String = splitOptions(0).ToLower.Trim(" ")
                        Select Case typeField
                            Case "pk"
                                ' _id bigint NOT NULL IDENTITY (1, 1)
                                tmpString = tmpString.Insert(0, "ISPK_")    ' Envio flag is pk para tratarlo como clave primaria
                                tmpString &= "bigint"

                                If splitOptions.Length > 1 Then
                                    ' Tercer split
                                    ' (auto):(1/1)
                                    splitOptions = splitOptions(1).Split(":")
                                    If (splitOptions.Length > 1) And (splitOptions(0).ToLower.Trim(" ") = "auto") Then
                                        tmpString &= " IDENTITY (" & splitOptions(1).Split("/")(0) & ", " & splitOptions(1).Split("/")(1) & ")"
                                    End If
                                End If

                            Case "string"
                                tmpString &= "nvarchar("
                                If splitOptions.Length > 1 Then
                                    If splitOptions(1).Trim(" ").StartsWith("lenght") Then
                                        tmpString &= splitOptions(1).Split(":")(1).ToUpper & ")"
                                    Else
                                        If splitOptions(1).Trim(" ").ToLower.StartsWith("max") Then
                                            tmpString &= "max)"
                                        Else
                                            tmpString &= "150)"
                                        End If
                                    End If
                                End If
                            Case "bin"
                                tmpString &= "varbinary("
                                If splitOptions.Length > 1 Then
                                    If splitOptions(1).Trim(" ").StartsWith("lenght") Then
                                        tmpString &= splitOptions(1).Split(":")(1).ToUpper & ")"
                                    Else
                                        If splitOptions(1).Trim(" ").ToLower.StartsWith("max") Then
                                            tmpString &= "max)"
                                        Else
                                            tmpString &= "150)"
                                        End If
                                    End If
                                End If

                            Case "text"
                                tmpString &= "text"

                            Case "date"
                                tmpString &= "date"

                            Case "time"
                                tmpString &= "time(7)"

                            Case "bin"
                                tmpString &= "varbinary(max)"

                            Case "timestamp"
                                tmpString &= "datetime"

                            Case "money"
                                tmpString &= "decimal(25, 13)"

                            Case "int"
                                tmpString &= "int"

                            Case "long"
                                tmpString &= "bigint"

                            Case "longint"
                                tmpString &= "bigint"

                            Case "pkfield"
                                tmpString &= "bigint"

                            Case "bool"
                                tmpString &= "bit"

                            Case "percent"
                                tmpString &= "decimal(5,2)"

                            Case "img"
                                tmpString &= "image"

                            Case "byte"
                                tmpString &= "char(1)"

                            Case Else
                                tmpString &= ""

                        End Select

                        If line.Contains("!NULL") Then
                            tmpString &= " NOT NULL,"
                        Else
                            tmpString &= " NULL,"
                        End If

                        Return tmpString
                    Else
                        Return ""
                    End If
                End If
        End Select
        Return ""
    End Function
End Class
