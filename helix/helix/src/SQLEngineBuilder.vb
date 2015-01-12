Public Class SQLEngineBuilder

    ''' <summary>
    ''' Nombre de la base de datos a crearse
    ''' </summary>
    ''' <value>Cadena con el nombre de la base a crearse</value>
    ''' <returns>El nombre de la base a crearse</returns>
    ''' <remarks>Si en el archivo de script no se encuentra el nombre de la base de datos se utiliza esta</remarks>
    Public Property DataBaseName As String = "helix"

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

End Class
