
Imports System
Imports System.IO


' <summary>
' Provides encapsulated view of AppStart configuration information.
' </summary>

Friend Class AppStartConfiguration
    #Region "Constructors"
    
    
    ' <summary>
    ' Default constructor.
    ' </summary>
    Public Sub New() 
    
    End Sub 'New
    
    #End Region
    
    #Region "Private members"
    
    ' <summary>
    ' The name of the executable file for the application.
    ' </summary>
    Private applicationExecutableName As String = String.Empty
    
    ' <summary>
    ' The name of the folder where the application was installed.
    ' </summary>
    Private applicationFolderName As String = String.Empty
    
    ' <summary>
    ' Specifes when the application should be run.
    ' </summary>
    Private whenToRunApplicationUpdate As UpdateTimeEnum = UpdateTimeEnum.BeforeStart
    
    ' <summary>
    ' The id of the application.
    ' </summary>
    Private applicationId As String
    
    ' <summary>
    ' The URI of the manifest.
    ' </summary>
    Private manifestUri As Uri
    
    ' <summary>
    ' The full path for the executable of the application.
    ' </summary>
    Private applicationExecutablePath As String = String.Empty
    
    #End Region
    
    #Region "Public Properties"
    
    ' <summary>
    ' The id of the application.
    ' </summary>
    
    Public Property GetApplicationID() As String 
        Get
            Return applicationId
        End Get
        Set
            applicationId = Value
        End Set
    End Property 
    ' <summary>
    ' The name of the executable file for the application.
    ' </summary>
    
    Public Property ExecutableName() As String 
        Get
            Return applicationExecutableName
        End Get
        Set
            applicationExecutableName = value
        End Set
    End Property 
    ' <summary>
    ' The name of the folder where the application was installed.
    ' </summary>
    
    Public Property FolderName() As String 
        Get
            Return applicationFolderName
        End Get
        Set
            applicationFolderName = value
        End Set
    End Property 
    ' <summary>
    ' Specifes when the application should be run.
    ' </summary>
    
    Public Property UpdateTime() As UpdateTimeEnum 
        Get
            Return whenToRunApplicationUpdate
        End Get
        Set
            whenToRunApplicationUpdate = Value
        End Set
    End Property 
    ' <summary>
    ' The URI of the manifest.
    ' </summary>
    
    Public Property GetManifestUri() As Uri
        Get
            Return manifestUri
        End Get
        Set(ByVal Value As Uri)
            manifestUri = Value
        End Set
    End Property
    ' <summary>
    ' The full path for the executable of the application.
    ' </summary>

    Public ReadOnly Property ExecutablePath() As String
        Get
            If applicationExecutablePath = String.Empty Then
                If Path.IsPathRooted(Me.FolderName) Then
                    applicationExecutablePath = Path.Combine(Me.FolderName, Me.ExecutableName)
                Else
                    applicationExecutablePath = Path.Combine(Environment.CurrentDirectory, Me.FolderName)
                    applicationExecutablePath = Path.GetFullPath(Path.Combine(applicationExecutablePath, Me.ExecutableName))
                End If
            End If
            Return applicationExecutablePath
        End Get
    End Property

#End Region
End Class 'AppStartConfiguration

' <summary>
' Specifies the options for the execution of the Updater before of after the application.
' </summary>

Friend Enum UpdateTimeEnum
    ' <summary>
    ' Check for updates before the application is executed.
    ' </summary>
    BeforeStart
    
    ' <summary>
    ' Check for updates after the application is shutdown.
    ' </summary>
    AfterShutdown
End Enum 'UpdateTimeEnum