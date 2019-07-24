
Imports System
Imports System.Configuration
Imports System.Diagnostics
Imports System.Xml


' <summary>
' Implements IConfigurationSectionHandler.  Creates an AppStartConfig object which encapsulates necessary configuration settings
' in a strong-typed reliable class for consumption by AppStart core process.
' </summary>

Friend Class ConfigSectionHandler
    Implements IConfigurationSectionHandler
    #Region "Constructors"
    
    
    ' <summary>
    ' Default constructor.
    ' </summary>
    Public Sub New() 
    
    End Sub 'New
    
    #End Region
    
    #Region "IConfigurationSectionHandler members"
    
    
    
    ' <summary>
    ' This class is responsible for reading the configuration and creating an instance of the configuration class.
    ' 
    ' &lt;ClientApplicationInfo&gt;
    ' &lt;appFolderName&gt;C:\temp\pag\AppUpdaterClient&lt;/appFolderName&gt;
    ' &lt;appExeName&gt;SampleApplicationForBitsDownload.exe&lt;/appExeName&gt;
    ' &lt;/ClientApplicationInfo&gt;
    ' </summary>
    ' <param name="parent"></param>
    ' <param name="configContext"></param>
    ' <param name="section"></param>
    ' <returns></returns>
    Function Create(ByVal parent As Object, ByVal configContext As Object, ByVal section As XmlNode) As Object  Implements IConfigurationSectionHandler.Create
        Dim config As AppStartConfiguration = Nothing
        Dim rootNode As XmlNode = Nothing
        Dim subNode As XmlNode = Nothing
        Dim temp As String = ""
        
        '  grab an instance of AppStartConfiguration
        config = New AppStartConfiguration()
        
        '  get the primary node "ClientApplicationInfo" so we can access others easily
        rootNode = section.SelectSingleNode("ClientApplicationInfo")
        
        '  access each subnode, validating values and adding them to our instance of config object			
        Try
            subNode = rootNode.SelectSingleNode("appFolderName")
            If Not (subNode Is Nothing) Then
                temp = subNode.InnerText
                '  check if terminal slash, add if missing
                If Not temp.EndsWith("\") Then
                    temp += "\"
                End If
                config.FolderName = temp
            End If

            subNode = rootNode.SelectSingleNode("appExeName")
            If Not (subNode Is Nothing) Then
                config.ExecutableName = subNode.InnerText
            End If

            subNode = rootNode.SelectSingleNode("appID")
            If Not (subNode Is Nothing) Then
                temp = subNode.InnerText
                config.GetApplicationID = subNode.InnerText
            End If

            subNode = rootNode.SelectSingleNode("updateTime")
            If Not (subNode Is Nothing) Then
                config.UpdateTime = CType([Enum].Parse(GetType(UpdateTimeEnum), subNode.InnerText), UpdateTimeEnum)
            End If

            subNode = rootNode.SelectSingleNode("manifestUri")
            If Not (subNode Is Nothing) Then
                config.GetManifestUri = New Uri(subNode.InnerText)
            End If

        Catch e As Exception
            Trace.WriteLine("AppStart:[ConfigSectionHandler.Create]: Error during parsing of app.config file:" + Environment.NewLine + e.Message)
            Throw e
        End Try
        
        Return config
    
    End Function 'IConfigurationSectionHandler.Create 
    #End Region
End Class 'ConfigSectionHandler