
Imports System
Imports System.Configuration
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.ApplicationBlocks.Updater
Imports System.Runtime.InteropServices
Imports System.Net


'<summary>
'This is the main class for AppStart.exe.
'</summary>

NotInheritable Public Class AppStart
    #Region "Private members"
    
    ' <summary>
    ' The configuration for the AppStart process.
    ' </summary>
    Private Shared config As AppStartConfiguration = Nothing
    
    ' <summary>
    ' The process of the running application.
    ' </summary>
    Private Shared applicationProcess As Process
    
    ' <summary>
    ' The mutex name for the AppStart.
    ' </summary>
    Private Shared appStartMutexGuid As New Guid(New Byte() {&H5F, &H4D, &H69, &H63, &H68, &H61, &H65, &H6C, &H20, &H53, &H74, &H75, &H61, &H72, &H74, &H5F})

    Private Shared ConnectionStateString As String


    #End Region
    Private Declare Function InternetGetConnectedState Lib _
             "wininet.dll" (ByRef lpSFlags As Int32, _
             ByVal dwReserved As Int32) As Boolean


    #Region "Constructors"
    
    
    ' <summary>
    ' Static constructor.
    ' </summary>
    Shared Sub New() 
        '  Grab our config instance which we use to read app.config params, figure out
        '  WHERE our target app is and WHAT version
        config = CType(ConfigurationSettings.GetConfig("appStart"), AppStartConfiguration)
    
    End Sub 'New
    
    
    ' <summary>
    ' No reason to ever "construct" this from outside, meant purely as invisible shim.
    ' </summary>
    Private Sub New() 
    
    End Sub 'New 
    #End Region

    Public Enum InetConnState
        modem = &H1
        lan = &H2
        proxy = &H4
        ras = &H10
        offline = &H20
        configured = &H40
    End Enum

#Region "Static members"


    ' <summary>
    ' Main entry point to process.  Checks to see if another instance is running already, disallows if so.
    ' </summary>
    ' <param name="args">Arguments are ignored.</param>
    <STAThread()> _
    Shared Sub Main(ByVal args() As String)
        'Check to see if AppStart is already running FOR the particular versioned folder of the target application
        Dim isOwned As Boolean = False
        Dim appStartMutex As New Mutex(True, config.ExecutableName + appStartMutexGuid.ToString(), isOwned)

        If Not isOwned Then
            MessageBox.Show(String.Format(CultureInfo.CurrentCulture, "There is already a copy of the application '{0}' running.  Please close that application before starting a new one.", config.ExecutableName))

            Environment.Exit(1)
        End If


        StartAppProcess()

    End Sub 'Main


    ' <summary>
    ' Main process code.
    ' </summary>
    Private Shared Sub StartAppProcess()



        Dim processStarted As Boolean = False

        'CheckInetConnection()
        Dim objPing As New cPinger
        objPing.PingHostName("OPCGW")

        ''If ((ConnectionStateString <> "Offline.") And (ConnectionStateString <> "Not Connected.")) Then
        If Not ((objPing.Status = -1) Or (objPing.Status = 11010)) Then
            If config.UpdateTime = UpdateTimeEnum.BeforeStart Then
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
                AppStart.UpdateApplication()
                System.Windows.Forms.Cursor.Current = Cursors.Default

            End If
        Else
            MessageBox.Show(Nothing, "No internet connection detected or OPCGW is not available at this time.  You must have an internet connection to check for updates.", "Muster Updater", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        objPing = Nothing

        'Start the application
        Try
            Dim p As New ProcessStartInfo(config.ExecutablePath)
            p.WorkingDirectory = Path.GetDirectoryName(config.ExecutablePath)
            applicationProcess = Process.Start(p)
            processStarted = True
            Debug.WriteLine("APPLICATION STARTER:  Started app:  " + config.ExecutablePath)
        Catch e As Exception
            Debug.WriteLine("APPLICATION STARTER:  Failed to start process at:  " + config.ExecutablePath)
            HandleTerminalError(e)
        End Try

        If processStarted AndAlso config.UpdateTime = UpdateTimeEnum.AfterShutdown Then
            applicationProcess.WaitForExit()
            AppStart.UpdateApplication()
        End If

    End Sub 'StartAppProcess


    Private Shared Sub UpdateApplication()
        Dim updaterManager As ApplicationUpdaterManager = ApplicationUpdaterManager.GetUpdater(config.GetApplicationID)
        Dim manifests As Manifest() = updaterManager.CheckForUpdates(config.GetManifestUri)

        If manifests.Length > 0 Then
            MessageBox.Show(Nothing, "Updates available. Please continue to apply.", "Muster Updater", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim manifest As Manifest
            For Each manifest In manifests
                manifest.Application.Location = Path.GetDirectoryName(config.ExecutablePath)
                manifest.Apply = True
            Next manifest
            updaterManager.Download(manifests, TimeSpan.MaxValue)
            updaterManager.Activate(manifests)
        Else
            MessageBox.Show(Nothing, "No updates available", "Muster Updater", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub 'UpdateApplication


    ' <summary>
    ' Prints out the console exception &amp; shuts down the application.
    ' </summary>
    Private Shared Sub HandleTerminalError(ByVal e As Exception)
        Debug.WriteLine("APPLICATION STARTER: Terminal error encountered.")
        Debug.WriteLine("APPLICATION STARTER: The following exception was encoutered:")
        Debug.WriteLine(e.ToString())
        Debug.WriteLine("APPLICATION STARTER: Shutting down")

        MessageBox.Show(String.Format(CultureInfo.CurrentCulture, "There was an error when trying to start the target application: {0}", e.Message))
        Environment.Exit(0)

    End Sub 'HandleTerminalError


    Private Shared Function CheckInetConnection() As Boolean

        Dim lngFlags As Long

        If InternetGetConnectedState(lngFlags, 0) Then
            ' True
            If lngFlags And InetConnState.lan Then
                ConnectionStateString = "LAN."
            ElseIf lngFlags And InetConnState.modem Then
                ConnectionStateString = "Modem."
            ElseIf lngFlags And InetConnState.configured Then
                ConnectionStateString = "Configured."
            ElseIf lngFlags And InetConnState.proxy Then
                ConnectionStateString = "Proxy"
            ElseIf lngFlags And InetConnState.ras Then
                ConnectionStateString = "RAS."
            ElseIf lngFlags And InetConnState.offline Then
                ConnectionStateString = "Offline."
            End If
        Else
            ' False
            ConnectionStateString = "Not Connected."
        End If

    End Function



#End Region
End Class 'AppStart


Public Class cPinger

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' code adapted from VB6 example. Used with permission. Copyright
    ' to original code shown below.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Copyright ©1996-2001 VBnet, Randy Birch, All Rights Reserved.
    ' Some pages may also contain other copyrights by the author.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' but you are expressly forbidden from selling or otherwise
    ' distributing this source code without prior written consent.
    ' This includes both posting free demo projects made from this
    ' code as well as reproducing the code in text or html format.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Const IP_SUCCESS As Integer = 0
    Private Const IP_STATUS_BASE As Integer = 11000
    Private Const IP_BUF_TOO_SMALL As Integer = (11000 + 1)
    Private Const IP_DEST_NET_UNREACHABLE As Integer = (11000 + 2)
    Private Const IP_DEST_HOST_UNREACHABLE As Integer = (11000 + 3)
    Private Const IP_DEST_PROT_UNREACHABLE As Integer = (11000 + 4)
    Private Const IP_DEST_PORT_UNREACHABLE As Integer = (11000 + 5)
    Private Const IP_NO_RESOURCES As Integer = (11000 + 6)
    Private Const IP_BAD_OPTION As Integer = (11000 + 7)
    Private Const IP_HW_ERROR As Integer = (11000 + 8)
    Private Const IP_PACKET_TOO_BIG As Integer = (11000 + 9)
    Private Const IP_REQ_TIMED_OUT As Integer = (11000 + 10)
    Private Const IP_BAD_REQ As Integer = (11000 + 11)
    Private Const IP_BAD_ROUTE As Integer = (11000 + 12)
    Private Const IP_TTL_EXPIRED_TRANSIT As Integer = (11000 + 13)
    Private Const IP_TTL_EXPIRED_REASSEM As Integer = (11000 + 14)
    Private Const IP_PARAM_PROBLEM As Integer = (11000 + 15)
    Private Const IP_SOURCE_QUENCH As Integer = (11000 + 16)
    Private Const IP_OPTION_TOO_BIG As Integer = (11000 + 17)
    Private Const IP_BAD_DESTINATION As Integer = (11000 + 18)
    Private Const IP_ADDR_DELETED As Integer = (11000 + 19)
    Private Const IP_SPEC_MTU_CHANGE As Integer = (11000 + 20)
    Private Const IP_MTU_CHANGE As Integer = (11000 + 21)
    Private Const IP_UNLOAD As Integer = (11000 + 22)
    Private Const IP_ADDR_ADDED As Integer = (11000 + 23)
    Private Const IP_GENERAL_FAILURE As Integer = (11000 + 50)
    Private Const MAX_IP_STATUS As Integer = (11000 + 50)
    Private Const IP_PENDING As Integer = (11000 + 255)
    Private Const PING_TIMEOUT As Integer = 1200
    Private Const WS_VERSION_REQD As Integer = &H101
    Private Const MIN_SOCKETS_REQD As Integer = 1
    Private Const SOCKET_ERROR As Integer = -1
    Private Const INADDR_NONE As Integer = &HFFFFFFFF
    Private Const MAX_WSADescription As Integer = 256
    Private Const MAX_WSASYSStatus As Integer = 128

    Private Structure ICMP_OPTIONS
        Public Ttl As Byte
        Public Tos As Byte
        Public Flags As Byte
        Public OptionsSize As Byte
        Public OptionsData As Integer
    End Structure

    Private Structure ICMP_ECHO_REPLY
        Public Address As Integer
        Public status As Integer
        Public RoundTripTime As Integer
        Public DataSize As Integer           'formerly integer
        'Public Reserved        As Integer
        Public DataPointer As Integer
        Public Options As ICMP_OPTIONS
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=250)> Dim myData As String
    End Structure

    Private Structure WSADATA
        Public wVersion As Integer
        Public wHighVersion As Integer
        'Public szDescription(0 To MAX_WSADescription) As Byte
        'Public szSystemStatus(0 To MAX_WSASYSStatus) As Byte
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_WSADescription + 1)> Public szDescription As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_WSASYSStatus + 1)> Public szSystemStatus As String
        Public wMaxSockets As Integer
        Public wMaxUDPDG As Integer
        Public dwVendorInfo As Integer
    End Structure

    Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Integer
    Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Integer) As Integer

    Private Declare Function IcmpSendEcho Lib "icmp.dll" _
       (ByVal IcmpHandle As Integer, _
     ByVal DestinationAddress As UInt32, _
     ByVal RequestData As String, _
     ByVal RequestSize As Integer, _
     ByVal RequestOptions As Integer, _
     ByRef ReplyBuffer As ICMP_ECHO_REPLY, _
     ByVal ReplySize As Integer, _
     ByVal Timeout As Integer) As Integer

    Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Integer
    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, ByRef lpWSADATA As WSADATA) As Integer
    Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Integer

    '---------------------------- end API declares ----------------------------
    Private FHostName As String
    Private FAddress As String
    Private FdwAddress As Long
    Private FEcho As ICMP_ECHO_REPLY
    Private FStatus As Integer
    Private FInitialized As Boolean

    ReadOnly Property Status() As Integer
        Get
            Return FStatus
        End Get
    End Property

    ReadOnly Property Address() As String
        Get
            Return FAddress
        End Get
    End Property

    ReadOnly Property RoundTripTime() As Integer
        Get
            Return FEcho.RoundTripTime
        End Get
    End Property

    ReadOnly Property pDataSize() As Integer
        Get
            Return FEcho.DataSize
        End Get
    End Property


    Private Sub ResolveHostName()

        'converts a host name to an IP address, both string and int form.

        Dim IPAddress As IPAddress
        Dim IPHE As IPHostEntry

        Try
            IPHE = Dns.GetHostByName(FHostName)
            If IPHE.AddressList.Length > 0 Then

                IPAddress = IPHE.AddressList(0)
                FAddress = IPAddress.ToString
                'La ligne suivante cause un Warning (obsolete) dans la compilation mais elle peut être ignorée sécuritairement
                'FdwAddress = IPAddress.Address
                '142.117.32.236  = 142 + (117 * 256 ^1) + (32 * 256 ^2) + (236 * 256 ^ 3) = 3961550222
                Dim intIPAddress As Int64
                intIPAddress = IPAddress.GetAddressBytes(0)
                intIPAddress += (IPAddress.GetAddressBytes(1) * 256)
                intIPAddress += (IPAddress.GetAddressBytes(2) * 65536)
                intIPAddress += CType(IPAddress.GetAddressBytes(3), Int64) * 16777216
                FdwAddress = intIPAddress

            Else
                FdwAddress = Convert.ToInt64(INADDR_NONE)
            End If

        Catch oEX As Exception
            FdwAddress = INADDR_NONE
        End Try
    End Sub

    Public Function PingHostName(ByVal cHostName As String) As String

        Dim hPort As Integer
        Const DATATOSEND As String = "TEST"

        FHostName = cHostName
        Call ResolveHostName()
        Call SocketsInitialize()
        Try

            'if a valid address..
            If FdwAddress <> Convert.ToInt64(INADDR_NONE) Then

                'open a port
                hPort = IcmpCreateFile()

                'and if successful,
                If hPort <> 0 Then

                    'ping it.
                    Call IcmpSendEcho(hPort, Convert.ToUInt32(FdwAddress), DATATOSEND, DATATOSEND.Length, 0, FEcho, Marshal.SizeOf(FEcho), PING_TIMEOUT)

                    'return the status as ping succes and close
                    FStatus = FEcho.status
                    Call IcmpCloseHandle(hPort)

                    PingHostName &= "Reply from " & FAddress
                    PingHostName &= ": bytes=32 time="
                    PingHostName &= FEcho.RoundTripTime & "ms TTL=" & FEcho.Options.Ttl
                End If

            Else
                'the address format was probably invalid
                FStatus = INADDR_NONE
                PingHostName = "host name not found"

            End If

        Finally
            Call SocketsCleanup()
        End Try
    End Function

    Private Sub SocketsCleanup()
        If WSACleanup() <> 0 Then
            MessageBox.Show("Windows Sockets error occurred in Cleanup.")
        End If
    End Sub


    Private Function SocketsInitialize() As Boolean

        Dim WSAD As WSADATA

        SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS

    End Function

    Public Function GetStatusText() As String

        Dim cMsg As String

        Select Case FStatus
            Case IP_SUCCESS : cMsg = "ip success"
            Case INADDR_NONE : cMsg = "inet_addr: bad IP format"
            Case IP_BUF_TOO_SMALL : cMsg = "ip buf too_small"
            Case IP_DEST_NET_UNREACHABLE : cMsg = "ip dest net unreachable"
            Case IP_DEST_HOST_UNREACHABLE : cMsg = "ip dest host unreachable"
            Case IP_DEST_PROT_UNREACHABLE : cMsg = "ip dest prot unreachable"
            Case IP_DEST_PORT_UNREACHABLE : cMsg = "ip dest port unreachable"
            Case IP_NO_RESOURCES : cMsg = "ip no resources"
            Case IP_BAD_OPTION : cMsg = "ip bad option"
            Case IP_HW_ERROR : cMsg = "ip hw_error"
            Case IP_PACKET_TOO_BIG : cMsg = "ip packet too_big"
            Case IP_REQ_TIMED_OUT : cMsg = "ip req timed out"
            Case IP_BAD_REQ : cMsg = "ip bad req"
            Case IP_BAD_ROUTE : cMsg = "ip bad route"
            Case IP_TTL_EXPIRED_TRANSIT : cMsg = "ip ttl expired transit"
            Case IP_TTL_EXPIRED_REASSEM : cMsg = "ip ttl expired reassem"
            Case IP_PARAM_PROBLEM : cMsg = "ip param_problem"
            Case IP_SOURCE_QUENCH : cMsg = "ip source quench"
            Case IP_OPTION_TOO_BIG : cMsg = "ip option too_big"
            Case IP_BAD_DESTINATION : cMsg = "ip bad destination"
            Case IP_ADDR_DELETED : cMsg = "ip addr deleted"
            Case IP_SPEC_MTU_CHANGE : cMsg = "ip spec mtu change"
            Case IP_MTU_CHANGE : cMsg = "ip mtu_change"
            Case IP_UNLOAD : cMsg = "ip unload"
            Case IP_ADDR_ADDED : cMsg = "ip addr added"
            Case IP_GENERAL_FAILURE : cMsg = "ip general failure"
            Case IP_PENDING : cMsg = "ip pending"
            Case PING_TIMEOUT : cMsg = "ping timeout"
            Case Else : cMsg = "unknown  cMsg returned"
        End Select

        Return CStr(FStatus) & "   [ " & cMsg & " ]"

    End Function
End Class
