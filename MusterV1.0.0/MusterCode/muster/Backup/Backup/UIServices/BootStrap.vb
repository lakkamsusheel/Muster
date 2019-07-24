Module BootStrap

    Private _progressScreenThread As System.Threading.Thread
    Private _ticklerScreenThread As System.Threading.Thread
    Private _ticklerAlertThread As System.Threading.Thread
    Public _container As MusterContainer


    Sub ProgressThreadStarted()

        Dim pScreen As New ProgressScreen(_container)

        _container.ProgressScreen = pScreen

    End Sub


    Sub AlertThreadStarted()

        Dim pAlert As New TicklerAlert(_container)

        _container.TAlert = pAlert

    End Sub


    Sub TicklerThreadStarted()

        Dim tScreen As New TickerScreen(_container)

        _container.TicklerScreen = tScreen

    End Sub



    Sub main(ByVal args() As String)

        _container = New MusterContainer

        _progressScreenThread = New System.Threading.Thread(AddressOf ProgressThreadStarted)

        _progressScreenThread.Start()

        _ticklerScreenThread = New System.Threading.Thread(AddressOf TicklerThreadStarted)

        _ticklerScreenThread.Start()

        _ticklerAlertThread = New System.Threading.Thread(AddressOf AlertThreadStarted)

        _ticklerAlertThread.Start()

        If args.Length > 0 AndAlso args(0) = "/I" Then
            _container.Inspector = True
        End If



        Application.Run(_container)

        _container.TAlert.Close()

        _progressScreenThread = Nothing

        _ticklerScreenThread = Nothing

        _ticklerAlertThread = Nothing

    End Sub


End Module
