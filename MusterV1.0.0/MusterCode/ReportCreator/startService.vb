Module startService

    Private WithEvents fMainModuleGUI As GUI
    Private rReport As ReportDeveloper


    Public Property MainModuleGUI() As GUI
        Get

            If fMainModuleGUI Is Nothing Then
                fMainModuleGUI = New GUI
            End If

            Return fMainModuleGUI
        End Get

        Set(ByVal Value As GUI)
            fMainModuleGUI = Value
        End Set

    End Property


    Sub main(ByVal ParamArray Args() As String)



        If Args Is Nothing OrElse Args.GetUpperBound(0) = -1 Then
            MainModuleGUI.ShowDialog()

        ElseIf Args(0) = "/h" OrElse Args(0) = "/?" Then
            Console.WriteLine("help: ReportCreator [crystalReportFullPath] [name of Created PDF File]  ([Args[..]])")

        ElseIf Args.GetUpperBound(0) < 1 Then
            Console.WriteLine("Please enter the new file name of the PDF file")


        Else

            ExecuteReport(Args)

        End If

    End Sub

    Sub ExecuteReport(ByVal Args() As String)

        Dim aList As ArrayList

        Try
            rReport = New ReportDeveloper(Args(0), System.Configuration.ConfigurationSettings.AppSettings.Item("SQLConnectionString"))


            rReport.InitializeReport()

            'Set Arguments For Parameters
            aList = New ArrayList
            For g As Integer = 2 To Args.GetUpperBound(0)
                aList.Add(Args(g))
            Next


            rReport.PushParametersInReport(aList)
            rReport.ExportFile(Args(1))

        Catch ex As Exception

            LogError(ex)

        Finally

            If Not rReport Is Nothing Then
                rReport.dispose()
            End If

            rReport = Nothing

            If Not aList Is Nothing Then
                aList.Clear()
            End If

            aList = Nothing
        End Try


    End Sub

    Sub LogError(ByVal ex As Exception)

        Console.WriteLine(String.Format("{0} - {1}", ex.Message, IIf(ex.InnerException Is Nothing, String.Empty, ex.InnerException.Message)))

#If DEBUG Then

        Debugger.Log(0, "Application Exception", String.Format("{0} - {1}", ex.Message, IIf(ex.InnerException Is Nothing, String.Empty, ex.InnerException.Message)))

#End If

    End Sub

    Sub FormClose(ByVal sender As Object, ByVal e As EventArgs) Handles fMainModuleGUI.Closed

        fMainModuleGUI.Dispose()

        End

    End Sub

End Module
