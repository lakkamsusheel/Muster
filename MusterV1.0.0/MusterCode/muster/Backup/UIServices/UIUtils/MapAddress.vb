Imports System.IO
Imports System.Net
Imports System.Text
Imports mshtml
Imports SHDocVw.WebBrowserClass


Public Class MapAddress

    Private strAddress As String
    Private strCity As String
    Private IEinfo As ProcessStartInfo
    Private strState As String
    Private bWebBrowser As SHDocVw.InternetExplorer
    Private yahooStr As String

    Public Property webBrowser() As SHDocVw.InternetExplorer
        Get
            If bWebBrowser Is Nothing Then

                Try

                    bWebBrowser = New SHDocVw.InternetExplorer

                    AddHandler bWebBrowser.WindowClosing, AddressOf DisposeScreen
                Catch ex As Exception
                    Throw New Exception(String.Format("Web browser Tool COM Access failed: {0}", ex.Message))
                End Try


            End If

            Return bWebBrowser
        End Get
        Set(ByVal Value As SHDocVw.InternetExplorer)
            bWebBrowser = Value
        End Set
    End Property

    Public ReadOnly Property YahooStrZoom(ByVal zoom As Integer)

        Get
            Return yahooStr.Replace("XXXX", zoom.ToString)
        End Get

    End Property

    Sub New(ByVal address As String, ByVal address2 As String, ByVal city As String, ByVal state As String, ByVal longitude As Decimal, ByVal latitude As Decimal)

        strCity = city
        strState = state

        Try
            If Not File.Exists("c:\windows\system32\ieframe.dll") Then
                File.Copy("ieframe.dll", "c:\windows\system32\ieframe.dll")
            End If
        Catch ex As Exception
            Throw New Exception("Missing IEFRAME.DLL: Need to find and copy This Internet explorer DLL into your Muster\1.0.0.0 folder")
        End Try

        Try
            If Process.GetProcessesByName("iexplore") Is Nothing OrElse Process.GetProcessesByName("iexplore").Length = 0 Then
                IEinfo = New ProcessStartInfo
                IEinfo.CreateNoWindow = False
                IEinfo.FileName = "IEXPLORE"
                IEinfo.WindowStyle = ProcessWindowStyle.Hidden

                Try
                    Process.Start(IEinfo)
                Catch
                    Throw New Exception("Please Install Internet explore 6 or higher to view maps")
                End Try

            End If


            If Not (longitude = -1 AndAlso latitude = -1) Then
                strAddress = String.Format("clon={0}&clat={1}", longitude * -1, latitude)

            ElseIf address.ToUpper.Replace(".", "").IndexOf("PO BOX") > -1 OrElse address.ToUpper.StartsWith("ROUTE") OrElse address = String.Empty Then
                If address2.ToUpper.Replace(".", "").IndexOf("PO BOX") > -1 OrElse address2.ToUpper.StartsWith("ROUTE") OrElse address2 = String.Empty Then
                    strAddress = String.Format("q1={0},+{1}", strCity.Replace(" ", "+"), strState.Replace(" ", "+"))
                Else
                    strAddress = String.Format("q1={0},+{1},+{2}", address2, strCity.Replace(" ", "+"), strState.Replace(" ", "+"))
                End If
            Else
                strAddress = String.Format("q1={0},+{1},+{2}", address, strCity.Replace(" ", "+"), strState.Replace(" ", "+"))
            End If




            yahooStr = String.Format("http://maps.yahoo.com/print?mvt=m&zoom=XXXX&{0}", _
                                     IIf(strAddress.Length > 0, strAddress.Replace(" ", "+"), "q1=MS"))


        Catch ex As Exception
            Throw New Exception(String.Format("Web Mapping Class Initializer: {0}", ex.Message))
        End Try


    End Sub

    Public Sub dispose()

        Try
            If Not IEinfo Is Nothing Then
                IEinfo = Nothing
            End If

            For Each p As Process In Process.GetProcessesByName("iexplore")
                If p.MainWindowHandle.ToInt32 <= 0 Then
                    p.Kill()
                End If
            Next
        Catch
        End Try

        KillWebBrowserCOM(True)

    End Sub

    Private Sub KillWebBrowserCOM(Optional ByVal hideErr As Boolean = False)

        Try

            If Not bWebBrowser Is Nothing Then
                bWebBrowser.Quit()
                bWebBrowser = Nothing
            End If

        Catch ex As Exception
            If Not hideErr Then
                Throw New Exception(String.Format("Quitting Internal Web Explorer: {0}", ex.Message))
            End If
        End Try

    End Sub

    Public Function LoadMapOnWeb(ByRef b As SHDocVw.InternetExplorer, Optional ByVal webPage As String = "About:Blank") As mshtml.IHTMLDocument2



        b.Navigate2(webPage)

        Do
        Loop Until b.ReadyState = SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE

        Return b.Document

    End Function

    Private Sub DisposeScreen(ByVal ischildWindow As Boolean, ByRef cancel As Boolean)
        KillWebBrowserCOM()
    End Sub


    Public Sub ShowOnScreen()

        _container.Cursor = Windows.Forms.Cursors.WaitCursor


        Dim i As New Mapdisplay(LoadImage(10), LoadImage(13), LoadImage(16))

        i.ShowDialog()
        _container.Cursor = Windows.Forms.Cursors.Arrow

    End Sub





    Public Function LoadImage(ByVal zoomLevel As Integer) As Image


        Dim doc As mshtml.IHTMLDocument2
        Dim img As Image

        Try

            bWebBrowser = Nothing

            doc = LoadMapOnWeb(webBrowser, YahooStrZoom(zoomLevel))

            Dim elem As mshtml.IHTMLElement
            Dim count As Integer = 0


            If webBrowser.ReadyState = SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE Then

                For Each elem2 As IHTMLElement In doc.all

                    If TypeOf elem2 Is HTMLCommentElement AndAlso Not elem2.id Is Nothing AndAlso elem2.id.ToUpper = "YMAP_F_IMG" Then
                        elem = elem2
                        Exit For
                    End If

                Next

                count += 1

            End If




            doc.close()

            If Not elem Is Nothing AndAlso Not elem.id Is Nothing Then


                Dim r As IHTMLControlRange

                r = DirectCast(doc.body, IHTMLElement).createControlRange

                r.add(elem)

                count = 0
                Do

                    r.execCommand("Copy")

                    If Clipboard.GetDataObject.GetDataPresent(DataFormats.Bitmap, True) Then
                        img = Clipboard.GetDataObject.GetData(DataFormats.Bitmap, True)
                        count += 1
                    Else
                        Throw New Exception("Error: Image has been found but was not properly formatted into the clipboard for viewing")
                        count = 100000
                    End If
                Loop Until img.Height > 100 Or count >= 100000


            Else
                Throw New Exception("The (yahoo maps page) did not show the proper ID for the MAP picture. Cannot load any images")
            End If

            KillWebBrowserCOM()

            Return img


        Catch ex As Exception

            Throw New Exception(String.Format("Map Address Module: {0}", ex.Message))
        End Try

    End Function

End Class
