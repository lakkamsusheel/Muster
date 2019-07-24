Public Class TicklerManager

#Region "private members"

    Private aMsgMemory As Collections.Specialized.StringDictionary
    Private aMsgMemory2 As Collections.Specialized.StringDictionary

    Private aMsgDataTable As Data.DataTable
    Private _readFlag As DataAccess.TicklerMessageDB.MessageBOOLEnum = DataAccess.TicklerMessageDB.MessageBOOLEnum.BothYesNo
    Private _completeFlag As DataAccess.TicklerMessageDB.MessageBOOLEnum = DataAccess.TicklerMessageDB.MessageBOOLEnum.No

    Private oMessage As BusinessLogic.pTicklerMessage
    Private _unreadMessages As Integer = 0

#End Region

#Region "public members"

    Public Property Message() As BusinessLogic.pTicklerMessage
        Get

            If oMessage Is Nothing Then
                oMessage = New BusinessLogic.pTicklerMessage
            End If

            Return oMessage
        End Get
        Set(ByVal Value As BusinessLogic.pTicklerMessage)
            oMessage = Value
        End Set
    End Property

    Public Event NewFound(ByVal ds As DataTable, ByVal force As Boolean)

    Public ReadOnly Property UnReadMessages() As Integer
        Get
            Return _unreadMessages
        End Get
    End Property



    Public Property MsgDataTable() As Data.DataTable

        Get

            Return aMsgDataTable

        End Get

        Set(ByVal Value As Data.DataTable)
            aMsgDataTable = Value
        End Set

    End Property

    Public Property MsgMemory() As Collections.Specialized.StringDictionary

        Get

            If aMsgMemory Is Nothing Then
                aMsgMemory = New Collections.Specialized.StringDictionary
            End If

            Return aMsgMemory
        End Get

        Set(ByVal Value As Collections.Specialized.StringDictionary)
            aMsgMemory = Value
        End Set
    End Property

    Public Property MsgMemory2() As Collections.Specialized.StringDictionary

        Get

            If aMsgMemory2 Is Nothing Then
                aMsgMemory2 = New Collections.Specialized.StringDictionary
            End If

            Return aMsgMemory2
        End Get

        Set(ByVal Value As Collections.Specialized.StringDictionary)
            aMsgMemory2 = Value
        End Set
    End Property

#End Region

#Region "construct"

    Sub New()

        MyBase.New()
    End Sub


    Sub dispose()

        If Not MsgDataTable Is Nothing Then
            MsgDataTable.Dispose()
        End If

        If Not oMessage Is Nothing Then
            oMessage = Nothing
        End If

        If Not MsgMemory Is Nothing Then

            MsgMemory.Clear()
            MsgMemory = Nothing
        End If

    End Sub

#End Region

#Region "public members"



    Sub refreshMessages(ByVal userID As String, Optional ByVal force As Boolean = False, Optional ByVal showCompleted As Boolean = False, Optional ByVal showFrom As Boolean = False)

        getData(userID, showCompleted, showFrom)
        CheckForNew(force)

        If Not MsgDataTable Is Nothing Then
            MsgDataTable.Dispose()
        End If

    End Sub


    Sub CheckForNew(Optional ByVal force As Boolean = False)

        Dim isNew As Boolean = False
        Dim isLong As Boolean = False
        Dim cnt As Integer = 0
        Dim AlreadyOpen As Boolean = False

        'clean scans 

        _unreadMessages = 0

        If Not _container Is Nothing AndAlso Not _container.TicklerScreen Is Nothing AndAlso _container.TicklerScreen.Visible Then
            AlreadyOpen = True
        End If

        If Not MsgDataTable Is Nothing Then

            For Each dr As DataRow In MsgDataTable.Rows

                If dr("MsgRead").ToString = "0" Then _unreadMessages += 1

                With DirectCast(dr("MsgID"), String)

                    If MsgMemory.ContainsKey(.ToString) Then

                        If (Now.Subtract(Convert.ToDateTime(Me.MsgMemory.Item(.ToString))).TotalHours >= 24 Or (Now.Subtract(Convert.ToDateTime(Me.MsgMemory.Item(.ToString))).TotalMinutes >= 15 And MsgMemory2.Item(.ToString) = "false")) Then
                            isLong = True
                            dr("IsNew") = "true"

                            If force OrElse AlreadyOpen Then
                                Me.MsgMemory2.Item(.ToString) = "true"
                            End If

                            MsgMemory.Item(.ToString) = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)

                        Else

                            dr("IsNew") = "false"

                            If force OrElse AlreadyOpen Then
                                Me.MsgMemory2.Item(.ToString) = "true"
                            End If

                        End If

                    Else
                        isNew = True
                        dr("IsNew") = "true"

                        MsgMemory.Add(.ToString, dr("DatePulled"))
                        MsgMemory2.Add(.ToString, "false")

                    End If
                End With

            Next

            MsgDataTable.DefaultView.Sort = "IsNew DESC, MSgRead, Completed"

            For Each dr As DataRowView In MsgDataTable.DefaultView

                dr("RowNum") = cnt + 1
                cnt += 1

            Next

            If (isNew Or isLong) Or force Then

                RaiseEvent NewFound(MsgDataTable, force)
            End If


        End If

    End Sub

#End Region

#Region "private members"


    Public Shared Function saveData(ByVal tickler As BusinessLogic.pTicklerMessage) As String

        Dim retVal As String = String.Empty

        tickler.Save(retVal)

        If retVal.Length > 0 Then
            Throw New Exception(String.Format("Error Saving Tickler Message: {0}", retVal))
        End If

    End Function


    Private Sub getData(ByVal userID As String, Optional ByVal showCompleted As Boolean = False, Optional ByVal showFrom As Boolean = False)

        Try

            If showFrom Then
                MsgDataTable = Message.GetMessagesFromUser(userID, _readFlag, IIf(showCompleted, DataAccess.TicklerMessageDB.MessageBOOLEnum.Yes, DataAccess.TicklerMessageDB.MessageBOOLEnum.No))
            Else
                MsgDataTable = Message.GetMessagesByUser(userID, _readFlag, IIf(showCompleted, DataAccess.TicklerMessageDB.MessageBOOLEnum.Yes, DataAccess.TicklerMessageDB.MessageBOOLEnum.No))

            End If


        Catch ex As Exception

            If Not oMessage Is Nothing Then
                oMessage = Nothing
            End If

        End Try
    End Sub

#End Region

End Class
