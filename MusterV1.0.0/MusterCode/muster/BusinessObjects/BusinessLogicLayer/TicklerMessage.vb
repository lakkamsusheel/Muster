
Namespace MUSTER.BusinessLogic
    <Serializable()> Public Class pTicklerMessage
        '-------------------------------------------------------------------------------
        ' MUSTER.BusinessLogic.pTicklerMessage
        '   Provides the operations required to manipulate a Tickler Message object.
        '
        ' Copyright (C) 2009  CIBER, Inc.
        ' All rights reserved.
        '
        ' Release   Initials             Date          Description
        '  1.0       Thomas Franey       5/29/2009      Original class definition
        '
        ' Function          Description
        '-------------------------------------------------------------------------------
        ' Attribute          Description
        '-------------------------------------------------------------------------------
#Region "Private Member Variables"
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private oTicklerMessageDB As MUSTER.DataAccess.TicklerMessageDB = New MUSTER.DataAccess.TicklerMessageDB
        Private oTicklerMessageInfo As MUSTER.Info.TicklerMessageInfo = New MUSTER.Info.TicklerMessageInfo
#End Region

#Region "Public Events"

        Public Delegate Sub TicklerChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub TicklerErrEventHandler(ByVal MsgStr As String)
        ' indicates change in the underlying TecDocInfo structure
        Public Delegate Sub TicklerMsgInfoChanged()

        Public Event TecDocChanged As TicklerChangedEventHandler
        Public Event TecDocErr As TicklerErrEventHandler

#End Region
#Region "Constructors"
        Public Sub New()
            oTicklerMessageInfo = New MUSTER.Info.TicklerMessageInfo
        End Sub
        Public Sub New(ByVal MsgID As String)
            oTicklerMessageInfo = New MUSTER.Info.TicklerMessageInfo

            Me.Retrieve(MsgID)
        End Sub

#End Region
#Region "Exposed Attributes"

        ' Gets/Sets the read flag 
        Public ReadOnly Property Read() As Boolean
            Get
                Return oTicklerMessageInfo.Read
            End Get
        End Property

        Public ReadOnly Property Completed() As Boolean
            Get
                Return oTicklerMessageInfo.Completed
            End Get
        End Property

        Public Property IsIssue() As Boolean
            Get
                Return oTicklerMessageInfo.IsIssue
            End Get
            Set(ByVal Value As Boolean)
                oTicklerMessageInfo.IsIssue = Value
            End Set
        End Property



        Public ReadOnly Property IsDirty() As Boolean
            Get
                If oTicklerMessageInfo.IsDirty Then
                    Return True
                End If
                Return False
            End Get
        End Property

        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oTicklerMessageInfo.CreatedOn
            End Get
        End Property

        ' The date on which user has read the message
        Public ReadOnly Property DateRead() As DateTime
            Get
                Return oTicklerMessageInfo.DateRead
            End Get
        End Property

        ' The date on which use has assign as completed
        Public ReadOnly Property DateCompleted() As DateTime
            Get
                Return oTicklerMessageInfo.DateCompleted
            End Get
        End Property

        ' The dateset for message to post
        Public Property PostDate() As Date
            Get
                Return oTicklerMessageInfo.PostDate
            End Get

            Set(ByVal Value As Date)
                oTicklerMessageInfo.PostDate = Value
            End Set
        End Property

        ' The module ID associated with tickler message
        Public Property ModuleID() As Integer
            Get
                Return oTicklerMessageInfo.ModuleID
            End Get

            Set(ByVal Value As Integer)
                Me.oTicklerMessageInfo.ModuleID = Value
            End Set
        End Property


        ' Gets/Sets the physical file name of the image associated with the tickler message
        Public Property ImageFile() As String
            Get
                Return oTicklerMessageInfo.ImageFile
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.ImageFile = Value
            End Set
        End Property

        ' Gets the ID of the tickler message
        Public Property ID() As String
            Get
                Return oTicklerMessageInfo.ID
            End Get

            Set(ByVal Value As String)
                oTicklerMessageInfo.ID = Value
            End Set
        End Property


        ' Gets/Sets the subject of the user tickler message
        Public Property Subject() As String
            Get
                Return oTicklerMessageInfo.Subject
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.Subject = Value
            End Set
        End Property


        ' Gets/Sets the message of the user tickler message
        Public Property Message() As String
            Get
                Return oTicklerMessageInfo.Message
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.Message = Value
            End Set
        End Property

        ' Gets/Sets the Object ID (Entity ID)  of the user tickler message
        Public Property ObjectID() As String
            Get
                Return oTicklerMessageInfo.ObjectID
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.ObjectID = Value
            End Set
        End Property

        ' Gets/Sets the keyword (Entity type)  of the user tickler message
        Public Property Keyword() As String
            Get
                Return oTicklerMessageInfo.Keyword
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.Keyword = Value
            End Set
        End Property

        ' Gets/Sets Sender
        Public Property FromID() As String
            Get
                Return oTicklerMessageInfo.FromID
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.FromID = Value
            End Set
        End Property

        ' Gets/Sets receiver
        Public Property ToID() As String
            Get
                Return oTicklerMessageInfo.toID
            End Get
            Set(ByVal Value As String)
                oTicklerMessageInfo.toID = Value
            End Set
        End Property

        ' Gets the info object
        Public ReadOnly Property InfoObject() As MUSTER.Info.TicklerMessageInfo
            Get
                Return oTicklerMessageInfo
            End Get
        End Property
#End Region
#Region "Exposed Methods"
#Region "General Operations"
        Public Sub Clear()

            oTicklerMessageInfo = Nothing

            oTicklerMessageInfo = New MUSTER.Info.TicklerMessageInfo
        End Sub
        Public Sub Reset()
            oTicklerMessageInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal MsgID As String, Optional ByVal setRead As Boolean = False, Optional ByVal setCompleted As Boolean = False) As MUSTER.Info.TicklerMessageInfo
            Try
                oTicklerMessageInfo = Me.oTicklerMessageDB.DBGetByID(MsgID, setRead, setCompleted)
                If oTicklerMessageInfo.ID = "0" Then
                    nID = String.Empty
                End If

                Return oTicklerMessageInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        ' Saves the data in the current Info object
        Public Sub Save(ByRef returnVal As String)

            Dim strModuleName As String = String.Empty
            Dim bolSubmitForCalendar As Boolean
            Try
                If Me.ValidateData(strModuleName) Then

                    oTicklerMessageDB.Put(oTicklerMessageInfo, returnVal)

                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If

                    oTicklerMessageInfo.Archive()
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                'If errStr.Length > 0 Or Not validateSuccess Then
                '    RaiseEvent LustEventErr(errStr)
                'End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

#End Region
#Region " Populate Routines "

        Public Function GetMessagesByUser(ByVal userID As String, ByVal read As DataAccess.TicklerMessageDB.MessageBOOLEnum, ByVal completed As DataAccess.TicklerMessageDB.MessageBOOLEnum) As DataTable

            Try
                Return oTicklerMessageDB.DbGetTicklerList(userID, read, completed)
            Catch ex As Exception
                Throw ex
                Return Nothing
            End Try


        End Function

        Public Function GetMessagesFromUser(ByVal userID As String, ByVal read As DataAccess.TicklerMessageDB.MessageBOOLEnum, ByVal completed As DataAccess.TicklerMessageDB.MessageBOOLEnum) As DataTable

            Try
                Return oTicklerMessageDB.DbGetTicklerListSent(userID, read, completed)
            Catch ex As Exception
                Throw ex
                Return Nothing
            End Try


        End Function


#End Region
#End Region
    End Class
End Namespace
