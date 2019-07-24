Namespace MUSTER.Info

    <Serializable()> _
Public Class UserAuditPoint
#Region "Private member variables"
        Private strUserModule As String
        Private dtUserEntered As DateTime
        Private dtUserExited As DateTime
        Private strGUID As String
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
        End Sub
        Sub New(ByVal ModuleName As String)
            Me.ModuleName = ModuleName

        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ModuleName() As String
            Get
                Return strUserModule
            End Get
            Set(ByVal Value As String)
                strUserModule = Value
                Me.EntryPoint = System.DateTime.Now
            End Set
        End Property

        Public Property EntryPoint() As DateTime
            Get

            End Get
            Set(ByVal Value As DateTime)
                dtUserEntered = Value
            End Set
        End Property

        Public Property ExitPoint() As DateTime
            Get

            End Get
            Set(ByVal Value As DateTime)
                dtUserExited = Value
            End Set
        End Property
        Public Property GUID() As String
            Get
                Return strGUID
            End Get
            Set(ByVal Value As String)
                strGUID = Value
            End Set
        End Property
#End Region
    End Class
End Namespace
