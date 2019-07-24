Public Class LookupProperty
    Private nId As Integer
    Private strType As String
    Private nParentId As Integer

    Public Sub New(ByVal strTypeValue As String, ByVal nIdValue As Integer)
        strType = strTypeValue
        nId = nIdValue
    End Sub

    Public Sub New(ByVal strTypeValue As String, ByVal nIdValue As Integer, ByVal nParentIdValue As Integer)
        strType = strTypeValue
        nId = nIdValue
        nParentId = nParentIdValue
    End Sub

    Public Property Type() As String
        Get
            Return strType
        End Get
        Set(ByVal Value As String)
            strType = Value
        End Set
    End Property

    Public Property Id() As Integer
        Get
            Return nId
        End Get
        Set(ByVal Value As Integer)
            nId = Value
        End Set
    End Property

    Public Property ParentId() As Integer
        Get
            Return nParentId
        End Get
        Set(ByVal Value As Integer)
            nParentId = Value
        End Set
    End Property
End Class

