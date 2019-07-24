'-------------------------------------------------------------------------------
' MUSTER.Info.CommentsInfo
'   Provides the container to persist MUSTER Comments state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        PN      12/13/04    Original class definition.
'  1.1        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        AB      02/17/05    Added AgeThreshold and IsAgedData Attributes
'  1.3        MNR     03/22/05    Updated Constructor New(ByVal drComments As DataRow) to check for System.DBNull.Value
'
'
' Function          Description
' New()             Instantiates an empty CommentsInfo object.
' New(ID, Name, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                   Instantiates a populated CommentsInfo object.
' New(dr)           Instantiates a populated CommentsInfo object taking member state
'                       from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
' IsDirty           Indicates if the user state has been altered since it was
'                       last loaded from or saved to the repository.
' AgeThreshold      Indicates the number of minutes old data can be before it should be 
'                       refreshed from the DB.  Data should only be refreshed when Retrieved
'                       and when IsDirty is false
' IsAgedData        Will return true if the data has been held longer than the AgeThreshold
'
'-------------------------------------------------------------------------------
Namespace MUSTER.Info

    <Serializable()> _
Public Class CommentsInfo

#Region "Public Events"
        Public Event CommentsInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private strModule As String
        Private nId As Integer
        Private strWorkFlow As String
        Private strComment As String
        Private dtCommentDate As Date
        Private strUserId As String
        Private nCommentID As Integer
        Private strCommentScope As String
        Private nEntityID As Integer
        Private strEntityAdditionalInfo As String
        Private nEntityType As Integer
        Private bolDeleted As Boolean

        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private ostrModule As String
        Private onId As Integer
        Private ostrWorkFlow As String
        Private ostrComment As String
        Private odtCommentDate As Date
        Private ostrUserId As String
        Private onCommentID As Integer
        Private ostrCommentScope As String
        Private onEntityID As Integer
        Private ostrEntityAdditionalInfo As String
        Private onEntityType As Integer
        Private obolDeleted As Boolean

        Private bolIsDirty As Boolean = False
        Private dtDataAge As DateTime

        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal CommentID As Integer, _
                ByVal EntityID As Integer, _
                ByVal EntityAdditionalInfo As String, _
                ByVal EntityType As Integer, _
                ByVal Comment As String, _
                ByVal CommentScope As String, _
                ByVal Deleted As Boolean, _
                ByVal UserId As String, _
                ByVal CommentDate As Date, _
                ByVal ModuleName As String, _
                ByVal CreatedBy As String, _
                ByVal CreatedOn As Date, _
                ByVal ModifiedBy As String, _
                ByVal LastEdited As Date)
            onCommentID = CommentID
            onEntityID = EntityID
            ostrEntityAdditionalInfo = EntityAdditionalInfo
            onEntityType = EntityType
            ostrComment = Comment
            ostrCommentScope = CommentScope
            obolDeleted = Deleted
            ostrUserId = UserId
            odtCommentDate = CommentDate
            ostrModule = ModuleName
            strCreatedBy = CreatedBy
            dtCreatedOn = CreatedOn
            strModifiedBy = ModifiedBy
            dtModifiedOn = LastEdited
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drComments As DataRow)
            Try
                onCommentID = drComments.Item("COMMENT_ID")
                onEntityID = drComments.Item("ENTITY ID")
                ostrEntityAdditionalInfo = IIf(drComments.Item("ENTITY_ADDITIONAL_INFO") Is DBNull.Value, String.Empty, drComments.Item("ENTITY_ADDITIONAL_INFO"))
                onEntityType = drComments.Item("ENTITY_TYPE")
                ostrComment = drComments.Item("COMMENT")
                ostrCommentScope = drComments.Item("VIEWABLE BY")
                obolDeleted = drComments.Item("DELETED")
                ostrUserId = drComments.Item("USER ID")
                odtCommentDate = drComments.Item("COMMENT_DATE")
                ostrModule = drComments.Item("MODULE")
                strCreatedBy = drComments.Item("CREATEDBY")
                dtCreatedOn = drComments.Item("CREATED ON")
                strModifiedBy = IIf(drComments.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drComments.Item("LAST_EDITED_BY"))
                dtModifiedOn = IIf(drComments.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drComments.Item("DATE_LAST_EDITED"))
                dtDataAge = Now()
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nCommentID = onCommentID
            nEntityID = onEntityID
            strEntityAdditionalInfo = ostrEntityAdditionalInfo
            nEntityType = onEntityType
            strComment = ostrComment
            strCommentScope = ostrCommentScope
            bolDeleted = obolDeleted
            strUserId = ostrUserId
            strModule = ostrModule
            dtCommentDate = odtCommentDate
            bolIsDirty = False
        End Sub
        Public Sub Archive()
            onCommentID = nCommentID
            onEntityID = nEntityID
            ostrEntityAdditionalInfo = strEntityAdditionalInfo
            onEntityType = nEntityType
            ostrComment = strComment
            ostrCommentScope = strCommentScope
            obolDeleted = bolDeleted
            ostrUserId = strUserId
            odtCommentDate = dtCommentDate
            ostrModule = strModule
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            bolIsDirty = (nCommentID <> onCommentID) Or _
                         (nEntityID <> onEntityID) Or _
                         (strEntityAdditionalInfo <> ostrEntityAdditionalInfo) Or _
                         (nEntityType <> onEntityType) Or _
                         (strComment <> ostrComment) Or _
                         (strCommentScope <> ostrCommentScope) Or _
                         (bolDeleted <> obolDeleted) Or _
                         (strUserId <> ostrUserId) Or _
                         (dtCommentDate <> odtCommentDate) Or _
                         (strModule <> ostrModule)
            If obolIsDirty <> bolIsDirty Then
                RaiseEvent CommentsInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onCommentID = 0
            onEntityID = 0
            ostrEntityAdditionalInfo = String.Empty
            onEntityType = 0
            ostrComment = String.Empty
            ostrCommentScope = String.Empty
            ostrUserId = String.Empty
            odtCommentDate = System.DateTime.Now
            obolDeleted = False
            ostrModule = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nCommentID
            End Get
            Set(ByVal Value As Integer)
                nCommentID = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EntityID() As Integer
            Get
                Return nEntityID
            End Get
            Set(ByVal Value As Integer)
                nEntityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EntityAdditionalInfo() As String
            Get
                Return strEntityAdditionalInfo
            End Get
            Set(ByVal Value As String)
                strEntityAdditionalInfo = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EntityType() As Integer
            Get
                Return nEntityType
            End Get
            Set(ByVal Value As Integer)
                nEntityType = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property CommentsScope() As String
            Get
                Return strCommentScope
            End Get
            Set(ByVal Value As String)
                strCommentScope = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property UserID() As String
            Get
                Return strUserId
            End Get
            Set(ByVal Value As String)
                strUserId = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModuleName() As String
            Get
                Return strModule
            End Get
            Set(ByVal Value As String)
                strModule = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WorkFlow() As String
            Get
                Return strWorkFlow
            End Get
            Set(ByVal Value As String)
                strWorkFlow = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Comments() As String
            Get
                Return strComment
            End Get
            Set(ByVal Value As String)
                strComment = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CommentDate() As Date
            Get
                Return dtCommentDate
            End Get
            Set(ByVal Value As Date)
                dtCommentDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property

        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get

            Set(ByVal value As Int16)
                nAgeThreshold = Int16.Parse(value)
            End Set
        End Property

        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get

        End Property


#Region "IAccessors"
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
#End Region
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace

