'-------------------------------------------------------------------------------
' MUSTER.Info.EntityInfo
'   Provides the container to persist MUSTER Entity data
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/19/04    Original class definition.
'  1.1        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        MNR       03/22/05    Added Constructor New(ByVal dr As DataRow)
'
' Function          Description
' New()             Instantiates an empty EntityInfo object.
' New(ID, Name, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                   Instantiates a populated EntityInfo object.
' New(ds)           Instantiates a populated EntityInfo object taking member state
'                       from the first row in the first table in the dataset provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
' Save()            Saves the object state to the repository.
'
'Attribute          Description
' ID                The unique identifier associated with the Entity in the repository.
' Name              The name of the Entity.
' IsDirty           Indicates if the Entity state has been altered since it was
'                       last loaded from or saved to the repository.
'-------------------------------------------------------------------------------

Namespace MUSTER.Info

    <Serializable()> _
      Public Class EntityInfo
        Implements iAccessors
#Region "Private member variables"

        Private oEntityID As Integer
        Private EntityID As Integer
        Private ostrEntityName As String
        Private strEntityName As String
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private ostrCreatedBy As String
        Private strCreatedBy As String
        Private ostrModifiedBy As String
        Private strModifiedBy As String
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
        End Sub
        Sub New(ByVal EntityID As Integer, _
            ByVal EntityName As String, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date)
            oEntityID = EntityID
            ostrEntityName = EntityName
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            Me.Reset()
        End Sub
        Sub New(ByVal dr As DataRow)
            oEntityID = dr.Item("ENTITY_ID")
            ostrEntityName = dr.Item("ENTITY_NAME")
            ostrCreatedBy = dr.Item("CREATED_BY")
            odtCreatedOn = dr.Item("CREATED_DATE")
            ostrModifiedBy = IIf(dr.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, dr.Item("LAST_EDITED_BY"))
            odtModifiedOn = IIf(dr.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_LAST_EDITED"))
            Me.Reset()
        End Sub
        Sub New(ByVal ds As DataSet)
            Try
                LoadEntity(ds)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"

        Public Sub Reset()

            dtCreatedOn = odtCreatedOn
            dtModifiedOn = odtModifiedOn
            EntityID = oEntityID
            strCreatedBy = ostrCreatedBy
            strEntityName = ostrEntityName
            strModifiedBy = ostrModifiedBy
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"

        Private Sub CheckDirty()

            bolIsDirty = (EntityID <> oEntityID) Or _
                         (strEntityName <> ostrEntityName)

        End Sub
        Private Sub Init()

            oEntityID = 0
            odtCreatedOn = System.DateTime.Now
            odtModifiedOn = System.DateTime.Now
            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty
            ostrEntityName = String.Empty
            Me.Reset()

        End Sub

        Private Sub LoadEntity(ByRef ds As DataSet)

            Try
                If ds.Tables.Count > 0 Then
                    Dim oDataTable As DataTable
                    oDataTable = ds.Tables(0)
                    '
                    ' Got too many rows - can only be one row...
                    '
                    If oDataTable.Rows.Count <> 1 Then
                        Throw New Exception("Attempt to get entity  returned " & oDataTable.Rows.Count.ToString & " values!")
                        Exit Sub
                    End If
                    Dim oRow As DataRow
                    oRow = oDataTable.Rows(0)
                    oEntityID = oRow.Item("ENTITY_ID")
                    ostrEntityName = oRow.Item("ENTITY_NAME")
                    ostrCreatedBy = oRow.Item("CREATED_BY")
                    odtCreatedOn = oRow.Item("CREATED_DATE")
                    odtModifiedOn = oRow.Item("DATE_LAST_EDITED")
                    ostrModifiedBy = oRow.Item("LAST_EDITED_BY")
                    Me.Reset()
                End If
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return EntityID
            End Get

            Set(ByVal value As Integer)
                EntityID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property Name() As String
            Get
                Return strEntityName
            End Get

            Set(ByVal value As String)
                strEntityName = value
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

#Region "iAccessors"
        Public ReadOnly Property CreatedBy() As String Implements iAccessors.CreatedBy
            Get
                Return strCreatedBy
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As Date Implements iAccessors.CreatedOn
            Get
                Return dtCreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedBy() As String Implements iAccessors.ModifiedBy
            Get
                Return strModifiedBy
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date Implements iAccessors.ModifiedOn
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
