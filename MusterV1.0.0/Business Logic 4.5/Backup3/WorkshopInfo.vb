'-------------------------------------------------------------------------------
' MUSTER.Info.WorkshopInfo
'   Provides the container to persist MUSTER Workshop state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1       JC         01/12/05    Added line of code to RESET to raise
'                                       data changed event when called.
'
' Function          Description
' New()             Instantiates an empty WorkshopInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated WorkshopInfo object
' New(dr)           Instantiates a populated WorkshopInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as Workshop to build other objects.
'       Replace keyword "Workshop" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class WorkshopInfo
        Implements iAccessors
#Region "Public Events"
        Public Event WorkshopInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private strName As String
        '********************************************************
        '
        ' Other private member variables for current state here
        '
        '********************************************************
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private ostrName As String
        '********************************************************
        '
        ' Other private member variables for previous state here
        '
        '********************************************************
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        '********************************************************
        '
        ' Update signature to reflect other attributes persisted
        '   by the object.
        '
        '********************************************************
        Sub New(ByVal ID As Int64, _
        ByVal Deleted As Boolean, _
        ByVal CreatedBy As String, _
        ByVal CreatedOn As Date, _
        ByVal ModifiedBy As String, _
        ByVal LastEdited As Date)
            onID = ID
            '********************************************************
            '
            ' Other private member variables for prior state here
            '
            '********************************************************
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            Me.Reset()
        End Sub
        Sub New(ByVal drWorkshop As DataRow)
            Try
                onID = drWorkshop.Item("ID")
                '********************************************************
                '
                ' Other private member variables for prior state here
                '
                '********************************************************
                obolDeleted = drWorkshop.Item("DELETED")
                ostrCreatedBy = drWorkshop.Item("CREATED_BY")
                odtCreatedOn = drWorkshop.Item("DATE_CREATED")
                ostrModifiedBy = drWorkshop.Item("LAST_EDITED_BY")
                odtModifiedOn = drWorkshop.Item("DATE_LAST_EDITED")
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nID = onID
            '********************************************************
            '
            ' Other assignments of current state to prior state here
            '
            '********************************************************
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent WorkshopInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            '********************************************************
            '
            ' Other assignments of prior state to current state here
            '
            '********************************************************
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        '********************************************************
        '
        ' Update eval in CheckDirty with other current and prior
        '    state variables.
        '
        '********************************************************
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nID <> onID) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent WorkshopInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onID = 0
            '********************************************************
            '
            ' Other assignments of current state to empty/false/etc here
            '
            '********************************************************
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Name() As String
            Get
                Return strName
            End Get
            Set(ByVal Value As String)
                strName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get

            Set(ByVal value As Boolean)
                bolDeleted = value
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
        Public Property WORKSHOP_NAME As Integer
Get
End Get
Set
End Set
End Property
Public Property WORKSHOP_ID As Integer
Get
End Get
Set
End Set
End Property
Public Property VIOLATION_ID As Integer
Get
End Get
Set
End Set
End Property
Public Property PASSED As Boolean
Get
End Get
Set
End Set
End Property
Public Property DATE_WORKSHOP_ATTENDED As Date
Get
End Get
Set
End Set
End Property
Public Property DATE_WORKSHOP_ASSIGNED As Date
Get
End Get
Set
End Set
End Property

Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace