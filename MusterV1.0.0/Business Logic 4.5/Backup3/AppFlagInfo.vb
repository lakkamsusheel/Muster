'-------------------------------------------------------------------------------
' MUSTER.Info.AppFlagInfo
'   Provides the container to persist MUSTER Template state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1      EN          02/22/2005  Added New Attribute ModuleName
'
' Function          Description
' New()             Instantiates an empty AppFlagInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated AppFlagInfo object
' New(dr)           Instantiates a populated AppFlagInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class AppFlagInfo
        Implements iAccessors
#Region "Public Events"

#End Region
#Region "Private member variables"
        '
        Private ostrKey As String
        Private strKey As String

        Private ostrModuleName As String
        Private strModuleName As String



        'GUID Object for each item.
        Private oObjValue As Object
        Private ObjValue As Object

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal ID As Int64, _
        ByVal Deleted As Boolean)
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            Me.strKey = Me.ostrKey
            Me.ObjValue = Me.oObjValue
            strModuleName = ostrModuleName
        End Sub
        Public Sub Archive()
            Me.ostrKey = Me.strKey
            Me.oObjValue = Me.ObjValue
            ostrModuleName = strModuleName
        End Sub
#End Region
#Region "Private Operations"
        Private Sub Init()
            strKey = String.Empty
            ObjValue = Nothing
            strModuleName = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property Key() As String
            Get
                Return strKey
            End Get
            Set(ByVal Value As String)
                strKey = Value
            End Set
        End Property
        Public Property Value() As Object
            Get
                Return ObjValue
            End Get
            Set(ByVal Value As Object)
                ObjValue = Value
            End Set
        End Property
        Public Property ModuleName() As String
            Get
                Return strModuleName
            End Get
            Set(ByVal Value As String)
                strModuleName = Value
            End Set
        End Property



#Region "iAccessors"
        Public ReadOnly Property CreatedBy() As String Implements iAccessors.CreatedBy
            Get
                'Return strCreatedBy
            End Get
        End Property

        Public ReadOnly Property CreatedOn() As Date Implements iAccessors.CreatedOn
            Get
                'Return dtCreatedOn
            End Get
        End Property

        Public ReadOnly Property ModifiedBy() As String Implements iAccessors.ModifiedBy
            Get
                'Return strModifiedBy
            End Get
        End Property

        Public ReadOnly Property ModifiedOn() As Date Implements iAccessors.ModifiedOn
            Get
                'Return dtModifiedOn
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
