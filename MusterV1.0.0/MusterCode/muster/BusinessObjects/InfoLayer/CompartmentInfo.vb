'-------------------------------------------------------------------------------
' MUSTER.Info.CompartmentInfo
'   Provides the container to persist MUSTER Owner state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN       12/16/04    Original class definition.
'  1.1        AN       12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        AB       02/17/05    Added AgeThreshold and IsAgedData Attributes
'  1.3        MNR      03/15/05    Added Constructor New(ByVal drCompartment As DataRow)
'  1.4        MNR      03/16/05    Removed strSrc from events
'  1.5        KKM      03/18/05    PipesCollection property is added
'
' Function          Description
'' New()             Instantiates an empty compartment object
''New(ByVal Tank As Integer, ByVal CompartmentNumber As Integer, ByVal Capacity As Integer, ByVal CCERCLA As Integer, ByVal Substance As Integer, ByVal FuelTypeId As Integer, ByVal Deleted As Boolean, ByVal CREATED_BY As String, ByVal DATE_CREATED As Date, ByVal LAST_EDITED_BY As String, ByVal DATE_LAST_EDITED As Date)
'                   Instantiates a populated Compartment object
'Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'Archive            Sets the object state to the old state when loaded from or
'                   last saved to the repository
'CheckDirty         'Check for dirty....
'Init               Intialise the object attributes...

'Attribute          Description
' ID                The unique key identifier associated with the compartment number and tankid in the Collection.
' Tankid            The unique identifier associated with the tank in the repository
' COMPARTMENTNumber  The unique identifier associated with the compartment in the repository
' IsDirty           Indicates if the Facility state has been altered since it was
'                       last loaded from or saved to the repository.
'Capacity
'CCERCLA
'Substance
'FuelTypeId
'Deleted
'AgeThreshold         Indicates the number of minutes old data can be before it should be 
'                           refreshed from the DB.  Data should only be refreshed when Retrieved
'                           and when IsDirty is false
'IsAgedData           Will return true if the data has been held longer than the AgeThreshold
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class CompartmentInfo

#Region "Public Events"
        Public Event CompartmentInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"

        Private nTankId As Integer
        Private nCompartmentNumber As Integer
        Private nCapacity As Integer
        Private nCCERCLA As Integer
        Private nSubstance As Integer
        Private nFuelTypeId As Integer
        Private strCreatedBy As String
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private nFacilityId As Integer
        Private nTankSiteID As Integer

        'Current values
        Private onTankId As Integer
        Private onCompartmentNumber As Integer
        Private onCapacity As Integer
        Private onCCERCLA As Integer
        Private onSubstance As Integer
        Private onFuelTypeId As Integer
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private onFacilityId As Integer
        Private onTankSiteID As Integer
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private dtDataAge As DateTime

        Private nAgeThreshold As Int16 = 5
        'added by kiran
        'Dim colPipes As MUSTER.Info.PipesCollection
        'end changes
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
            'added by kiran 
            'colPipes = New MUSTER.Info.PipesCollection
            'end changes
        End Sub
        Public Sub New(ByVal Tank As Integer, ByVal CompartmentNumber As Integer, ByVal Capacity As Integer, ByVal CCERCLA As Integer, ByVal Substance As Integer, ByVal FuelTypeId As Integer, ByVal Deleted As Boolean, ByVal CREATED_BY As String, ByVal DATE_CREATED As Date, ByVal LAST_EDITED_BY As String, ByVal DATE_LAST_EDITED As Date)
            onTankId = Tank
            onCompartmentNumber = CompartmentNumber
            onCapacity = Capacity
            onCCERCLA = CCERCLA
            onSubstance = Substance
            onFuelTypeId = FuelTypeId
            obolDeleted = Deleted
            ostrCreatedBy = CREATED_BY
            odtCreatedOn = DATE_CREATED
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            onFacilityId = 0
            onTankSiteID = 0
            dtDataAge = Now()
            'added by kiran 
            'colPipes = New MUSTER.Info.PipesCollection
            'end changes
            Me.Reset()
        End Sub
        Sub New(ByVal drCompartment As DataRow)
            Try
                Me.Init()
                onTankId = drCompartment.Item("TANK_ID")
                onCompartmentNumber = drCompartment.Item("COMPARTMENT_NUMBER")
                onCapacity = IIf(drCompartment.Item("CAPACITY") Is System.DBNull.Value, 0, drCompartment.Item("CAPACITY"))
                onCCERCLA = IIf(drCompartment.Item("CERCLA#") Is System.DBNull.Value, 0, drCompartment.Item("CERCLA#"))
                onSubstance = IIf(drCompartment.Item("SUBSTANCE") Is System.DBNull.Value, 0, drCompartment.Item("SUBSTANCE"))
                onFuelTypeId = IIf(drCompartment.Item("FUEL_TYPE_ID") Is System.DBNull.Value, 0, drCompartment.Item("FUEL_TYPE_ID"))
                obolDeleted = drCompartment.Item("DELETED")
                ostrCreatedBy = drCompartment.Item("CREATED_BY")
                odtCreatedOn = drCompartment.Item("DATE_CREATED")
                ostrModifiedBy = IIf(drCompartment.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drCompartment.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drCompartment.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drCompartment.Item("DATE_LAST_EDITED"))
                'onFacilityId = 0
                'onTankSiteID = 0
                dtDataAge = Now()
                'colPipes = New MUSTER.Info.PipesCollection
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nCompartmentNumber >= 0 Then
                nCompartmentNumber = onCompartmentNumber
            End If
            If nTankId >= 0 Then
                nTankId = onTankId
            End If
            nCapacity = onCapacity
            nCCERCLA = onCCERCLA
            nSubstance = onSubstance
            nFuelTypeId = onFuelTypeId
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            nFacilityId = onFacilityId
            nTankSiteID = onTankSiteID
            bolIsDirty = False
        End Sub
        Public Sub Archive()
            onTankId = nTankId
            onCompartmentNumber = nCompartmentNumber
            onCapacity = nCapacity
            onCCERCLA = nCCERCLA
            onSubstance = nSubstance
            onFuelTypeId = nFuelTypeId
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            onFacilityId = nFacilityId
            onTankSiteID = nTankSiteID
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            '(onCompartmentNumber <> nCompartmentNumber)
            bolIsDirty = (onTankId <> nTankId) Or _
                            (onCapacity <> nCapacity) Or _
                            (onCCERCLA <> nCCERCLA) Or _
                            (onSubstance <> nSubstance) Or _
                            (onFuelTypeId <> nFuelTypeId) Or _
                            (obolDeleted <> bolDeleted)
            If bolOldState <> bolIsDirty Then
                'MsgBox("Info F:" + FacilityId.ToString + " TI:" + TankSiteID.ToString + " CID:" + ID)
                RaiseEvent CompartmentInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onTankId = 0
            onCompartmentNumber = 0
            onCapacity = 0
            onCCERCLA = 0
            onSubstance = 0
            onFuelTypeId = 0
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            obolIsDirty = False
            onFacilityId = 0
            onTankSiteID = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        'added by kumar on March 11th
        'Public Property pipesCollection() As MUSTER.Info.PipesCollection
        '    Get
        '        Return colPipes
        '    End Get
        '    Set(ByVal Value As MUSTER.Info.PipesCollection)
        '        colPipes = Value
        '    End Set
        'End Property
        'End Changes
        Public Property ID() As String
            Get
                Return CType(nTankId, String) & "|" & CType(nCompartmentNumber, String)
            End Get
            Set(ByVal value As String)
                Dim arrVals() As String
                arrVals = value.Split("|")
                nTankId = Integer.Parse(arrVals(0))
                nCompartmentNumber = Integer.Parse(arrVals(1))
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankId() As Integer
            Get
                Return Me.nTankId
            End Get
            Set(ByVal value As Integer)
                Me.nTankId = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property COMPARTMENTNumber() As Integer
            Get
                Return Me.nCompartmentNumber
            End Get
            Set(ByVal value As Integer)
                Me.nCompartmentNumber = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Capacity() As Integer
            Get
                Return Me.nCapacity
            End Get
            Set(ByVal value As Integer)
                Me.nCapacity = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CCERCLA() As Integer

            Get
                Return Me.nCCERCLA
            End Get
            Set(ByVal value As Integer)
                Me.nCCERCLA = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Substance() As Integer

            Get
                Return Me.nSubstance
            End Get
            Set(ByVal value As Integer)
                Me.nSubstance = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FuelTypeId() As Integer

            Get
                Return Me.nFuelTypeId
            End Get
            Set(ByVal value As Integer)
                Me.nFuelTypeId = value
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
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityId() As Integer
            Get
                Return nFacilityId
            End Get
            Set(ByVal Value As Integer)
                nFacilityId = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankSiteID() As Integer
            Get
                Return nTankSiteID
            End Get
            Set(ByVal Value As Integer)
                nTankSiteID = Value
                Me.CheckDirty()
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
#End Region
#Region "iAccessors"
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
        Public ReadOnly Property ModifiedOn()
            Get
                Return dtModifiedOn
            End Get
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
