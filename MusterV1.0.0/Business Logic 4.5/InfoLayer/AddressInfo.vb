'-------------------------------------------------------------------------------
' MUSTER.Info.AddressInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/06/04    Original class definition.
'  1.1        KJ      12/23/04    Added Archive Function and also added descriptions in Header.
'  1.2        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        KJ      12/30/04    Added event for notification of data changes.
'                                 Added firing of event in CHECKDIRTY() if
'                                   dirty state changed.
'  1.4        MNR     01/13/05    Commented - firing of event in CheckDirty() if dirty state changed,
'                                 Added Events
'  1.5        AB      01/31/05    Corrected CheckDirty RaiseEvent
'  1.6        AB      02/17/05    Added AgeThreshold and IsAgedData Attributes
'  1.7        MNR     03/15/05    Updated Constructor New(ByVal drAddress As DataRow) to check for System.DBNull.Value
'  1.8        MNR     03/16/05    Removed strSrc from events
'  1.9        MNR     07/25/05    Added County Property
'  2.0 Thomas Franey  12/02/09    Added Physical Town Property
'
' Function          Description
'  New()             Instantiates an empty AddressInfo object.
'  New(AddressId, nAddressTypeId, nEntityType, strAddressLine1, strAddressLine2, strCity, strState, strZip, strFIPSCode, 
'           strState, dtStartDate, dtEndDate, bolDeleted, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                    Instantiates a populated AddressInfo object.
'  New(dr)           Instantiates a populated AddressInfo object taking member state
'                       from the datarow provided
'  Archive()         Sets the object state to the new state 
'  Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
'Attribute                      Description
'  AddressId            The unique identifier associated with the Address in the repository.
'  AddressTypeID        The AddressType ID 
'  EntityType           The Entity Type associated with AddressInfo
'  AddressLine1         The First Line in the AddressInfo
'  AddressLine2         The second Line in the AddressInfo
'  City                 The City in the AddressInfo
'  State                The State in the AddressInfo
'  Zip                  The Zip Code in the Address
'  FIPSCode             The FIPS-Federal Information Processing .. Code
'  StartDate            Start Date at the Address
'  EndDate              The End Date at the Address
'  Deleted              The Boolean Flag indicating if the address has been deleted
'  IsDirty              Indicates if the Address state has been altered since it was
'                           last loaded from or saved to the repository.
'  AgeThreshold         Indicates the number of minutes old data can be before it should be 
'                           refreshed from the DB.  Data should only be refreshed when Retrieved
'                           and when IsDirty is false
'  IsAgedData           Will return true if the data has been held longer than the AgeThreshold
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
Public Class AddressInfo

#Region "Private member variables"

        'Original Values
        Private onAddressId As Integer
        Private onAddressTypeId As Integer
        Private onEntityType As Integer
        Private ostrAddressLine1 As String
        Private ostrAddressLine1ForEnsite As String
        Private ostrAddressLine2 As String
        Private ostrAddressLine2ForEnsite As String
        Private ostrState As String
        Private ostrCity As String
        Private ostrPhysicalTown As String

        Private ostrZip As String
        Private ostrFIPSCode As String
        Private ostrCounty As String
        Private odtStartDate As DateTime
        Private odtEndDate As DateTime
        Private obolDeleted As Boolean

        'Current Values
        Private nAddressId As Integer
        Private nAddressTypeId As Integer
        Private nEntityType As Integer
        Private strAddressLine1 As String
        Private strAddressLine1ForEnsite As String
        Private strAddressLine2 As String
        Private strAddressLine2ForEnsite As String
        Private strState As String
        Private strCity As String
        Private strPhysicalTown As String
        Private strZip As String
        Private strFIPSCode As String
        Private strCounty As String
        Private dtStartDate As DateTime
        Private dtEndDate As DateTime
        Private bolDeleted As Boolean

        Private strCreatedBy As String
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString

        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private dtDataAge As DateTime

        Private nAgeThreshold As Int16 = 5

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        'Public Event AddressChanged(ByVal bolValue As Boolean)
        Public Event AddressInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Init()
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal AddressId As Integer, _
            ByVal AddressTypeId As Integer, _
            ByVal EntityType As Integer, _
            ByVal AddressLineOne As String, _
            ByVal AddressTwo As String, _
            ByVal City As String, _
            ByVal State As String, _
            ByVal Zip As String, _
            ByVal FIPSCode As String, _
            ByVal StartDate As Date, _
            ByVal EndDate As Date, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date, _
            ByVal county As String, _
            ByVal AddressLineOneForEnsite As String, _
            ByVal AddressTwoForEnsite As String, _
            Optional ByVal PhysicalTown As String = "")
            onAddressId = AddressId
            onAddressTypeId = AddressTypeId
            onEntityType = EntityType
            ostrAddressLine1 = AddressLineOne.Trim
            ostrAddressLine1ForEnsite = AddressLineOneForEnsite.Trim
            ostrAddressLine2 = AddressTwo.Trim
            ostrAddressLine2ForEnsite = AddressTwoForEnsite.Trim
            ostrCity = City.Trim
            ostrState = State.Trim
            ostrZip = Zip.Trim
            ostrFIPSCode = FIPSCode.Trim
            odtStartDate = StartDate
            odtEndDate = EndDate
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            ostrCounty = county.Trim
            ostrPhysicalTown = PhysicalTown
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drAddress As DataRow)
            Try
                onAddressId = drAddress.Item("ADDRESS_ID")
                onAddressTypeId = IIf(drAddress.Item("ADDRESS_TYPE_ID") Is System.DBNull.Value, 0, drAddress.Item("ADDRESS_TYPE_ID"))
                onEntityType = IIf(drAddress.Item("ENTITY_TYPE") Is System.DBNull.Value, 0, drAddress.Item("ENTITY_TYPE"))
                ostrAddressLine1 = drAddress.Item("ADDRESS_LINE_ONE")
                ostrAddressLine1 = ostrAddressLine1.Trim
                ostrAddressLine1ForEnsite = IIf(drAddress.Item("ADDRESS_LINE_ONE_FOR_ENSITE") Is System.DBNull.Value, ostrAddressLine1, drAddress.Item("ADDRESS_LINE_ONE_FOR_ENSITE"))
                ostrAddressLine1ForEnsite = ostrAddressLine1ForEnsite.Trim
                ostrAddressLine2 = IIf(drAddress.Item("ADDRESS_TWO") Is System.DBNull.Value, String.Empty, drAddress.Item("ADDRESS_TWO"))
                ostrAddressLine2 = ostrAddressLine2.Trim
                ostrAddressLine2ForEnsite = IIf(drAddress.Item("ADDRESS_TWO_FOR_ENSITE") Is System.DBNull.Value, ostrAddressLine2, drAddress.Item("ADDRESS_TWO_FOR_ENSITE"))
                ostrAddressLine2ForEnsite = ostrAddressLine2ForEnsite.Trim
                ostrCity = drAddress.Item("CITY")
                ostrCity = ostrCity.Trim
                ostrState = drAddress.Item("STATE")
                ostrState = ostrState.Trim
                ostrZip = drAddress.Item("ZIP")
                ostrZip = ostrZip.Trim
                ostrFIPSCode = drAddress.Item("FIPS_CODE")
                ostrFIPSCode = ostrFIPSCode.Trim
                odtStartDate = IIf(drAddress.Item("START_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drAddress.Item("START_DATE"))
                odtEndDate = IIf(drAddress.Item("END_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drAddress.Item("END_DATE"))
                obolDeleted = drAddress.Item("DELETED")
                ostrCreatedBy = drAddress.Item("CREATED_BY_ADD")
                odtCreatedOn = drAddress.Item("DATE_CREATED_ADD")
                ostrModifiedBy = IIf(drAddress.Item("LAST_EDITED_BY_ADD") Is System.DBNull.Value, String.Empty, drAddress.Item("LAST_EDITED_BY_ADD"))
                odtModifiedOn = IIf(drAddress.Item("DATE_LAST_EDITED_ADD") Is System.DBNull.Value, CDate("01/01/0001"), drAddress.Item("DATE_LAST_EDITED_ADD"))
                ostrCounty = IIf(drAddress.Item("COUNTY") Is DBNull.Value, String.Empty, drAddress.Item("COUNTY"))
                ostrPhysicalTown = IIf(drAddress.Item("PHYSICALTOWN") Is DBNull.Value, String.Empty, drAddress.Item("PHYSICALTOWN"))
                ostrCounty = ostrCounty.Trim
                dtDataAge = Now()
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nAddressId >= 0 Then
                nAddressId = onAddressId
            End If
            nAddressTypeId = onAddressTypeId
            nEntityType = onEntityType
            strAddressLine1 = ostrAddressLine1
            strAddressLine1ForEnsite = ostrAddressLine1ForEnsite
            strAddressLine2 = ostrAddressLine2
            strAddressLine2ForEnsite = ostrAddressLine2ForEnsite
            strState = ostrState
            strCity = ostrCity
            strPhysicalTown = ostrPhysicalTown
            strZip = ostrZip
            strFIPSCode = ostrFIPSCode
            strCounty = ostrCounty
            dtStartDate = odtStartDate
            dtEndDate = odtEndDate
            bolDeleted = obolDeleted

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolIsDirty = False
            RaiseEvent AddressInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onAddressId = nAddressId
            onAddressTypeId = nAddressTypeId
            onEntityType = nEntityType
            ostrAddressLine1 = strAddressLine1
            ostrAddressLine1ForEnsite = strAddressLine1ForEnsite
            ostrAddressLine2 = strAddressLine2
            ostrAddressLine2ForEnsite = strAddressLine2ForEnsite
            ostrPhysicalTown = strPhysicalTown

            ostrState = strState
            ostrCity = strCity
            ostrZip = strZip
            ostrFIPSCode = strFIPSCode
            ostrCounty = strCounty
            odtStartDate = dtStartDate
            odtEndDate = dtEndDate
            obolDeleted = bolDeleted

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            bolIsDirty = (nAddressTypeId <> onAddressTypeId Or _
                        nEntityType <> onEntityType Or _
                        strAddressLine1 <> ostrAddressLine1 Or _
                        strAddressLine1ForEnsite <> ostrAddressLine1ForEnsite Or _
                        strAddressLine2 <> ostrAddressLine2 Or _
                        strAddressLine2ForEnsite <> ostrAddressLine2ForEnsite Or _
                        strState <> ostrState Or _
                        strCity <> ostrCity Or _
                        strZip <> ostrZip Or _
                        strFIPSCode <> ostrFIPSCode Or _
                        strCounty <> ostrCounty Or _
                        dtStartDate <> odtStartDate Or _
                        dtEndDate <> odtEndDate Or _
                        strPhysicalTown.ToUpper <> ostrPhysicalTown.ToUpper Or _
                        bolDeleted <> obolDeleted)
            If obolIsDirty <> bolIsDirty Then
                RaiseEvent AddressInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onAddressId = 0
            onAddressTypeId = 0
            onEntityType = 0
            ostrAddressLine1 = String.Empty
            ostrAddressLine1ForEnsite = String.Empty
            ostrAddressLine2 = String.Empty
            ostrAddressLine2ForEnsite = String.Empty
            ostrCity = String.Empty
            ostrState = String.Empty
            ostrZip = String.Empty
            ostrFIPSCode = String.Empty
            ostrCounty = String.Empty
            odtStartDate = CDate("01/01/0001")
            odtEndDate = CDate("01/01/0001")
            obolDeleted = False
            odtCreatedOn = CDate("01/01/0001")
            odtModifiedOn = CDate("01/01/0001")
            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty
            ostrPhysicalTown = String.Empty

            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property PhsycalTown() As String
            Get
                Return strPhysicalTown
            End Get
            Set(ByVal Value As String)
                strPhysicalTown = Value
            End Set
        End Property

        Public Property AddressId() As Integer
            Get
                Return nAddressId
            End Get

            Set(ByVal value As Integer)
                nAddressId = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property AddressTypeId() As Integer
            Get
                Return nAddressTypeId
            End Get

            Set(ByVal value As Integer)
                nAddressTypeId = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property EntityType() As Integer
            Get
                Return nEntityType
            End Get

            Set(ByVal value As Integer)
                nEntityType = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property AddressLine1() As String
            Get
                Return strAddressLine1
            End Get

            Set(ByVal value As String)
                strAddressLine1 = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AddressLine1ForEnsite() As String
            Get
                Return strAddressLine1ForEnsite
            End Get
            Set(ByVal Value As String)
                strAddressLine1ForEnsite = Value
                CheckDirty()
            End Set
        End Property

        Public Property AddressLine2() As String
            Get
                Return strAddressLine2
            End Get

            Set(ByVal value As String)
                strAddressLine2 = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AddressLine2ForEnsite() As String
            Get
                Return strAddressLine2ForEnsite
            End Get
            Set(ByVal Value As String)
                strAddressLine2ForEnsite = Value
                CheckDirty()
            End Set
        End Property

        Public Property City() As String
            Get
                Return strCity
            End Get

            Set(ByVal value As String)
                strCity = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property State() As String
            Get
                Return strState
            End Get

            Set(ByVal value As String)
                strState = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Zip() As String
            Get
                Return strZip
            End Get

            Set(ByVal value As String)
                strZip = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FIPSCode() As String
            Get
                Return strFIPSCode
            End Get

            Set(ByVal value As String)
                strFIPSCode = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property County() As String
            Get
                Return strCounty
            End Get
            Set(ByVal Value As String)
                strCounty = Value
                Me.CheckDirty()
            End Set
        End Property

        Public WriteOnly Property CountyFirstTime() As String
            Set(ByVal Value As String)
                strCounty = Value
                ostrCounty = Value
            End Set
        End Property

        Public Property StartDate() As DateTime
            Get
                Return dtStartDate
            End Get

            Set(ByVal value As DateTime)
                dtStartDate = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EndDate() As DateTime
            Get
                Return dtEndDate
            End Get

            Set(ByVal value As DateTime)
                dtEndDate = value
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
