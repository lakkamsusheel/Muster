'-------------------------------------------------------------------------------
' MUSTER.Info.ComAddressInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR      5/24/05    Original class definition.
'
' Function          Description
'  New()             Instantiates an empty ComAddressInfo object.
'  New(COM_ADDRESS_ID,COMPANY_ID,LICENSEE_ID,PROVIDER_ID,ADDRESS_LINE_ONE,ADDRESS_LINE_TWO,
'           CITY,STATE,ZIP,FIPS_CODE,PHONE_NUMBER_ONE,EXT_ONE,PHONE_ONE_COMMENT,
'           PHONE_NUMBER_TWO,EXT_TWO,PHONE_TWO_COMMENT,CELL_NUMBER,FAX_NUMBER,
'           CREATED_BY,DATE_CREATED,LAST_EDITED_BY,DATE_LAST_EDITED,DELETED)
'                    Instantiates a populated ComAddressInfo object.
'  New(dr)           Instantiates a populated ComAddressInfo object taking member state
'                       from the datarow provided
'  Archive()         Sets the object state to the new state 
'  Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
'Attribute                      Description
'
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
Public Class ComAddressInfo
        ' Implements iAccessors

#Region "Private member variables"

        'Original Values
        Private onAddressId As Integer = 0
        Private onCompanyId As Integer
        Private onLicenseeId As Integer
        Private onProviderId As Integer
        Private ostrAddressLine1 As String = String.Empty
        Private ostrAddressLine2 As String = String.Empty
        Private ostrState As String = String.Empty
        Private ostrCity As String = String.Empty
        Private ostrZip As String = String.Empty
        Private ostrFIPSCode As String = String.Empty
        Private ostrPhone1 As String = String.Empty
        Private ostrPhone2 As String = String.Empty
        Private ostrExt1 As String = String.Empty
        Private ostrExt2 As String = String.Empty
        Private ostrPhone1Comment As String = String.Empty
        Private ostrPhone2Comment As String = String.Empty
        Private ostrCell As String = String.Empty
        Private ostrFax As String = String.Empty
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime


        'Current Values
        Private nAddressId As Integer = 0
        Private nCompanyId As Integer
        Private nLicenseeId As Integer
        Private nProviderId As Integer
        Private strAddressLine1 As String = String.Empty
        Private strAddressLine2 As String = String.Empty
        Private strState As String = String.Empty
        Private strCity As String = String.Empty
        Private strZip As String = String.Empty
        Private strFIPSCode As String = String.Empty
        Private strPhone1 As String = String.Empty
        Private strPhone2 As String = String.Empty
        Private strExt1 As String = String.Empty
        Private strExt2 As String = String.Empty
        Private strPhone1Comment As String = String.Empty
        Private strPhone2Comment As String = String.Empty
        Private strCell As String = String.Empty
        Private strFax As String = String.Empty
        Private bolDeleted As Boolean

        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

#End Region
#Region "Public Events"
        Public Event AddressInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal AddressId As Integer, _
            ByVal CompanyId As Integer, _
            ByVal LicenseeId As Integer, _
            ByVal ProviderId As Integer, _
            ByVal AddressLineOne As String, _
            ByVal AddressLineTwo As String, _
            ByVal City As String, _
            ByVal State As String, _
            ByVal Zip As String, _
            ByVal FIPSCode As String, _
            ByVal Phone1 As String, _
            ByVal Ext1 As String, _
            ByVal Phone1Comment As String, _
            ByVal Phone2 As String, _
            ByVal Ext2 As String, _
            ByVal Phone2Comment As String, _
            ByVal Cell As String, _
            ByVal Fax As String, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date)

            onAddressId = AddressId
            onCompanyId = CompanyId
            onLicenseeId = LicenseeId
            onProviderId = ProviderId
            ostrAddressLine1 = AddressLineOne
            ostrAddressLine2 = AddressLineTwo
            ostrCity = City
            ostrState = State
            ostrZip = Zip
            ostrFIPSCode = FIPSCode
            ostrPhone1 = Phone1
            ostrPhone2 = Phone2
            ostrExt1 = Ext1
            ostrExt2 = Ext2
            ostrPhone1Comment = Phone1Comment
            ostrPhone2Comment = Phone2Comment
            ostrCell = Cell
            ostrFax = Fax
            obolDeleted = Deleted
            strCreatedBy = CreatedBy
            dtCreatedOn = CreatedOn
            strModifiedBy = ModifiedBy
            dtModifiedOn = LastEdited
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drAddress As DataRow)
            'Try
            '    onAddressId = drAddress.Item("ADDRESS_ID")
            '    'onAddressTypeId = IIf(drAddress.Item("ADDRESS_TYPE_ID") Is System.DBNull.Value, 0, drAddress.Item("ADDRESS_TYPE_ID"))
            '    'onEntityType = IIf(drAddress.Item("ENTITY_TYPE") Is System.DBNull.Value, 0, drAddress.Item("ENTITY_TYPE"))
            '    ostrAddressLine1 = drAddress.Item("ADDRESS_LINE_ONE")
            '    ostrAddressLine2 = IIf(drAddress.Item("ADDRESS_TWO") Is System.DBNull.Value, String.Empty, drAddress.Item("ADDRESS_TWO"))
            '    ostrCity = drAddress.Item("CITY")
            '    ostrState = drAddress.Item("STATE")
            '    ostrZip = drAddress.Item("ZIP")
            '    ostrFIPSCode = drAddress.Item("FIPS_CODE")
            '    'odtStartDate = IIf(drAddress.Item("START_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drAddress.Item("START_DATE"))
            '    'odtEndDate = IIf(drAddress.Item("END_DATE") Is System.DBNull.Value, CDate("01/01/0001"), drAddress.Item("END_DATE"))
            '    'obolDeleted = drAddress.Item("DELETED")
            '    'strCreatedBy = drAddress.Item("CREATED_BY_ADD")
            '    'dtCreatedOn = drAddress.Item("DATE_CREATED_ADD")
            '    'strModifiedBy = IIf(drAddress.Item("LAST_EDITED_BY_ADD") Is System.DBNull.Value, String.Empty, drAddress.Item("LAST_EDITED_BY_ADD"))
            '    'dtModifiedOn = IIf(drAddress.Item("DATE_LAST_EDITED_ADD") Is System.DBNull.Value, CDate("01/01/0001"), drAddress.Item("DATE_LAST_EDITED_ADD"))
            '    dtDataAge = Now()
            '    Me.Reset()
            'Catch ex As Exception
            '    MusterException.Publish(ex, Nothing, Nothing)
            '    Throw ex
            'End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nAddressId >= 0 Then
                nAddressId = onAddressId
            End If
            nCompanyId = onCompanyId
            nLicenseeId = onLicenseeId
            nProviderId = onProviderId
            strAddressLine1 = ostrAddressLine1
            strAddressLine2 = ostrAddressLine2
            strCity = ostrCity
            strState = ostrState
            strZip = ostrZip
            strFIPSCode = ostrFIPSCode
            strPhone1 = ostrPhone1
            strPhone2 = ostrPhone2
            strExt1 = ostrExt1
            strExt2 = ostrExt2
            strPhone1Comment = ostrPhone1Comment
            strPhone2Comment = ostrPhone2Comment
            strCell = ostrCell
            strFax = ostrFax
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
            onCompanyId = nCompanyId
            onLicenseeId = nLicenseeId
            onProviderId = nProviderId
            ostrAddressLine1 = strAddressLine1
            ostrAddressLine2 = strAddressLine2
            ostrCity = strCity
            ostrState = strState
            ostrZip = strZip
            ostrFIPSCode = strFIPSCode
            ostrPhone1 = strPhone1
            ostrPhone2 = strPhone2
            ostrExt1 = strExt1
            ostrExt2 = strExt2
            ostrPhone1Comment = strPhone1Comment
            ostrPhone2Comment = strPhone2Comment
            ostrCell = strCell
            ostrFax = strFax
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
            bolIsDirty = (nCompanyId <> onCompanyId Or _
                        nLicenseeId <> onLicenseeId Or _
                        nProviderId <> onProviderId Or _
                        strAddressLine1 <> ostrAddressLine1 Or _
                        strAddressLine2 <> ostrAddressLine2 Or _
                        strState <> ostrState Or _
                        strCity <> ostrCity Or _
                        strZip <> ostrZip Or _
                        strFIPSCode <> ostrFIPSCode Or _
                        strPhone1 <> ostrPhone1 Or _
                        strPhone2 <> ostrPhone2 Or _
                        strExt1 <> ostrExt1 Or _
                        strExt2 <> ostrExt2 Or _
                        strPhone1Comment <> ostrPhone1Comment Or _
                        strPhone2Comment <> ostrPhone2Comment Or _
                        strCell <> ostrCell Or _
                        strFax <> ostrFax Or _
                        bolDeleted <> obolDeleted)
            If obolIsDirty <> bolIsDirty Then
                RaiseEvent AddressInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onAddressId = 0
            onCompanyId = 0
            onLicenseeId = 0
            onProviderId = 0
            ostrAddressLine1 = String.Empty
            ostrAddressLine2 = String.Empty
            ostrCity = String.Empty
            ostrState = String.Empty
            ostrZip = String.Empty
            ostrFIPSCode = String.Empty
            ostrPhone1 = String.Empty
            ostrPhone2 = String.Empty
            ostrExt1 = String.Empty
            ostrExt2 = String.Empty
            ostrPhone1Comment = String.Empty
            ostrPhone2Comment = String.Empty
            ostrCell = String.Empty
            ostrFax = String.Empty
            obolDeleted = False

            dtCreatedOn = System.DateTime.Now
            dtModifiedOn = System.DateTime.Now
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property AddressId() As Integer
            Get
                Return nAddressId
            End Get

            Set(ByVal value As Integer)
                nAddressId = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property
        Public Property CompanyId() As Integer
            Get
                Return nCompanyId
            End Get

            Set(ByVal value As Integer)
                nCompanyId = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return nLicenseeId
            End Get

            Set(ByVal value As Integer)
                nLicenseeId = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property
        Public Property ProviderID() As Integer
            Get
                Return nProviderId
            End Get

            Set(ByVal value As Integer)
                nProviderId = Integer.Parse(value)
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
        Public Property AddressLine2() As String
            Get
                Return strAddressLine2
            End Get

            Set(ByVal value As String)
                strAddressLine2 = value
                Me.CheckDirty()
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
        Public Property Phone1() As String
            Get
                Return strPhone1
            End Get

            Set(ByVal value As String)
                strPhone1 = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Phone2() As String
            Get
                Return strPhone2
            End Get

            Set(ByVal value As String)
                strPhone2 = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Ext1() As String
            Get
                Return strExt1
            End Get

            Set(ByVal value As String)
                strExt1 = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Ext2() As String
            Get
                Return strExt2
            End Get

            Set(ByVal value As String)
                strExt2 = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Phone1Comment() As String
            Get
                Return strPhone1Comment
            End Get

            Set(ByVal value As String)
                strPhone1Comment = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Phone2Comment() As String
            Get
                Return strPhone2Comment
            End Get

            Set(ByVal value As String)
                strPhone2Comment = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Cell() As String
            Get
                Return strCell
            End Get

            Set(ByVal value As String)
                strCell = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Fax() As String
            Get
                Return strFax
            End Get

            Set(ByVal value As String)
                strFax = value
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
                If strCreatedBy = Nothing Then
                    Return String.Empty
                Else
                    Return strCreatedBy
                End If
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                If strModifiedBy = Nothing Then
                    Return String.Empty
                Else
                    Return strModifiedBy
                End If
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
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


