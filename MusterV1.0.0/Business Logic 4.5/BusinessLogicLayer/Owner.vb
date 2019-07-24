'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Owner
'   Provides the operations required to manipulate an Owner object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0        MNR      12/03/04    Original class definition.
'   1.1        AN       12/16/04    Added Address Object
'   1.2        EN       12/29/04    Changed from GetAllByOwnerID to GetALL.
'   1.3        AN       01/03/05    Added Try catch and Exception Handling/Logging
'   1.4        MNR      01/05/05    Added set value for colIsDirty function
'   1.5        KJ       01/07/05    Changed the GetAddress to Retrive as I have removed that function.
'   1.6        EN       01/10/05    Added BPersona Property.
'   1.7        MNR      01/11/05    Added GetFacilities function, 
'                                   Added Events,
'                                   Changed Retrieve function to handle hirearchy,
'                                   Commented unreferenced private member variables
'   1.8        MNR      01/14/05    Added ValidateData(), commented oFacility - unreferrenced variable,
'                                   modified flush()
'   1.9        MNR      01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   2.0        MNR      01/26/05    Added ValidateEmail(..) ValidatePhone(..) Functions
'   2.1        MNR      01/27/05    Added remove from collection after deleted flag is set to true
'   2.2        MNR      01/28/05    Implemented ChangeKey functionality
'   2.3        AN       02/02/05    Added Comments object
'   2.4        JVC2     02/08/05    Exposed info created and modified attributes
'   2.5        EN       02/10/05    Added Events too handle facilityCAP STATUS.
'   2.6        AB       02/16/05    Streamlined the Retrieve Function
'   2.7        AB       02/21/05    Added DataAge check to the Retrieve function
'   2.8        MR       02/22/05    Added PopulateOwnerName() to Retrieve OwnerNames.
'   2.9        MNR      03/15/05    Added RetrieveAll Sub - Speed Loading of Objects
'   3.0        MNR      03/16/05    Removed strSrc from events
'   3.1        AB       03/21/05    Added FacilitiesLUSTSummaryTable / GetFacilitiesLUSTSummary
'   3.2        KKM      03/18/05    Events for handling local FacilityCollection and CommentsCollection are added
'   3.3        MNR      07/25/05    handling events for address changed
'
' Function          Description
' Retrieve(ID)      Returns an Info Object requested by the int arg ID
' Save()            Saves the Info Object
' GetAll()          Returns a collection with all the relevant information
' Add(ID)           Adds an Info Object identified by the int arg ID
'                   to the Owners Collection
' Add(Entity)       Adds the Entity passed as an argument
'                   to the Owners Collection
' Remove(ID)        Removes an Info Object identified by the int arg ID
'                   from the Owners Collection
' Remove(Entity)    Removes the Entity passed as an argument
'                   from the Owners Collection
' Flush()           Marshalls all modified/new Onwer Info objects in the
'                   Owner Collection to the repository
' ListOwnerIDs(showdeleted) Returns a datatable with a list of all Owner ID's
' EntityTable()     Returns a datatable containing all columns for the Entity
'                   objects in the Owners Collection
' EntityCombo()     Returns a two-column datatable containing Owner ID and Org ID for
'                   the Entity objects in the Owners Collection
'-------------------------------------------------------------------------------
'
' TODO - Integrate with solution 2/9/05 JVC2

Namespace MUSTER.BusinessLogic
    <Serializable()> _
        Public Class pOwner
#Region "Public Events"
        Public Event FlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer)
        Public Event evtOwnerErr(ByVal MsgStr As String)
        Public Event evtOwnerChanged(ByVal bolValue As Boolean)
        Public Event evtOwnersChanged(ByVal bolValue As Boolean)
        Public Event evtValidationErr(ByVal ID As Integer, ByVal MsgStr As String)
        Public Event evtOwnerCommentsChanged(ByVal bolValue As Boolean)

        'facility
        Public Event evtFacilityChanged(ByVal bolValue As Boolean)
        Public Event evtFacilitiesChanged(ByVal bolValue As Boolean)
        Public Event evtFacilityCommentsChanged(ByVal bolValue As Boolean)
        'Added By Elango 
        Public Event evtOwnFacilityCAPStatusChanged(ByVal BolValue As Boolean, ByVal facID As Integer)

        'address
        Public Event evtAddressChanged(ByVal bolValue As Boolean)
        Public Event evtAddressesChanged(ByVal bolValue As Boolean)

        'persona
        Public Event evtPersonaChanged(ByVal bolValue As Boolean)
        Public Event evtPersonasChanged(ByVal bolValue As Boolean)


        'Tank
        'Public Event evtTankCommentsChanged(ByVal bolValue As Boolean)
        'Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String)

        'Pipe
        'Public Event evtPipeCommentsChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private strBankruptChapter As String
        Private dtBankruptDate As Date
        Private strBP2KOwnerID As String
        Private WithEvents colOwners As MUSTER.Info.OwnersCollection
        Private WithEvents oOwnerInfo As MUSTER.Info.OwnerInfo
        Private oOwnerDB As MUSTER.DataAccess.OwnerDB
        Private WithEvents oOwnFacility As MUSTER.BusinessLogic.pFacility
        Private WithEvents oOwnAddress As MUSTER.BusinessLogic.pAddress
        Private WithEvents oOwnPersona As MUSTER.BusinessLogic.pPersona
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private nID As Int64 = -1
        Private MusterException As MUSTER.Exceptions.MusterExceptions
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Owner").ID
        Private dtFacTnkTable As DataTable
        Private dtFacLUSTTable As DataTable
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oOwnerInfo = New MUSTER.Info.OwnerInfo
            colOwners = New MUSTER.Info.OwnersCollection
            oOwnerDB = New MUSTER.DataAccess.OwnerDB(strDBConn, MusterXCEP)
            oOwnFacility = New MUSTER.BusinessLogic.pFacility(strDBConn, MusterXCEP, oOwnerInfo)
            oOwnAddress = New MUSTER.BusinessLogic.pAddress
            oOwnPersona = New MUSTER.BusinessLogic.pPersona
            oComments = New MUSTER.BusinessLogic.pComments
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property OwnerInfo() As MUSTER.Info.OwnerInfo
            Get
                Return oOwnerInfo
            End Get
        End Property
        Public Property CreatedBy() As String
            Get
                Return oOwnerInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oOwnerInfo.CreatedBy = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oOwnerInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oOwnerInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oOwnerInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oOwnerInfo.ModifiedOn
            End Get
        End Property
        Public ReadOnly Property ID() As Integer
            Get
                Return oOwnerInfo.ID
            End Get

            'Set(ByVal value As Integer)
            '    oOwnerInfo.ID = Integer.Parse(value)
            'End Set
        End Property

        Public Property OrganizationID() As Integer
            Get
                Return oOwnerInfo.OrganizationID
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.OrganizationID = Integer.Parse(value)
            End Set
        End Property

        Public Property PersonID() As Integer
            Get
                Return oOwnerInfo.PersonID
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.PersonID = Integer.Parse(value)
            End Set
        End Property

        Public Property PhoneNumberOne() As String
            Get
                Return oOwnerInfo.PhoneNumberOne
            End Get

            Set(ByVal value As String)
                oOwnerInfo.PhoneNumberOne = value
            End Set
        End Property

        Public Property PhoneNumberTwo() As String
            Get
                Return oOwnerInfo.PhoneNumberTwo
            End Get

            Set(ByVal value As String)
                oOwnerInfo.PhoneNumberTwo = value
            End Set
        End Property

        Public Property Fax() As String
            Get
                Return oOwnerInfo.Fax
            End Get

            Set(ByVal value As String)
                oOwnerInfo.Fax = value
            End Set
        End Property

        Public Property EmailAddress() As String
            Get
                Return oOwnerInfo.EmailAddress
            End Get

            Set(ByVal value As String)
                oOwnerInfo.EmailAddress = value
            End Set
        End Property

        Public Property EmailAddressPersonal() As String
            Get
                Return oOwnerInfo.EmailAddressPersonal
            End Get

            Set(ByVal value As String)
                oOwnerInfo.EmailAddressPersonal = value
            End Set
        End Property

        Public Property AddressId() As Long
            Get
                Return oOwnerInfo.AddressId
            End Get

            Set(ByVal value As Long)
                oOwnerInfo.AddressId = Long.Parse(value)
            End Set
        End Property

        'Public Property AddressLine1() As String
        '    Get
        '        Return oOwnerInfo.AddressLine1
        '    End Get

        '    Set(ByVal value As String)
        '        oOwnerInfo.AddressLine1 = value
        '    End Set
        'End Property

        'Public Property AddressLine2() As String
        '    Get
        '        Return oOwnerInfo.AddressLine2
        '    End Get

        '    Set(ByVal value As String)
        '        oOwnerInfo.AddressLine2 = value
        '    End Set
        'End Property

        'Public Property City() As String
        '    Get
        '        Return oOwnerInfo.City
        '    End Get

        '    Set(ByVal value As String)
        '        oOwnerInfo.City = value
        '    End Set
        'End Property

        'Public Property State() As String
        '    Get
        '        Return oOwnerInfo.State
        '    End Get

        '    Set(ByVal value As String)
        '        oOwnerInfo.State = value
        '    End Set
        'End Property

        'Public Property Zip() As String
        '    Get
        '        Return oOwnerInfo.Zip
        '    End Get

        '    Set(ByVal value As String)
        '        oOwnerInfo.Zip = value
        '    End Set
        'End Property

        'Public Property FIPSCode() As String
        '    Get
        '        Return oOwnerInfo.FIPSCode
        '    End Get

        '    Set(ByVal value As String)
        '        oOwnerInfo.FIPSCode = value
        '    End Set
        'End Property

        Public Property DateCapSignUp() As Date
            Get
                Return oOwnerInfo.DateCapSignUp
            End Get

            Set(ByVal value As Date)
                oOwnerInfo.DateCapSignUp = value
            End Set
        End Property

        Public Property CapCurrentStatus() As Boolean
            Get
                Return oOwnerInfo.CapCurrentStatus
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.CapCurrentStatus = value
            End Set
        End Property

        Public Property OwnerType() As Integer
            Get
                Return oOwnerInfo.OwnerType
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.OwnerType = Integer.Parse(value)
            End Set
        End Property

        Public Property BP2KType() As Integer
            Get
                Return oOwnerInfo.BP2KType
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.BP2KType = Integer.Parse(value)
            End Set
        End Property

        Public Property FeesProfileID() As Integer
            Get
                Return oOwnerInfo.FeesProfileID
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.FeesProfileID = Integer.Parse(value)
            End Set
        End Property

        Public Property FeesStatus() As Boolean
            Get
                Return oOwnerInfo.FeesStatus
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.FeesStatus = value
            End Set
        End Property

        Public Property ComplianceProfileID() As Integer
            Get
                Return oOwnerInfo.ComplianceProfileID
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.ComplianceProfileID = Integer.Parse(value)
            End Set
        End Property

        Public Property ComplianceStatus() As Boolean
            Get
                Return oOwnerInfo.ComplianceStatus
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.ComplianceStatus = value
            End Set
        End Property

        Public Property Active() As Boolean
            Get
                Return oOwnerInfo.Active
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.Active = value
            End Set
        End Property

        Public Property FeeActive() As Boolean
            Get
                Return oOwnerInfo.FeeActive
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.FeeActive = value
            End Set
        End Property

        Public Property EnsiteOrganizationID() As Integer
            Get
                Return oOwnerInfo.EnsiteOrganizationID
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.EnsiteOrganizationID = Integer.Parse(value)
            End Set
        End Property

        Public ReadOnly Property EnsiteOrganization() As MUSTER.Info.PersonaInfo
            Get
                Try
                    Return oOwnPersona.Retrieve(oOwnerInfo.EnsiteOrganizationID)
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property

        Public Property EnsitePersonID() As Integer
            Get
                Return oOwnerInfo.EnsitePersonID
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.EnsitePersonID = Integer.Parse(value)
            End Set
        End Property
        'strBankruptChapter = ""
        'dtBankruptDate = tmpDate
        'strBP2KOwnerID = ""

        'Public ReadOnly Property BP2KOwnerID() As String
        '    Get
        '        Try
        '            Return strBP2KOwnerID
        '        Catch Ex As Exception
        '            If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '            Throw Ex
        '        End Try
        '    End Get
        'End Property
        Public ReadOnly Property BankruptChapter() As String
            Get
                Try
                    Return strBankruptChapter
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property
        Public ReadOnly Property BankruptDate() As Date
            Get
                Try
                    Return dtBankruptDate
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property
        Public ReadOnly Property EnsitePerson() As MUSTER.Info.PersonaInfo
            Get
                Try
                    Return oOwnPersona.Retrieve(oOwnerInfo.EnsitePersonID)
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property

        Public Property EnsiteAgencyInterestID() As Boolean
            Get
                Return oOwnerInfo.EnsiteAgencyInterestID
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.EnsiteAgencyInterestID = value
            End Set
        End Property

        '  Public Property Description() As Integer
        '     Get
        '        Return oOwnerInfo.Description
        '   End Get

        '   Set(ByVal value As Integer)
        '      oOwnerInfo.Description = Integer.Parse(value)
        '   End Set
        '  End Property

        Public Property CustEntityCode() As Integer
            Get
                Return oOwnerInfo.CustEntityCode
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.CustEntityCode = Integer.Parse(value)
            End Set
        End Property

        Public Property CustTypeCode() As Integer
            Get
                Return oOwnerInfo.CustTypeCode
            End Get

            Set(ByVal value As Integer)
                oOwnerInfo.CustTypeCode = Integer.Parse(value)
            End Set
        End Property
        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property

        Public Property Deleted() As Boolean
            Get
                Return oOwnerInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.Deleted = value
            End Set
        End Property

        Public Property OwnerL2CSnippet() As Boolean
            Get
                Return oOwnerInfo.OwnerL2CSnippet
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.OwnerL2CSnippet = value
            End Set
        End Property

        Public ReadOnly Property BP2KOwnerID() As String
            Get
                Return oOwnerInfo.BP2KOwnerID
            End Get
        End Property

        Public Property CAPParticipationLevel() As String
            Get
                Return oOwnerInfo.CapParticipationLevel
            End Get
            Set(ByVal Value As String)
                oOwnerInfo.CapParticipationLevel = Value
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oOwnerInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oOwnerInfo.IsDirty = value
            End Set
        End Property

        Public ReadOnly Property colIsDirty(Optional ByVal self As Boolean = True) As Boolean
            Get
                If self Then
                    If oOwnerInfo.IsDirty Or oOwnFacility.colIsDirty Or oOwnPersona.colIsDirty Then
                        Return True
                        Exit Property
                    End If
                Else
                    For Each oownerInfoLocal As MUSTER.Info.OwnerInfo In colOwners.Values
                        If oownerInfoLocal.IsDirty Then
                            Return True
                            Exit Property
                        End If
                        For Each facInfo As MUSTER.Info.FacilityInfo In oownerInfoLocal.facilityCollection.Values
                            If facInfo.IsDirty Then
                                Return True
                                Exit Property
                            End If
                            For Each tankInfo As MUSTER.Info.TankInfo In facInfo.TankCollection.Values
                                If tankInfo.IsDirty Then
                                    Return True
                                    Exit Property
                                End If
                                For Each compInfo As MUSTER.Info.CompartmentInfo In tankInfo.CompartmentCollection.Values
                                    If compInfo.IsDirty Then
                                        Return True
                                        Exit Property
                                    End If
                                Next
                                For Each pipeInfo As MUSTER.Info.PipeInfo In tankInfo.pipesCollection.Values
                                    If pipeInfo.IsDirty Then
                                        Return True
                                        Exit Property
                                    End If
                                Next
                            Next
                        Next
                    Next
                End If
                Return False
            End Get
        End Property

        Public ReadOnly Property Organization() As MUSTER.Info.PersonaInfo
            Get
                'modified By elango on Dec  27 2004 
                Try
                    Return oOwnPersona.Retrieve("O" & "|" & oOwnerInfo.OrganizationID)
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property

        Public ReadOnly Property Address() As MUSTER.Info.AddressInfo
            Get
                Try
                    If oOwnerInfo.AddressId > 0 Then
                        Return oOwnAddress.Retrieve(oOwnerInfo.AddressId)
                    Else
                        Return oOwnAddress.Retrieve(oOwnAddress.AddressId)
                    End If

                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property

        Public Property Addresses() As MUSTER.BusinessLogic.pAddress
            Get
                Return oOwnAddress
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pAddress)
                If Not Value Is Nothing AndAlso Value.AddressId > 0 Then
                    oOwnAddress = Value
                End If

            End Set
        End Property

        Public ReadOnly Property Persona() As MUSTER.Info.PersonaInfo
            Get
                Try
                    Return oOwnPersona.Retrieve("P" & "|" & oOwnerInfo.PersonID)
                Catch Ex As Exception
                    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
            End Get
        End Property

        Public Property BPersona() As MUSTER.BusinessLogic.pPersona
            Get
                Return oOwnPersona
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pPersona)
                oOwnPersona = Value
            End Set
        End Property

        Public Property Facilities() As MUSTER.BusinessLogic.pFacility
            Get
                Return Me.oOwnFacility
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pFacility)
                Me.oOwnFacility = Value
            End Set
        End Property

        Public Property Facility() As MUSTER.Info.FacilityInfo
            Get
                Return oOwnFacility.Facility
            End Get
            Set(ByVal Value As MUSTER.Info.FacilityInfo)
                oOwnFacility.Facility = Value
            End Set
        End Property

        Public Property Comments() As MUSTER.BusinessLogic.pComments
            Get
                Return oComments
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pComments)
                oComments = Value
            End Set
        End Property

        Public ReadOnly Property FacilitiesFinancialSummaryTable() As DataTable
            Get
                Return GetFacilitiesFinancialSummaryTable()
            End Get
        End Property

        Public ReadOnly Property FacilitiesLUSTSummaryTable() As DataTable
            Get
                Return GetFacilitiesLUSTSummary()
            End Get
        End Property

        Public ReadOnly Property FacilitiesTankStatusTable() As DataTable
            Get
                Return GetFacilitiesTankStatus()
            End Get
        End Property

        Public ReadOnly Property OwnerCollection() As MUSTER.Info.OwnersCollection
            Get
                Return colOwners
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Sub RetrieveAll(ByVal ownerID As Integer, Optional ByVal [Module] As String = "", Optional ByVal showDeleted As Boolean = False, Optional ByVal facID As Int64 = 0, Optional ByVal tankID As Int64 = 0)
            Try
                Dim ds As New DataSet
                ds = oOwnerDB.DBGetDS(ownerID, [Module], showDeleted, facID, tankID)
                ' 0 - Owner+address
                ' 1 - Person / Organization
                ' 2 - Facilities + addresses
                ' 3 - Tanks
                ' 4 - Compartments
                ' 5 - Pipes
                ' 6 - Compartments_Pipes
                ds.Tables(0).TableName = "Owner"
                ds.Tables(1).TableName = "OrgPerson"
                ds.Tables(2).TableName = "Facilities"
                ds.Tables(3).TableName = "Tanks"
                ds.Tables(4).TableName = "Compartments"
                ds.Tables(5).TableName = "Pipes"
                ' Owner
                If ds.Tables("Owner").Rows.Count > 0 Then
                    'oComments.Clear()
                    oOwnerInfo = New MUSTER.Info.OwnerInfo(ds.Tables("Owner").Rows(0))
                    colOwners.Add(oOwnerInfo)
                    ' Owner Address
                    oOwnAddress.Load(ds.Tables("Owner").Rows(0))
                    ds.Tables.Remove("Owner")

                    ' Person / Organization
                    If ds.Tables("OrgPerson").Rows.Count > 0 Then
                        If ds.Tables("OrgPerson").Columns(0).ColumnName <> "PERSON_ID" Then
                            oOwnPersona.Load(ds, "O")
                        Else
                            oOwnPersona.Load(ds, "P")
                        End If
                    Else
                        oOwnPersona.Retrieve("P|0")
                    End If
                    ds.Tables.Remove("OrgPerson")

                    ' Facility, Tank, Comp, Pipe
                    oOwnFacility.Load(oOwnerInfo, ds, [Module])
                    LoadFeesSpecific(ownerID)

                    'Comments
                    'oComments.Load(ds, [Module], nEntityTypeID, oOwnerInfo.ID)
                    'oComments.GetByModule("", nEntityTypeID, oOwnerInfo.ID)
                    ds = Nothing
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByVal ID As Integer, Optional ByVal strDepth As String = "SELF", _
                                Optional ByVal showDeleted As Boolean = False, _
                                Optional ByVal bolLoading As Boolean = False, _
                                Optional ByVal intInspectionID As Integer = Integer.MinValue) As MUSTER.Info.OwnerInfo
            Dim bolDataAged As Boolean = False
            Try
                Dim strCallDepth As String
                Dim bolValidDepth As Boolean

                If Not (oOwnerInfo.Deleted Or oOwnerInfo.ID = 0 Or Not oOwnerInfo.IsDirty Or bolLoading) Then
                    Me.ValidateData()
                End If

                Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
                Select Case UCase(strDepth).Trim
                    Case "SELF", "CHILD"
                        bolValidDepth = True
                        strCallDepth = "SELF"

                    Case "GRANDCHILD", "ALL"
                        bolValidDepth = True
                        strCallDepth = "CHILD"

                    Case Else
                        bolValidDepth = False
                End Select
                If bolValidDepth Then
                    oOwnerInfo = colOwners.Item(ID)

                    ' Check for Aged Data here.
                    If Not (oOwnerInfo Is Nothing) Then
                        If oOwnerInfo.IsAgedData = True And oOwnerInfo.IsDirty = False Then
                            bolDataAged = True
                            colOwners.Remove(oOwnerInfo)
                        End If
                    End If

                    If bolDataAged Then
                        RetrieveAll(ID, , showDeleted, , )
                    ElseIf oOwnerInfo Is Nothing Then
                        Add(ID, showDeleted, intInspectionID)
                    End If
                    'If oOwnerInfo Is Nothing Then Or bolDataAged Then
                    '    Add(ID, showDeleted)
                    'End If

                    If UCase(strDepth).Trim <> "SELF" AndAlso oOwnFacility.ID > 0 Then
                        oOwnFacility.Retrieve(oOwnerInfo, oOwnerInfo.ID, strCallDepth, , showDeleted, bolLoading)

                        'Dim bolOwnerActive As Boolean = GetOwnerFacilityStatus()
                        'If oOwnerInfo.Active <> bolOwnerActive Then
                        '    oOwnerInfo.Active = bolOwnerActive
                        '    oOwnerDB.PutOwnerActive(oOwnerInfo.Active, oOwnerInfo.ID)
                        'End If
                    End If
                    If oOwnerInfo.AddressId > 0 Then
                        oOwnAddress.Retrieve(oOwnerInfo.AddressId, strCallDepth, showDeleted, bolLoading)
                    End If

                    If Me.OrganizationID = 0 And Me.PersonID <> 0 Then
                        oOwnPersona.Retrieve("P" & "|" & oOwnerInfo.PersonID, showDeleted)
                    ElseIf Me.OrganizationID <> 0 And Me.PersonID = 0 Then
                        oOwnPersona.Retrieve("O" & "|" & oOwnerInfo.OrganizationID, showDeleted)
                    Else
                        oOwnPersona.Retrieve("P|0")
                    End If
                Else
                ' if input is invalid, raise error
                RaiseEvent evtOwnerErr("Pass correct param for Owner retrieve")
                End If
                'oComments.Clear()
                'oComments.GetByModule("", nEntityTypeID, oOwnerInfo.ID)

                ' no not need to calculate since the value in the db is correct
                'Dim bolOwnerActive As Boolean = GetOwnerFacilityStatus()
                'If oOwnerInfo.Active <> bolOwnerActive Then
                '    oOwnerInfo.ActiveOriginal = bolOwnerActive
                '    oOwnerInfo.Active = bolOwnerActive
                '    oOwnerDB.PutOwnerActive(oOwnerInfo.Active, oOwnerInfo.ID)
                'End If

                UpdateOwnerActive()

                SetInfoInChild()
                LoadFeesSpecific(ID)
                Return oOwnerInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated And Not oOwnerInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oOwnerInfo.ID < 0 And oOwnerInfo.Deleted) Then
                    oldID = oOwnerInfo.ID
                    'oOwnerInfo.Active = GetOwnerFacilityStatus()
                    If oOwnerInfo.IsDirty Then
                        oOwnerDB.Put(oOwnerInfo, moduleID, staffID, returnVal, strUser)
                        If Not returnVal = String.Empty Then
                            Exit Function
                        End If
                    End If

                    If Not bolValidated Then
                        If oldID < 0 Then
                            colOwners.ChangeKey(oldID, oOwnerInfo.ID)
                        End If
                        'oOwnFacility.Flush(strUser, strModule)
                        'oOwnAddress.Flush()
                        'oOwnPersona.Flush()
                        'oComments.Flush()
                    End If

                    SetInfoInChild()
                    oOwnFacility.Flush(moduleID, staffID, returnVal, strUser)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    oOwnPersona.Flush(moduleID, staffID, returnVal)
                    oOwnerInfo.Archive()
                    oOwnerInfo.IsDirty = False
                End If
                If Not bolValidated And bolDelete Then
                    If oOwnerInfo.Deleted Then
                        ' check if other owners are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oOwnerInfo.ID Then
                            If strPrev = oOwnerInfo.ID Then
                                RaiseEvent evtOwnerErr("Owner " + oOwnerInfo.ID.ToString + " deleted")
                                colOwners.Remove(oOwnerInfo)
                                If bolDelete Then
                                    oOwnerInfo = New MUSTER.Info.OwnerInfo
                                Else
                                    oOwnerInfo = Me.Retrieve(0)
                                End If
                            Else
                                RaiseEvent evtOwnerErr("Owner " + oOwnerInfo.ID.ToString + " deleted")
                                colOwners.Remove(oOwnerInfo)
                                oOwnerInfo = Me.Retrieve(strPrev)
                            End If
                        Else
                            RaiseEvent evtOwnerErr("Owner " + oOwnerInfo.ID.ToString + " deleted")
                            colOwners.Remove(oOwnerInfo)
                            oOwnerInfo = Me.Retrieve(strNext)
                        End If
                    End If
                End If
                RaiseEvent evtOwnerChanged(oOwnerInfo.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
                Return False
            End Try

        End Function
        Private Function GetFacilitiesFinancialSummaryTable() As DataTable
            Try
                dtFacLUSTTable = oOwnerDB.GetFacilitiesFinancialSummaryTable(oOwnerInfo.ID).Tables(0)
                dtFacLUSTTable.Columns("FAC_ID").ColumnName = "FacilityID"
                dtFacLUSTTable.Columns("NAME").ColumnName = "Facility Name"
                dtFacLUSTTable.Columns("ADDRESS").ColumnName = "Address"
                dtFacLUSTTable.Columns("CITY").ColumnName = "City"
                dtFacLUSTTable.Columns("COUNTY").ColumnName = "County"
                dtFacLUSTTable.Columns("OpenEvents").ColumnName = "Open Events"
                dtFacLUSTTable.Columns("ClosedEvents").ColumnName = "Closed Events"
                dtFacLUSTTable.Columns("TotalEvents").ColumnName = "Total"

                Return dtFacLUSTTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function GetFacilitiesLUSTSummary() As DataTable
            Try
                dtFacLUSTTable = oOwnerDB.GetFacilitiesLUSTSummary(oOwnerInfo.ID).Tables(0)
                dtFacLUSTTable.Columns("FAC_ID").ColumnName = "FacilityID"
                dtFacLUSTTable.Columns("NAME").ColumnName = "Facility Name"
                dtFacLUSTTable.Columns("ADDRESS").ColumnName = "Address"
                dtFacLUSTTable.Columns("CITY").ColumnName = "City"
                dtFacLUSTTable.Columns("COUNTY").ColumnName = "County"
                dtFacLUSTTable.Columns("OpenEvents").ColumnName = "Open Events"
                dtFacLUSTTable.Columns("ClosedEvents").ColumnName = "Closed Events"
                dtFacLUSTTable.Columns("TotalEvents").ColumnName = "Total"

                Return dtFacLUSTTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function GetFacilitiesTankStatus() As DataTable
            Try
                dtFacTnkTable = oOwnerDB.GetFacilitiesTankStatus(oOwnerInfo.ID).Tables(0)
                dtFacTnkTable.Columns("FAC_ID").ColumnName = "FacilityID"
                dtFacTnkTable.Columns("NAME").ColumnName = "Facility Name"
                dtFacTnkTable.Columns("TOU").ColumnName = "TOS"
                dtFacTnkTable.Columns("TOT").ColumnName = "Total"
                Dim col As New DataColumn
                Dim colData As New Collection
                Dim i As Integer
                col = dtFacTnkTable.Columns("Total")
                For i = 0 To dtFacTnkTable.Rows.Count - 1
                    colData.Add(CType(dtFacTnkTable.Rows(i).Item("Total"), Int32))
                Next
                dtFacTnkTable.Columns.Remove(col)
                dtFacTnkTable.Columns.Add(col)
                For i = 0 To dtFacTnkTable.Rows.Count - 1
                    dtFacTnkTable.Rows(i).Item("Total") = CType(colData.Item(i + 1), Int32)
                Next
                Return dtFacTnkTable
                'If dtFacTnkTable Is Nothing Then ' Or dtFacTnkTable.Rows.Count < 1 Then
                '    dtFacTnkTable = oOwnerDB.GetFacilitiesTankStatus(oOwnerInfo.ID).Tables(0)
                '    dtFacTnkTable.Columns("TOT").ColumnName = "Total"
                '    dtFacTnkTable.Columns("FAC_ID").ColumnName = "FacilityID"
                '    dtFacTnkTable.Columns("NAME").ColumnName = "Facility Name"
                '    dtFacTnkTable.Columns("TOU").ColumnName = "TOS"
                '    Dim col As New DataColumn
                '    Dim colData As New Collection
                '    Dim i As Integer
                '    col = dtFacTnkTable.Columns("Total")
                '    For i = 0 To dtFacTnkTable.Rows.Count - 1
                '        colData.Add(CType(dtFacTnkTable.Rows(i).Item("Total"), Int32))
                '    Next
                '    dtFacTnkTable.Columns.Remove(col)
                '    dtFacTnkTable.Columns.Add(col)
                '    For i = 0 To dtFacTnkTable.Rows.Count - 1
                '        dtFacTnkTable.Rows(i).Item("Total") = CType(colData.Item(i + 1), Int32)
                '    Next
                'Else
                '    ' build data table from collections ?
                'End If
                'Return dtFacTnkTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetFacilitiesCAESummaryTable() As DataSet
            Dim strSQL As String = String.Empty
            Try
                strSQL = "SELECT * FROM dbo.v_CAE_FACILITY_DISPLAY_DATA WHERE OWNER_ID = " + oOwnerInfo.ID.ToString
                Return oOwnerDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            ' to be completed according to DDD specs for registration / technical ... - Manju 01/14/05
            Try
                Dim errStr As String = ""
                Dim validateSuccess As Boolean = True
                Select Case [module]
                    Case "Registration"
                        ' User adds new owner by providing atleast the foll info
                        ' name, add, owner category, org_entity_code (if owner category is org)
                        ' owner type(private, commercial, local govt, etc)
                        If oOwnerInfo.ID <> 0 Then
                            If oOwnerInfo.AddressId <= 0 Then
                                errStr += "Owner Address is required" + vbCrLf
                                validateSuccess = False
                            End If
                            If oOwnerInfo.EmailAddress.Trim <> String.Empty Then
                                If Not ValidateEmail(oOwnerInfo.EmailAddress) Then
                                    errStr += "Owner Email Address Validation Failed" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If
                            If oOwnerInfo.PhoneNumberOne.Trim <> String.Empty Then
                                If Not ValidatePhone(oOwnerInfo.PhoneNumberOne) Then
                                    errStr += "Owner Phone Validation Failed" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If
                            If oOwnerInfo.PhoneNumberTwo.Trim <> String.Empty Then
                                If Not ValidatePhone(oOwnerInfo.PhoneNumberTwo) Then
                                    errStr += "Owner Phone 2 Validation Failed" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If
                            If oOwnerInfo.Fax.Trim <> String.Empty Then
                                If Not ValidatePhone(oOwnerInfo.Fax) Then
                                    errStr += "Owner Fax Validation Failed" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If
                            If Not validateSuccess Then
                                Exit Select
                            End If
                            If oOwnerInfo.OwnerType <> 0 Then
                                ''If oOwnFacility.ValidateData() Then
                                If oOwnPersona.ValidateData() Then
                                    validateSuccess = True
                                    'If oOwnAddress.ValidateData() Then
                                    '    validateSuccess = True
                                    'Else
                                    '    errStr += "Owner Address Validations Failed" + vbCrLf
                                    '    validateSuccess = False
                                    'End If
                                Else
                                    'errStr += "Owner Persona Validations Failed" + vbCrLf
                                    'validateSuccess = False
                                    Return False
                                End If
                                ''Else
                                ''errStr += "Owner Facility Validations Failed" + vbCrLf
                                ''validateSuccess = False
                                ''End If
                            Else
                                errStr += "Owner Type cannot be empty" + vbCrLf
                                validateSuccess = False
                            End If
                        End If
                        Exit Select
                        'Case "Technical"
                End Select
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent evtOwnerErr(errStr)
                    Exit Function
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function DeleteOwner(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String) As Boolean
            Try
                'Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
                'For Each oFacilityInfoLocal In oOwnFacility.FacilityCollection.Values
                '    If oFacilityInfoLocal.OwnerID = oOwnerInfo.ID Then
                '        RaiseEvent evtOwnerErr("The Specified owner has associated Facilities. Delete Facilities before deleting the owner")
                '        Return False
                '    End If
                'Next
                If oOwnerInfo.ID > 0 Then
                    Dim ds As DataSet
                    ds = oOwnerDB.DBGetDS("EXEC spCheckDependancy " + oOwnerInfo.ID.ToString + ",NULL,NULL,0,NULL")
                    ds = oOwnerDB.DBGetDS("SELECT dbo.udfCheckIfFeeExists(" + oOwnerInfo.ID.ToString + ",NULL,NULL)")
                    If ds.Tables(0).Rows(0)(0) Then
                        RaiseEvent evtOwnerErr(IIf(ds.Tables(0).Rows(0)("MSG") Is DBNull.Value, "Owner has dependants", ds.Tables(0).Rows(0)("MSG")))
                        Return False
                    End If
                End If

                oOwnerInfo.Deleted = True
                Return Me.Save(moduleID, staffID, returnVal, "", True, True)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
        Public Sub SetInfoInChild()
            If Not oOwnFacility Is Nothing Then
                oOwnFacility.OwnerInfo = oOwnerInfo
            End If
        End Sub
        Public Sub UpdateOwnerActive()
            Try
                Dim ds As DataSet = oOwnerDB.DBGetDS("select active from tblreg_owner where owner_id = " + oOwnerInfo.ID.ToString)
                If Not ds Is Nothing Then
                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            oOwnerInfo.ActiveOriginal = ds.Tables(0).Rows(0).Item("active")
                            oOwnerInfo.Active = ds.Tables(0).Rows(0).Item("active")
                        End If
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub GetCAPParticipationLevel()
            Try
                Dim ds As DataSet = oOwnerDB.DBGetDS("select cap_participation_level from tblreg_owner where owner_id = " + oOwnerInfo.ID.ToString)
                If Not ds Is Nothing Then
                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            If ds.Tables(0).Rows(0)(0) Is DBNull.Value Then
                                oOwnerInfo.CapParticipationLevel = "NONE (0/0)"
                            Else
                                oOwnerInfo.CapParticipationLevel = ds.Tables(0).Rows(0)(0)
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.OwnersCollection
            Try
                colOwners.Clear()
                colOwners = oOwnerDB.GetAllInfo
                Return colOwners
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer, Optional ByVal showDeleted As Boolean = False, Optional ByVal intInspectionID As Integer = Integer.MinValue)
            Try
                oOwnerInfo = oOwnerDB.DBGetByID(ID, showDeleted, intInspectionID)
                If oOwnerInfo.ID = 0 Then
                    oOwnerInfo.ID = nID
                    nID -= 1
                End If
                colOwners.Add(oOwnerInfo)
                oOwnFacility.OwnerInfo = oOwnerInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oOwner As MUSTER.Info.OwnerInfo)
            Try
                oOwnerInfo = oOwner
                colOwners.Add(oOwnerInfo)
                oOwnFacility.OwnerInfo = oOwnerInfo
                oOwnFacility.GetAllInfo(oOwnerInfo.ID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Int64)
            Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
            Try
                oOwnerInfoLocal = colOwners.Item(ID)
                If Not (oOwnerInfoLocal Is Nothing) Then
                    colOwners.Remove(oOwnerInfoLocal)
                    oOwnerInfo = New MUSTER.Info.OwnerInfo
                    oOwnFacility = New MUSTER.BusinessLogic.pFacility(, , oOwnerInfo)
                    oOwnAddress = New MUSTER.BusinessLogic.pAddress
                    oOwnPersona = New MUSTER.BusinessLogic.pPersona
                    oComments = New MUSTER.BusinessLogic.pComments
                    Exit Sub
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Owner " & ID.ToString & " is not in the collection of Owners.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oOwnrInf As MUSTER.Info.OwnerInfo)
            Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
            Try
                oOwnerInfoLocal = colOwners.Item(oOwnrInf.ID)
                If Not (oOwnerInfoLocal Is Nothing) Then
                    colOwners.Remove(oOwnrInf)
                    oOwnerInfo = New MUSTER.Info.OwnerInfo
                    oOwnFacility = New MUSTER.BusinessLogic.pFacility(, , oOwnerInfo)
                    oOwnAddress = New MUSTER.BusinessLogic.pAddress
                    oOwnPersona = New MUSTER.BusinessLogic.pPersona
                    oComments = New MUSTER.BusinessLogic.pComments
                    Exit Sub
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Owner " & oOwnrInf.ID & " is not in the collection of Owners.")
        End Sub
        Public Sub RemoveAll()
            Try
                colOwners.Clear()
                oOwnerInfo = New MUSTER.Info.OwnerInfo
                oOwnFacility = New MUSTER.BusinessLogic.pFacility(, , oOwnerInfo)
                oOwnAddress = New MUSTER.BusinessLogic.pAddress
                oOwnPersona = New MUSTER.BusinessLogic.pPersona
                oComments = New MUSTER.BusinessLogic.pComments
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String)
            Try
                Dim IDs As New Collection
                Dim delIDs As New Collection
                Dim index As Integer
                Dim xOwnerInfo As MUSTER.Info.OwnerInfo
                For Each xOwnerInfo In colOwners.Values
                    If xOwnerInfo.IsDirty Then
                        oOwnerInfo = xOwnerInfo
                        If oOwnerInfo.Deleted Then
                            If oOwnerInfo.ID < 0 Then
                                delIDs.Add(oOwnerInfo.ID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, strUser, True)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Me.ValidateData Then
                                If oOwnerInfo.ID < 0 Then
                                    IDs.Add(oOwnerInfo.ID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, strUser, True)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            Else : Exit For
                            End If
                        End If
                        'If Me.ValidateData() Then
                        '    If oOwnerInfo.ID < 0 And _
                        '        Not oOwnerInfo.Deleted Then
                        '        IDs.Add(oOwnerInfo.ID)
                        '    End If
                        '    Me.Save(strUser, strModule, True)
                        'Else : Exit For
                        'End If
                    ElseIf xOwnerInfo.ID > 0 And xOwnerInfo.ChildrenDirty Then
                        oOwnerInfo = xOwnerInfo
                        SetInfoInChild()
                        oOwnFacility.Flush(moduleID, staffID, returnVal, strUser)
                        If Not returnVal = String.Empty Then
                            Exit Sub
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        xOwnerInfo = colOwners.Item(CType(delIDs.Item(index), String))
                        colOwners.Remove(xOwnerInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xOwnerInfo = colOwners.Item(colKey)
                        colOwners.ChangeKey(colKey, xOwnerInfo.ID)
                    Next
                End If
                'oOwnFacility.Flush(strUser, strModule)
                'oOwnAddress.Flush()
                'oOwnPersona.Flush()
                'oComments.Flush()
                RaiseEvent evtOwnersChanged(oOwnerInfo.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colOwners.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colOwners.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            ElseIf colOwners.Count <> 0 Then
                Return colOwners.Item(nArr.GetValue(colIndex)).ID.ToString
            Else
                Return "0"
            End If
        End Function
        Public Function GetOwnerFacilityStatus() As Boolean
            Return oOwnFacility.GetOwnerFacilityStatus(oOwnerInfo)
        End Function
        Public Sub RemoveOwner(ByVal ID As Int64)
            Try
                oOwnerInfo = colOwners.Item(ID)
                If Not (oOwnerInfo Is Nothing) Then
                    Dim strNext As String = Me.GetNext()
                    Dim strPrev As String = Me.GetPrevious()
                    If strNext = oOwnerInfo.ID Then
                        If strPrev = oOwnerInfo.ID Then
                            'colOwners.Remove(oOwnerInfo)
                            oOwnerInfo = New MUSTER.Info.OwnerInfo
                        Else
                            'colOwners.Remove(oOwnerInfo)
                            oOwnerInfo = Me.Retrieve(strPrev, , , True)
                        End If
                    Else
                        'colOwners.Remove(oOwnerInfo)
                        oOwnerInfo = Me.Retrieve(strNext, , , True)
                    End If
                    Exit Sub
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Throw New Exception("Owner " & ID.ToString & " is not in the collection of Owners.")
        End Sub
#End Region
#Region "General Operations"
        Public Function ListOwnerIDs(ByVal bolshowdeleted As Boolean) As DataTable

            Dim dtOwnerIDs As DataTable

            Dim strSQL As String
            Dim dsset As New DataSet
            strSQL = "SELECT '' as OWNER_ID UNION SELECT  OWNER_ID FROM tblREG_OWNER "
            strSQL += IIf(Not bolshowdeleted, " WHERE ACTIVE = 1", "")
            Try
                dsset = oOwnerDB.DBGetDS(strSQL)
                If dsset.Tables(0).Rows.Count > 0 Then
                    dtOwnerIDs = dsset.Tables(0)
                Else
                    dtOwnerIDs = Nothing
                End If
                Return dtOwnerIDs
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Clear(Optional ByVal strDepth As String = "ALL")
            oOwnerInfo = New MUSTER.Info.OwnerInfo
            'oOwnAddress.Clear(strDepth)
            oOwnPersona.Clear(strDepth)
            'If strDepth = "ALL" Then
            '    oOwnFacility = New MUSTER.BusinessLogic.pFacility
            'End If
        End Sub
        Public Sub Reset(Optional ByVal strDepth As String = "ALL")
            oOwnerInfo.Reset()
            'oOwnAddress.Reset(strDepth)
            oOwnPersona.Reset(strDepth)
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the entities in the collection
        Public Function EntityTable() As DataTable

            Dim oOwnerInfoLocal As New MUSTER.Info.OwnerInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Entity ID")
                tbEntityTable.Columns.Add("Organization ID")
                tbEntityTable.Columns.Add("Person ID")
                tbEntityTable.Columns.Add("Phone Number One")
                tbEntityTable.Columns.Add("Phone Number Two")
                tbEntityTable.Columns.Add("Fax Number")
                tbEntityTable.Columns.Add("Email")
                tbEntityTable.Columns.Add("Personal Email")
                'tbEntityTable.Columns.Add("Address ID")
                tbEntityTable.Columns.Add("ADDRESS_LINE_ONE")
                tbEntityTable.Columns.Add("ADDRESS_TWO")
                tbEntityTable.Columns.Add("CITY")
                tbEntityTable.Columns.Add("STATE")
                tbEntityTable.Columns.Add("ZIP")
                tbEntityTable.Columns.Add("FIPS_CODE")
                tbEntityTable.Columns.Add("Date CAP Signup")
                tbEntityTable.Columns.Add("CAP Current Status")
                tbEntityTable.Columns.Add("Owner Type")
                tbEntityTable.Columns.Add("BP2K Owner Type")
                tbEntityTable.Columns.Add("Fees Profile ID")
                tbEntityTable.Columns.Add("Fees Status")
                tbEntityTable.Columns.Add("Compliance Profile ID")
                tbEntityTable.Columns.Add("Compliace Status")
                tbEntityTable.Columns.Add("Active", Type.GetType("System.Boolean"))
                tbEntityTable.Columns.Add("Fee Active")
                tbEntityTable.Columns.Add("Ensite Organization ID")
                tbEntityTable.Columns.Add("Ensite Person ID")
                tbEntityTable.Columns.Add("Ensite Agency Interest ID")
                tbEntityTable.Columns.Add("Owner Description")
                tbEntityTable.Columns.Add("Cust Entity Code")
                tbEntityTable.Columns.Add("Cust Type Code")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Created On")
                tbEntityTable.Columns.Add("Modified By")
                tbEntityTable.Columns.Add("Modified On")
                tbEntityTable.Columns.Add("Owner L2C Snippet")

                For Each oOwnerInfoLocal In colOwners.Values
                    dr = tbEntityTable.NewRow()
                    dr("Entity ID") = oOwnerInfoLocal.ID
                    dr("Organization ID") = oOwnerInfoLocal.OrganizationID
                    dr("Person ID") = oOwnerInfoLocal.PersonID
                    dr("Phone Number One") = oOwnerInfoLocal.PhoneNumberOne
                    dr("Phone Number Two") = oOwnerInfoLocal.PhoneNumberTwo
                    dr("Fax Number") = oOwnerInfoLocal.Fax
                    dr("Email") = oOwnerInfoLocal.EmailAddress
                    dr("Personal Email") = oOwnerInfoLocal.EmailAddressPersonal
                    'dr("Address ID") = oOwnerInfoLocal.AddressId
                    'dr("ADDRESS_LINE_ONE") = oOwnerInfoLocal.AddressLine1
                    'dr("ADDRESS_TWO") = oOwnerInfoLocal.AddressLine2
                    'dr("CITY") = oOwnerInfoLocal.City
                    'dr("STATE") = oOwnerInfoLocal.State
                    'dr("ZIP") = oOwnerInfoLocal.Zip
                    'dr("FIPS_CODE") = oOwnerInfoLocal.FIPSCode
                    dr("Date CAP Signup") = oOwnerInfoLocal.DateCapSignUp.Date
                    dr("CAP Current Status") = oOwnerInfoLocal.CapCurrentStatus
                    dr("Owner Type") = oOwnerInfoLocal.OwnerType
                    dr("BP2K Owner Type") = oOwnerInfoLocal.BP2KType
                    dr("Fees Profile ID") = oOwnerInfoLocal.FeesProfileID
                    dr("Fees Status") = oOwnerInfoLocal.FeesStatus
                    dr("Compliance Profile ID") = oOwnerInfoLocal.ComplianceProfileID
                    dr("Compliace Status") = oOwnerInfoLocal.ComplianceStatus
                    dr("Active") = oOwnerInfoLocal.Active
                    dr("Fee Active") = oOwnerInfoLocal.FeeActive
                    dr("Ensite Organization ID") = oOwnerInfoLocal.EnsiteOrganizationID
                    dr("Ensite Person ID") = oOwnerInfoLocal.EnsitePersonID
                    dr("Ensite Agency Interest ID") = oOwnerInfoLocal.EnsiteAgencyInterestID
                    dr("Cust Entity Code") = oOwnerInfoLocal.CustEntityCode
                    dr("Cust Type Code") = oOwnerInfoLocal.CustTypeCode
                    dr("Deleted") = oOwnerInfoLocal.Deleted
                    dr("Created By") = oOwnerInfoLocal.CreatedBy
                    dr("Created On") = oOwnerInfoLocal.CreatedOn
                    dr("Modified By") = oOwnerInfoLocal.ModifiedBy
                    dr("Modified On") = oOwnerInfoLocal.ModifiedOn
                    dr("Owner L2C Snippet") = oOwnerInfoLocal.OwnerL2CSnippet
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Returns a two-column datatable of the entities in the collection column names Entity ID and Organization ID
        Public Function OwnerCombo() As DataTable

            Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable

            Try
                tbEntityTable.Columns.Add("Owner ID")
                tbEntityTable.Columns.Add("Organization ID")

                For Each oOwnerInfoLocal In colOwners.Values
                    dr = tbEntityTable.NewRow()
                    dr("Owner ID") = oOwnerInfoLocal.ID
                    dr("Organization ID") = oOwnerInfoLocal.OrganizationID
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Private Function ValidateEmail(ByVal strEmail As String) As Boolean
            Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
            Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
            If rx.IsMatch(strEmail) Then
                Return True
            Else
                Return False
            End If
        End Function
        Private Function ValidatePhone(ByVal strPhone As String) As Boolean
            Dim strRegex As String = "(\(\d\d\d\))?\s*(\d\d\d)\s*[\-]?\s*(\d\d\d\d)"
            '"^\(?\d{3}\)?\s|-\d{3}-\d{4}$" -  matches (555) 555-5555, or 555-555-5555
            Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
            If rx.IsMatch(strPhone) Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Sub CheckWriteAccessForRegistration(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oOwnerDB.CheckRegistrationActivityRights(moduleID, staffID, returnVal)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Public Function GetOwnerSummaryLustSites() As DataSet
        '    Dim dsOwnerSummaryLustSites As DataSet
        '    Dim strSQL As String
        '    Try
        '        strSQL = "select * from vOWNERSUMMARY_LustSites where OWNER_ID =" + oOwnerInfo.ID.ToString
        '        dsOwnerSummaryLustSites = oOwnerDB.DBGetDS(strSQL)
        '        Return dsOwnerSummaryLustSites
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function

        Public Function GetOwnerSummary() As DataSet
            Dim dsOwnerSummary As DataSet
            Try
                dsOwnerSummary = oOwnerDB.DBGetOwnerSummary(oOwnerInfo.ID)
                Return dsOwnerSummary
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetOwnerSummaryFeesTotals() As DataTable
            Dim dsFeesTotals As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try
                '                Prior(Balance)
                '                Current(Fees)
                '                Late(Fees)
                '                Total(Due)
                '                Current(Payments)
                '                Current(Credits)
                '                Current(Adjustments)
                'To Date Balance

                strSQL = "Select OWNER_ID, Convert(varchar,PriorBalanceTotal,1) as PriorBalanceTotal, Convert(varchar,CurrentFeesTotal,1) as CurrentFeesTotal,Convert(varchar,LateFeesTotal,1) as LateFeesTotal, "
                strSQL &= "Convert(varchar,TotalDueTotal,1) as TotalDueTotal, Convert(varchar,CurrentPaymentsTotal,1) as CurrentPaymentsTotal,  "
                strSQL &= "Convert(varchar,CurrentCreditsTotal,1) as CurrentCreditTotal, Convert(varchar,CurrentAdjustmentsTotal,1) as CurrentAdjustmentsTotal,Convert(varchar,ToDateBalanceTotal,1) as ToDateBalanceTotal from  "
                strSQL &= "(select OWNER_ID, sum(cast([PriorBalance] as money)) as PriorBalanceTotal, sum(cast([CurrentFees] as money)) as CurrentFeesTotal, sum(cast([LateFees] as money)) as LateFeesTotal, sum(cast([TotalDue] as money))  as TotalDueTotal ,sum(cast([CurrentPayments] as money))  as CurrentPaymentsTotal,sum(cast([CurrentCredits] as money))  as CurrentCreditsTotal,sum(cast([CurrentAdjustments] as money))  as CurrentAdjustmentsTotal,sum(cast([ToDateBalance] as money))  as ToDateBalanceTotal "
                strSQL &= "  from vFees_FacilitySummaryGrid "

                strSQL &= "where OWNER_ID = " & oOwnerInfo.ID & " "
                strSQL &= "group by OWNER_ID) as tmpView "

                dsFeesTotals = oOwnerDB.DBGetDS(strSQL)

                If dsFeesTotals.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsFeesTotals.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub RollOverTankCAPDates(ByVal facID As Integer, ByVal tnkID As Integer, ByVal dtCPDate As Date, ByVal dtLIInspected As Date, ByVal dtTTDate As Date, ByVal dtSpillTested As Date, ByVal dtOverfillInspected As Date, ByVal dtTankSecondary As Date, ByVal dtTankElectronic As Date, ByVal dtATG As Date, ByVal userID As String)
            Try
                oOwnerDB.DBRollOverTankCAPDates(facID, tnkID, dtCPDate, dtLIInspected, dtTTDate, dtSpillTested, dtOverfillInspected, dtTankSecondary, dtTankElectronic, dtATG, userID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub RollOverPipeCAPDates(ByVal facID As Integer, ByVal pipeID As Integer, ByVal dtCPDate As Date, ByVal dtTermCPTestDate As Date, ByVal dtALLDTestDate As Date, ByVal dtTTDate As Date, ByVal dtShear As Date, ByVal dtPipeSecondary As Date, ByVal dtPipeElectronic As Date, ByVal userID As String)
            Try
                oOwnerDB.DBRollOverPipeCAPDates(facID, pipeID, dtCPDate, dtTermCPTestDate, dtALLDTestDate, dtTTDate, dtShear, dtPipeSecondary, dtPipeElectronic, userID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub TransferOwnerBilling(ByVal facilities As String, ByVal oldOwner As Integer, ByVal newOwner As Integer)
            Try
                oOwnerDB.DBTransferOwnerBillingByFacilities(facilities, oldOwner, newOwner)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub


        Public Function CheckWriteAccess(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal entityType As Integer) As Boolean
            Try
                Return oOwnerDB.DBCheckWriteAccess(moduleID, staffID, entityType)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Sub SaveCAPAnnualSummary(ByVal processingYear As Integer, ByVal linePosition As Integer, ByVal ownerID As Integer, _
                                            ByVal ownerName As String, ByVal facilityID As Integer, ByVal facName As String, _
                                            ByVal facAddr1 As String, ByVal facAddrCity As String, ByVal facAddrState As String, _
                                            ByVal facAddrZip As String, ByVal desc As String, ByVal isDescPeriodicTestReq As Boolean, _
                                            ByVal isDescHeading As Boolean, ByVal isDescSubHeading As Boolean, ByVal createdBy As String, ByVal mode As Integer)
            Try
                oOwnerDB.DBSaveCAPAnnualSummary(processingYear, linePosition, ownerID, ownerName, facilityID, facName, _
                                                facAddr1, facAddrCity, facAddrState, facAddrZip, desc, isDescPeriodicTestReq, _
                                                isDescHeading, isDescSubHeading, createdBy, mode)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub SaveCAPAnnualCalendar(ByVal processingYear As Integer, ByVal ownerID As Integer, ByVal ownerName As String, _
                                         ByVal month As Integer, ByVal facilityID As Integer, ByVal facName As String, _
                                        ByVal City As String, ByVal requirements As String, ByVal createdBy As String, _
                                        Optional ByVal mode As Integer = 0, Optional ByVal procMonth As Integer = -1)
            Try
                oOwnerDB.DBSaveCAPAnnualCalendar(processingYear, ownerID, ownerName, month, facilityID, facName, _
                                                 City, requirements, createdBy, mode, procMonth)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region

#Region "Look Up Operations"
        Public Function PopulateOwnerType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vOWNERTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateOpenEntitiesForOwnership(ByVal showAll As Boolean) As DataTable
            Try
                Dim dtReturn As DataSet = Me.RunSQLQuery(String.Format("exec spEntitiesForOwnerShip {0}", IIf(showAll, 1, 0)))
                If Not dtReturn Is Nothing AndAlso dtReturn.Tables.Count > 0 Then
                    Return dtReturn.Tables(0)
                Else
                    Return Nothing
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateOwnerName(ByVal nOwnerID As Int64) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("v_OWNER_NAME_SEARCH", nOwnerID, False)
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateOwnerNameAndOwnerID(ByVal nOwnerID As Int64) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("v_OWNER_NAMEID_SEARCH", nOwnerID, False)
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateOwnerFacilities(ByVal nOwnerID As Int64, Optional ByVal SortByID As Boolean = False) As DataTable
            Dim tablename As String

            Try
                tablename = "vOWNER_FACILITIES_LIST"
                If SortByID And nOwnerID = 0 Then
                    tablename &= " Order By Facility_ID "
                End If
                Dim dtReturn As DataTable = GetDataTable(tablename, nOwnerID, True)
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetDataTable(ByVal DBViewName As String, Optional ByVal nOwnerID As Int64 = 0, Optional ByVal bInclusive As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                If nOwnerID > 0 Then
                    If bInclusive Then
                        strSQL = "SELECT * FROM " & DBViewName & " Where O_ID = " + nOwnerID.ToString + " Order by Facility_Name"
                    Else
                        strSQL = "SELECT * FROM " & DBViewName & " Where O_ID <> " + nOwnerID.ToString + " Order by O_NAME"
                    End If
                Else
                    strSQL = "SELECT * FROM " & DBViewName
                End If

                dsReturn = oOwnerDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                Else
                    dtReturn = Nothing
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PreviousFacilities() As DataSet
            Try
                Return oOwnerDB.DBGetPreviousFacs(oOwnerInfo.ID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RunSQLQuery(ByVal strSQL As String) As DataSet
            Try
                Return oOwnerDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Sub ClearCAPAnnualSummary(ByVal processingYear As Integer, ByVal mode As Integer, ByVal ownerID As Integer, Optional ByVal fac As String = "")
            Try
                oOwnerDB.DBClearCAPAnnualSummary(processingYear, mode, ownerID, fac)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Sub ClearCAPAnnualCalendar(ByVal processingYear As Integer, ByVal ownerID As Integer, Optional ByVal mode As Integer = 0, Optional ByVal procMonth As Integer = -1, Optional ByVal fac As String = "")
            Try
                oOwnerDB.DBClearCAPAnnualCalendar(processingYear, ownerID, mode, procMonth, fac)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub LoadFeesSpecific(ByVal OwnerID As Int64)
            Dim dsRemSys As New DataSet
            Dim dtReturn As Single
            Dim strSQL As String
            Dim tmpDate As Date
            Try
                strBankruptChapter = ""
                dtBankruptDate = tmpDate
                strBP2KOwnerID = ""

                strSQL = "Select Bankrupt_Chapter, Bankrupt_Date, BP2K_Owner_ID from tblREG_Owner where Owner_ID =  " & OwnerID

                dsRemSys = oOwnerDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    If IsDBNull(dsRemSys.Tables(0).Rows(0)("Bankrupt_Chapter")) = False Then
                        strBankruptChapter = dsRemSys.Tables(0).Rows(0)("Bankrupt_Chapter")
                    End If
                    If IsDBNull(dsRemSys.Tables(0).Rows(0)("Bankrupt_Date")) = False Then
                        dtBankruptDate = dsRemSys.Tables(0).Rows(0)("Bankrupt_Date")
                    End If
                    If IsDBNull(dsRemSys.Tables(0).Rows(0)("BP2K_Owner_ID")) = False Then
                        strBP2KOwnerID = dsRemSys.Tables(0).Rows(0)("BP2K_Owner_ID")
                    End If
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#End Region
#Region "Event Handlers"
        'Events related to Owner
        Private Sub FacFlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer) Handles oOwnFacility.FlagsChanged
            RaiseEvent FlagsChanged(entityID, entityType)
        End Sub
        Private Sub OwnersChanged(ByVal strSrc As String) Handles colOwners.OwnerColChanged
            RaiseEvent evtOwnersChanged(Me.colIsDirty)
        End Sub
        Private Sub OwnerChanged(ByVal bolValue As Boolean) Handles oOwnerInfo.OwnerInfoChanged
            RaiseEvent evtOwnerChanged(bolValue)
        End Sub
        'Private Sub OwnerCommentsChanged(ByVal bolValue As Boolean) Handles oComments.InfoBecameDirty
        '    RaiseEvent evtOwnerCommentsChanged(bolValue)
        'End Sub

        'Events related to Facility
        Private Sub FacilityErr(ByVal MsgStr As String) Handles oOwnFacility.evtFacilityErr
            RaiseEvent evtOwnerErr(MsgStr)
        End Sub
        Private Sub FacilityChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtFacilityChanged
            RaiseEvent evtFacilityChanged(bolValue)
        End Sub
        Private Sub FacilitiesChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtFacilitiesChanged
            RaiseEvent evtFacilitiesChanged(bolValue)
        End Sub
        Private Sub FacilityValidationErr(ByVal FacID As Integer, ByVal MsgStr As String) Handles oOwnFacility.evtFacilityValidationErr
            RaiseEvent evtValidationErr(FacID, MsgStr)
            Exit Sub
        End Sub
        'Private Sub FacilityCommentsChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtFacilityCommentsChanged
        '    RaiseEvent evtFacilityCommentsChanged(bolValue)

        'End Sub
        Private Sub facilityCapstatusChanged(ByVal facID As Integer) Handles oOwnFacility.evtFacilityCAPStatus
            RaiseEvent evtOwnFacilityCAPStatusChanged(True, facID)
        End Sub
        'Private Sub FacilityStatusChanged(ByVal bolStatus As Boolean) Handles oOwnFacility.evtFacilitySaved
        '    If oOwnerInfo.Active <> bolStatus Then
        '        oOwnerInfo.Active = bolStatus
        '    End If
        'End Sub
        'Private Sub evtFacilityIDChanged(ByVal oldID As Integer, ByVal newID As Integer) Handles oOwnFacility.evtFacilityIDChanged
        '    Try
        '        If Not (dtFacTnkTable Is Nothing) Then
        '            Dim dr As DataRow
        '            Dim index As Integer = 0
        '            If newID = 0 Then
        '                ' delete row from table
        '                'For Each dr In dtFacTnkTable.Rows
        '                '    If CType(dtFacTnkTable.Rows(index).Item("FacilityID"), Integer) = oldID Then
        '                '        dtFacTnkTable.Rows.Remove(dr)
        '                '    End If
        '                'Next

        '                'For index = 0 To dtFacTnkTable.Rows.Count - 1
        '                '    If CType(dtFacTnkTable.Rows(index).Item("FacilityID"), Integer) = oldID Then
        '                '        dtFacTnkTable.Rows(index).Delete()
        '                '        dtFacTnkTable.AcceptChanges()
        '                '        Exit For
        '                '    End If
        '                'Next
        '            Else
        '                For index = 0 To dtFacTnkTable.Rows.Count - 1
        '                    If CType(dtFacTnkTable.Rows(index).Item("FacilityID"), Integer) = oldID Then
        '                        dtFacTnkTable.Rows(index).Item("FacilityID") = newID
        '                        Exit For
        '                    End If
        '                Next
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub

        ' Events related to Tanks
        'Private Sub TankCommentsChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtTankCommentsChanged
        '    RaiseEvent evtTankCommentsChanged(bolValue)
        'End Sub
        'Private Sub TankValidationError(ByVal tnkID As Integer, ByVal strMessage As String) Handles oOwnFacility.evtTankValidationErr
        '    RaiseEvent evtTankValidationErr(tnkID, strMessage)
        'End Sub
        Private Sub PersonaErr(ByVal MsgStr As String) Handles oOwnPersona.evtPersonaErr
            RaiseEvent evtOwnerErr(MsgStr)
            Exit Sub
        End Sub
        Private Sub PersonaChanged(ByVal bolValue As Boolean) Handles oOwnPersona.evtPersonaChanged
            RaiseEvent evtPersonaChanged(bolValue)
        End Sub
        Private Sub PersonasChanged(ByVal bolValue As Boolean) Handles oOwnPersona.evtPersonasChanged
            RaiseEvent evtPersonasChanged(bolValue)
        End Sub

        'Private Sub PipeCommentsChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtPipeCommentsChanged
        '    RaiseEvent evtPipeCommentsChanged(bolValue)
        'End Sub
        'Events added by kumar
        'event to handle the facility collection of a specific owner
        'Private Sub FacColOwner(ByVal ownerID As Integer, ByVal facilityCol As MUSTER.info.FacilityCollection) Handles oOwnFacility.evtFacColOwner
        '    Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
        '    Try
        '        oOwnerInfoLocal = colOwners.Item(ownerID)
        '        If Not (oOwnerInfoLocal Is Nothing) Then
        '            oOwnerInfoLocal.facilityCollection = facilityCol
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'event to handle the Comments collection of a specific owner
        'Private Sub CommentsColOwner(ByVal ownerID As Integer, ByVal CommentsCol As MUSTER.info.CommentsCollection) Handles oComments.evtCommentColOwner
        '    Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
        '    Try
        '        oOwnerInfoLocal = colOwners.Item(ownerID)
        '        If Not (oOwnerInfoLocal Is Nothing) Then
        '            oOwnerInfoLocal.commentsCollection = CommentsCol
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'event to add the tank collection to its specific facility
        'Private Sub evtTankColFacOwner(ByVal facID As Integer, ByVal tankCol As MUSTER.Info.TankCollection) Handles oOwnFacility.evtTankColFacOwner
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Try
        '        oFacilityInfoLocal = oOwnerInfo.facilityCollection.Item(facID)
        '        If Not (oFacilityInfoLocal Is Nothing) Then
        '            oFacilityInfoLocal.TankCollection = tankCol
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'event to add the compartment collection to its specific tank
        'Private Sub evtCompColTank(ByVal TankID As Integer, ByVal compartmentCol As MUSTER.Info.CompartmentCollection, ByVal facID As Integer) Handles oOwnFacility.evtCompartmentCol
        '    Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    Try
        '        oTankInfoLocal = oOwnerInfo.facilityCollection.Item(facID).TankCollection.Item(TankID)
        '        If Not (oTankInfoLocal Is Nothing) Then
        '            oTankInfoLocal.CompartmentCollection = compartmentCol
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'event to add the comments collection to its specific facility
        'Private Sub FaccommentsCol(ByVal FacId As Integer, ByVal CommentsCol As MUSTER.Info.CommentsCollection)
        '    Dim ofacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Try
        '        ofacilityInfoLocal = oOwnerInfo.facilityCollection.Item(FacId)
        '        If Not (ofacilityInfoLocal Is Nothing) Then
        '            ofacilityInfoLocal.CommentsCollection = CommentsCol
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'event to add the comments collection to its specific Tank
        'Private Sub TankCommentsCol(ByVal tankId As Integer, ByVal CommentsCol As MUSTER.Info.CommentsCollection, ByVal facID As Integer) Handles oOwnFacility.evtCommentsColTank
        '    Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    Try
        '        oTankInfoLocal = oOwnerInfo.facilityCollection.Item(facID).TankCollection.Item(tankId)
        '        If Not (oTankInfoLocal Is Nothing) Then
        '            oTankInfoLocal.commentsCollection = CommentsCol
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub PipeCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal tankID As Integer, ByVal facID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oOwnFacility.evtPipeCommentsCol
        '    Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
        '    Try
        '        oPipeInfoLocal = oOwnerInfo.facilityCollection.Item(facID).TankCollection.Item(tankID).pipesCollection.Item(tankID.ToString + "|" + compID.ToString + "|" + pipeID.ToString)
        '        If Not (oPipeInfoLocal Is Nothing) Then
        '            oPipeInfoLocal.commentsCollection = commentsCol
        '        End If
        '        'if not 
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'end changes
        'Private Sub CommentInfoOwner(ByVal ownerID As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo) Handles oComments.evtCommentInfoOwner
        '    Dim oOwnerInfoLocal As MUSTER.Info.OwnerInfo
        '    Try
        '        oOwnerInfoLocal = colOwners.Item(ownerID)
        '        If Not (oOwnerInfoLocal Is Nothing) Then
        '            If oOwnerInfoLocal.commentsCollection.Contains(commentsInfo) Then
        '                oOwnerInfoLocal.commentsCollection.Remove(commentsInfo)
        '            End If
        '            oOwnerInfoLocal.commentsCollection.Add(commentsInfo)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CommentInfoFac(ByVal Facid As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo) Handles oComments.evtCommentInfoFac
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Try
        '        oFacilityInfoLocal = oOwnerInfo.facilityCollection.Item(Facid)
        '        If Not (oFacilityInfoLocal Is Nothing) Then
        '            If oFacilityInfoLocal.CommentsCollection.Contains(commentsInfo) Then
        '                oFacilityInfoLocal.CommentsCollection.Remove(commentsInfo)
        '            End If
        '            oFacilityInfoLocal.CommentsCollection.Add(commentsInfo)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub

        'Private Sub TankInfoFac(ByVal tankInfo As MUSTER.Info.TankInfo, ByVal strDesc As String) Handles oOwnFacility.evtTankInfoFac
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Try
        '        oFacilityInfoLocal = oOwnerInfo.facilityCollection.Item(tankInfo.FacilityId)
        '        If Not (oFacilityInfoLocal Is Nothing) Then
        '            If oFacilityInfoLocal.TankCollection.Contains(tankInfo) Then
        '                oFacilityInfoLocal.TankCollection.Remove(tankInfo)
        '            End If
        '            If UCase(strDesc).Trim = "ADD" Then
        '                oFacilityInfoLocal.TankCollection.Add(tankInfo)
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CompInfoTank(ByVal compartmentInfo As MUSTER.Info.CompartmentInfo, ByVal strDesc As String) Handles oOwnFacility.evtCompInfoTank
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    Try
        '        For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
        '            If oFacilityInfoLocal.TankCollection.Contains(compartmentInfo.TankId) Then
        '                oTankInfoLocal = oFacilityInfoLocal.TankCollection.Item(compartmentInfo.TankId)
        '            End If
        '        Next
        '        If Not (oTankInfoLocal Is Nothing) Then
        '            If oTankInfoLocal.CompartmentCollection.Contains(compartmentInfo) Then
        '                oTankInfoLocal.CompartmentCollection.Remove(compartmentInfo)
        '            End If
        '            If UCase(strDesc).Trim = "ADD" Then
        '                oTankInfoLocal.CompartmentCollection.Add(compartmentInfo)
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub PipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String) Handles oOwnFacility.evtPipeInfoCompartment
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
        '    Try
        '        oFacilityInfoLocal = oOwnerInfo.facilityCollection.Item(pipeInfo.FacilityID)
        '        If Not (oFacilityInfoLocal Is Nothing) Then
        '            oTankInfoLocal = oFacilityInfoLocal.TankCollection.Item(pipeInfo.TankID)
        '        End If
        '        If Not (oTankInfoLocal Is Nothing) Then
        '            oCompartmentInfoLocal = oTankInfoLocal.CompartmentCollection.Item(pipeInfo.TankID.ToString() + "|" + pipeInfo.CompartmentNumber.ToString())
        '        End If
        '        If Not (oCompartmentInfoLocal Is Nothing) Then
        '            If oTankInfoLocal.pipesCollection.Contains(pipeInfo) Then
        '                oTankInfoLocal.pipesCollection.Remove(pipeInfo)
        '            End If
        '            If UCase(strDesc).Trim = "ADD" Then
        '                oTankInfoLocal.pipesCollection.Add(pipeInfo)
        '            End If
        '        End If

        '        'For Each oFacilityInfoLocal In oOwnerInfo.facilityCollection.Values
        '        '    If oFacilityInfoLocal.TankCollection.Contains(pipeInfo.TankID) Then
        '        '        oTankInfoLocal = oFacilityInfoLocal.TankCollection.Item(pipeInfo.TankID)
        '        '    End If
        '        'Next
        '        'For Each oCompartmentInfoLocal In oTankInfoLocal.CompartmentCollection.Values
        '        '    If oCompartmentInfoLocal.COMPARTMENTNumber = pipeInfo.CompartmentNumber Then
        '        '        Exit For
        '        '    End If
        '        'Next
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub FacInfoOwner(ByVal facInfo As MUSTER.Info.FacilityInfo, ByVal strDesc As String) Handles oOwnFacility.evtFacInfoOwner
        '    Try
        '        Dim colTankLocal As New MUSTER.Info.TankCollection
        '        If oOwnerInfo.facilityCollection.Contains(facInfo) Then
        '            If oOwnerInfo.facilityCollection.Item(facInfo.ID).TankCollection.Count > 0 Then
        '                colTankLocal = oOwnerInfo.facilityCollection.Item(facInfo.ID).TankCollection
        '            End If
        '            oOwnerInfo.facilityCollection.Remove(facInfo)
        '        End If
        '        If UCase(strDesc).Trim = "ADD" Then
        '            oOwnerInfo.facilityCollection.Add(facInfo)
        '            oOwnerInfo.facilityCollection.Item(facInfo.ID).TankCollection = colTankLocal
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub FacInfoFacID(ByVal facID As Integer) Handles oOwnFacility.evtFacInfoFacID
        '    Try
        '        oOwnerInfo.facilityCollection.Remove(facID)
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub OwnerInfoFacCol(ByRef colFac As MUSTER.Info.FacilityCollection) Handles oOwnFacility.evtOwnerInfoFacCol
        '    colFac = oOwnerInfo.facilityCollection
        'End Sub
        'Private Sub OwnerInfoFacColByOwnerID(ByVal ownerID As Integer, ByRef colFac As MUSTER.Info.FacilityCollection) Handles oOwnFacility.evtOwnerInfoFacColByOwnerID
        '    Try
        '        oOwnerInfo = colOwners.Item(ownerID)
        '        If Not (oOwnerInfo Is Nothing) Then
        '            colFac = oOwnerInfo.facilityCollection
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub FacilityInfoTankColByFacilityID(ByVal facID As Integer, ByRef colTnk As MUSTER.Info.TankCollection) Handles oOwnFacility.evtFacilityInfoTankColByFacilityID
        '    Dim oFacilityInfoLocal As MUSTER.Info.FacilityInfo
        '    Try
        '        oFacilityInfoLocal = oOwnerInfo.facilityCollection.Item(facID)
        '        If Not (oFacilityInfoLocal Is Nothing) Then
        '            colTnk = oFacilityInfoLocal.TankCollection
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub FacilityChangeKey(ByVal oldID As Integer, ByVal newID As Integer) Handles oOwnFacility.evtFacilityChangeKey
        '    Try
        '        If oOwnerInfo.facilityCollection.Contains(oldID) Then
        '            oOwnerInfo.facilityCollection.ChangeKey(oldID, newID)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        ' Manju END
        Private Sub OwnerAddressChanged(ByVal bolValue As Boolean) Handles oOwnAddress.evtAddressChanged
            RaiseEvent evtAddressChanged(bolValue)
            oOwnerInfo.AddressId = oOwnAddress.AddressId
        End Sub
        Private Sub OwnerAddressesChanged(ByVal bolValue As Boolean) Handles oOwnAddress.evtAddressesChanged
            RaiseEvent evtAddressesChanged(bolValue)
            oOwnerInfo.AddressId = oOwnAddress.AddressId
        End Sub
        Private Sub FacilityAddressChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtAddressChanged
            RaiseEvent evtAddressChanged(bolValue)
            oOwnFacility.Facility.AddressID = oOwnFacility.AddressID

        End Sub
        Private Sub FacilityAddressesChanged(ByVal bolValue As Boolean) Handles oOwnFacility.evtAddressesChanged
            RaiseEvent evtAddressesChanged(bolValue)
            oOwnFacility.Facility.AddressID = oOwnFacility.AddressID

        End Sub
#End Region
    End Class
End Namespace
