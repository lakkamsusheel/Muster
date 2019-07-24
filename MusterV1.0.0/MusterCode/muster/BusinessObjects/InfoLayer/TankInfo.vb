'-------------------------------------------------------------------------------
' MUSTER.Info.TankInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/15/04    Original class definition.
'  1.1        KJ      12/22/04    Added the Archive Method and also Added more text amd descriptions in the header
'  1.2        KJ      12/30/04    Added event for data update notification.
'                                   Added firing of event in CHECKDIRTY() if
'                                     dirty state changed.
'  1.3        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.4        EN      01/21/05    Added source column in Event. 
'  1.5        MR      03/07/05    Change Modified By and Modified On to read/write.
'  1.6        MNR     03/10/05    Added OriginalTankStatus Property which exposes Tank Status value when Tank was originally loaded / last time saved
'  1.7        MR      03/14/05    Change Created By and Created On to read/write.
'  1.8        MNR     03/15/05    Updated Constructor New(ByVal drTank As DataRow) to check for System.DBNull.Value
'  1.9        AB      03/16/05    Added AgeThreshold and IsAgedData Attributes
'  2.0        MNR     03/16/05    Removed strSrc from events
'  2.1        KKM     03/18/05    CompartmentCollection and commentsCollection properties are added
'
' Function                              Description
' New()                         Instantiates an empty TankInfo object.
' New(TANK_ID,TANK_INDEX,FACILITY_ID,TANKSTATUS,DATERECEIVED,MANIFOLD,COMPARTMENT,TANKCAPACITY,SUBSTANCE,CASNUMBER
'           SUBSTANCECOMMENTS_ID,DATELASTUSED,DATECLOSURERECEIVED,DATECLOSED,CLOSURESTATUSDESC,INERTMATERIAL,TANKMATDESC
'           TANKMODDESC,TANKOTHERMATERIAL,OVERFILLINSTALLED,SPILLINSTALLED,LICENSEEID,CONTRACTORID,DATESIGNED,DATEINSTALLEDTANK
'           SMALLDELIVERY,TANKEMERGEN,PLANNEDINSTDATE,LASTTCPDATE,LINEDINTERIORINSTALLDATE,LINEDINTERIORINSPECTDATE,TCPINSTALLDATE
'           TTTDATE,TANKLD,OVERFILLTYPE,TIGHTFILLADAPTERS,DROPTUBE,TANKCPTYPE,PLACEDINSERVICEDATE,TANKTYPES,TANKLOCATION_DESCRIPTION
'           TANKMANUFACTURER,DELETED,CREATED_BY,DATE_CREATED,LAST_EDITED_BY,DATE_LAST_EDITED)
'                               Instantiates a populated TankInfo object.
' New(dr)                       Instantiates a populated TankInfo object taking member state
'                                   from the datarow provided.
' Archive()                     Sets the object state to the new state 
' Reset()                       Sets the object state to the original state when loaded from or
'                                   last saved to the repository.
'
' Read-Write Attributes
' Attribute                               Description
'   TANK_ID                     The primary key associated with the TankInfo in the repository.
'   TANK_INDEX                  The Index ID to show Numbers of Tanks for a Facility
'   FACILITY_ID                 The facility ID associated with TankInfo
'   TANKSTATUS                  The Status of TankInfo object. CIU, TOS, POS, TOSI, U
'   DATERECEIVED
'   MANIFOLD
'   COMPARTMENT
'   TANKCAPACITY
'   SUBSTANCE
'   CASNUMBER
'   SUBSTANCECOMMENTS_ID
'   DATELASTUSED                Tank Last Used(Closure info)
'   DATECLOSURERECEIVED         Tank closure Received Date
'   DATECLOSED                  Tank Closed Date(Closure)
'   CLOSURESTATUSDESC           Tank Closure Status Description
'   INERTMATERIAL               Tank Inert fill(Closure)
'   TANKMATDESC                 Tank Material Description
'   TANKMODDESC                 Same as Tank Secondary Option - corresponding to TankLD               
'   TANKOTHERMATERIAL
'   OVERFILLINSTALLED           Over fill protected
'   SPILLINSTALLED              Spill protected
'   LICENSEEID
'   CONTRACTORID
'   DATESIGNED                  Installer signed date
'   DATEINSTALLEDTANK           Tank Install date
'   SMALLDELIVERY               
'   TANKEMERGEN                 Tank Used for emergency
'   PLANNEDINSTDATE             Date tank planned to install(Reg pending)
'   LASTTCPDATE                 Tank CP Test date
'   LINEDINTERIORINSTALLDATE    Lined Interior Install Date
'   LINEDINTERIORINSPECTDATE    Lined Interior Install Inspect
'   TCPINSTALLDATE              Tank CP Install date
'   TTTDATE                     Tank Tightness Test Date
'   TANKLD                      Tank Release/Leak Detection
'   OVERFILLTYPE                Tank Over fill type                 
'   TIGHTFILLADAPTERS           Tight fill Adapter    
'   DROPTUBE                    Drop tube for Inventory Control      
'   TANKCPTYPE                  Tank CP Type
'   PLACEDINSERVICEDATE         The date on which the Tank was placed in Service
'   TANKTYPES                   Type of Tank
'   TANKLOCATION_DESCRIPTION    Description of Tank Location.
'   TANKMANUFACTURER            The manufacturer of the Tank
'   DELETED                     Shows if the Tank has been deleted.
'   IsDirty                     Indicates if the Tank state has been altered since it was
'                                   last loaded from or saved to the repository.
' Read-Only Attributes
' CreatedBy                     The name of the user that created the TankInfo object.
' CreatedOn                     The date that the TankInfo object was created.
' ModifiedBy                    The name of the user that last modified the TankInfo object.
' ModifiedOn                    The date that the TankInfo object was last modified.
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/05 - JVC2
' TODO - check properties and operations against list.
'

Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Namespace MUSTER.Info
    <Serializable()> _
Public Class TankInfo

#Region "Private member variables"
        ' Declaring the Original Properties for the Tank - Value Object
        Private onTankId As Integer
        Private onTankIndex As Integer
        Private onFacilityId As Integer
        Private onTankStatus As Integer
        Private odtDateReceived As Date
        Private obolManifold As Boolean
        Private obolCompartment As Boolean
        Private onTankCapacity As Integer
        Private onSubstance As Integer
        Private onCASNumber As Integer                           'CERCLA No
        Private onSubstanceCommentsID As Integer
        Private odtDateLastUsed As Date                          'Tank Last Used(Closure info)
        Private odtDateClosureReceived As Date                   'Tank closure Received Date
        Private odtDateClosed As Date                            'Tank Closed Date(Closure)
        Private onClosureStatusDesc As Integer
        Private onClosureType As Integer
        Private onInertMaterial As Integer                       'Tank Inert fill(Closure)
        Private onTankMatDesc As Integer
        Private onTankModDesc As Integer
        Private onTankOtherMaterial As Integer
        Private obolOverFillInstalled As Boolean                 'Over fill protected
        Private obolSpillInstalled As Boolean                    'Spill protected
        Private onLicenseeId As Integer
        Private onContractorId As Integer
        Private odtDateSigned As Date                            'Installer signed date
        Private odtDateInstalledTank As Date                     'Tank Install date
        Private odtDateSpillInstalled As Date                    'Tank  Spill Install date
        Private odtDateSpillTested As Date                       'Tank  Spill Test date
        Private odtDateOverfillInstalled As Date                 'Tank  Overfill Installed date
        Private odtDateOverfillTested As Date                    'Tank  Overfill Test date
        'Private odtDateTankSecInsp As Date                       'Tank  Secondary Insp date
        Private odtDateTankElecInsp As Date                      'Tank  Electronic Insp date
        Private odtDateATGInsp As Date                           'Tank  ATG Test date
        Private obolSmallDelivery As Boolean                     'Dont know
        Private obolTankEmergen As Boolean                       'Tank Used for emergency
        Private odtPlannedInstDate As Date                       'Date tank planned to install(Reg pending)
        Private odtLastTCPDate As Date                           'Tank CP Test date

        Private odtLinedInteriorInstallDate As Date              'Lined Interior Install Date
        Private odtLinedInteriorInspectDate As Date              'Lined Interior Install Inspect
        Private odtTCPInstallDate As Date                        'Tank CP Install date
        Private odtTTTDate As Date                               'Tank Tightness Test Date
        Private onTankLD As Integer                              'Tank Release Detection
        Private onOverFillType As Integer                        'Tank Over fill type
        Private onRevokeReason As Integer                        'Tank Revoke Reason
        Private onRevokeDate As Date                             'Tank Revoke Date
        Private onDatePhysicallyTagged As Date                   'Tank Prohibition DatePhysicallyTagged
        Private obolProhibition As Boolean                       'Prohibition
        Private obolTightFillAdapters As Boolean                 'Tight fill Adapter
        Private obolDropTube As Boolean                          'Drop tube for Inventory Control
        Private onTankCPType As Integer                          'Tank CP Type
        Private odtPlacedInServiceDate As Date                   'Tank placed in service 
        Private onTankTypes As Integer                           'Tank types
        Private ostrTankLocationDesc As String                   'Tank Location Description
        Private onTankManufacturer As Integer
        Private obolDeleted As Boolean                           'Deleted Flag

        ' Declaring the Current Properties for the Tank - Value Object
        Private nTankId As Integer
        Private nTankIndex As Integer
        Private nFacilityId As Integer
        Private nTankStatus As Integer
        Private dtDateReceived As Date
        Private bolManifold As Boolean
        Private bolCompartment As Boolean
        Private nTankCapacity As Integer
        Private nSubstance As Integer
        Private nCASNumber As Integer                           'CERCLA No
        Private nSubstanceCommentsID As Integer
        Private dtDateLastUsed As Date                          'Tank Last Used(Closure info)
        Private dtDateClosureReceived As Date                   'Tank closure Received Date
        Private dtDateClosed As Date                            'Tank Closed Date(Closure)
        Private nClosureStatusDesc As Integer
        Private nClosureType As Integer
        Private nInertMaterial As Integer                       'Tank Inert fill(Closure)
        Private nTankMatDesc As Integer
        Private nTankModDesc As Integer
        Private nTankOtherMaterial As Integer
        Private bolOverFillInstalled As Boolean                 'Over fill protected
        Private bolSpillInstalled As Boolean                    'Spill protected
        Private nLicenseeId As Integer
        Private nContractorId As Integer
        Private dtDateSigned As Date                            'Installer signed date
        Private dtDateInstalledTank As Date                     'Tank Install date
        Private dtDateSpillInstalled As Date                    'Tank Spill Install date
        Private dtDateSpillTested As Date                       'Tank  Spill Test date
        Private dtDateOverfillInstalled As Date                 'Tank  Overfill Installed date
        Private dtDateOverfillTested As Date                    'Tank  Overfill Test date
        'Private dtDateTankSecInsp As Date                       'Tank  Secondary Insp date
        Private dtDateTankElecInsp As Date                      'Tank  Electronic Insp date
        Private dtDateATGInsp As Date                           'Tank  ATG Test date
        Private bolSmallDelivery As Boolean                     'Dont know
        Private bolTankEmergen As Boolean                       'Tank Used for emergency
        Private dtPlannedInstDate As Date                       'Date tank planned to install(Reg pending)
        Private dtLastTCPDate As Date                           'Tank CP Test date
        Private dtLinedInteriorInstallDate As Date              'Lined Interior Install Date
        Private dtLinedInteriorInspectDate As Date              'Lined Interior Install Inspect
        Private dtTCPInstallDate As Date                        'Tank CP Install date
        Private dtTTTDate As Date                               'Tank Tightness Test Date
        Private nTankLD As Integer                              'Tank Release Detection
        Private nOverFillType As Integer                        'Tank Over fill type
        Private nRevokeReason As Integer                        'Tank Revoke Reason
        Private nRevokeDate As Date                             'Tank Revoke Date
        Private nDatePhysicallyTagged As Date                   'tank prohibition DatePhysicallyTagged
        Private bolProhibition As Boolean                       'Prohibition
        Private bolTightFillAdapters As Boolean                 'Tight fill Adapter
        Private bolDropTube As Boolean                          'Drop tube for Inventory Control
        Private nTankCPType As Integer                          'Tank CP Type
        Private dtPlacedInServiceDate As Date                   'Tank placed in service 
        Private nTankTypes As Integer                           'Tank types
        Private strTankLocationDesc As String                   'Tank Location Description
        Private nTankManufacturer As Integer
        Private bolDeleted As Boolean                           'Deleted Flag

        Private bolPOU As Boolean
        Private bolNonPre88 As Boolean
        Private bolFacilityPowerOff As Boolean

        Private strCreatedBy As String
        Private dtCreatedOn As Date
        Private strModifiedBy As String
        Private dtModifiedOn As Date

        Private bolIsDirty As Boolean = False
        ' the variable is a place holder for pipe to have access to facility cap status
        Private nFacCapStatus As Integer = 0

        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'added by kiran 
        Private colCompartment As MUSTER.Info.CompartmentCollection
        Private colPipe As MUSTER.Info.PipesCollection
        Private colComments As MUSTER.Info.CommentsCollection
        'end changes
#End Region
#Region "Public Events"
        Public Event evtTankInfoChanged(ByVal bolValue As Boolean)
        'Public Event InfoBecameDirty(ByVal DirtyState As Boolean, ByVal strSrc As String)
        Public Event eInfoTankStatus(ByVal nOldStatus As Integer, ByVal nNewStatus As Integer)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
            'Added by kiran on 03/09/2005
            colCompartment = New MUSTER.Info.CompartmentCollection
            colComments = New MUSTER.Info.CommentsCollection
            colPipe = New MUSTER.Info.PipesCollection
            'end changes
        End Sub

        '   ByVal odtDateTankSecInsp As Date, _                   
        Sub New(ByVal TankID As Integer, _
       ByVal TankIndex As Integer, _
       ByVal FacilityID As Integer, _
       ByVal TankStatus As Integer, _
       ByVal DateReceived As Date, _
       ByVal Manifold As Boolean, _
       ByVal Compartment As Boolean, _
       ByVal TankCapacity As Integer, _
       ByVal Substance As Integer, _
       ByVal CASNumber As Integer, _
       ByVal SubstanceCommentsID As Integer, _
       ByVal DateLastUsed As Date, _
       ByVal DateClosureReceived As Date, _
       ByVal DateClosed As Date, _
       ByVal ClosureStatusDesc As Integer, _
       ByVal ClosureType As Integer, _
       ByVal InertMaterial As Integer, _
       ByVal TankMatDesc As Integer, _
       ByVal TankModDesc As Integer, _
       ByVal TankOtherMaterial As Integer, _
       ByVal OverFillInstalled As Boolean, _
       ByVal SpillInstalled As Boolean, _
       ByVal LicenseeID As Integer, _
       ByVal ContractorID As Integer, _
       ByVal DateSigned As Date, _
       ByVal DateInstalledTank As Date, _
       ByVal DateSpillInstalled As Date, _
       ByVal odtDateSpillTested As Date, _
       ByVal odtDateOverfillInstalled As Date, _
       ByVal odtDateOverfillTested As Date, _
       ByVal odtDateTankElecInsp As Date, _
       ByVal odtDateATGInsp As Date, _
       ByVal SmallDelivery As Boolean, _
       ByVal TankEmergen As Boolean, _
       ByVal PlannedInstDate As Date, _
       ByVal LastTCPDate As Date, _
       ByVal LinedInteriorInstallDate As Date, _
       ByVal LinedInteriorInspectDate As Date, _
       ByVal TCPInstallDate As Date, _
       ByVal TTTDate As Date, _
       ByVal TankLD As Integer, _
       ByVal OverFillType As Integer, _
       ByVal RevokeReason As Integer, _
       ByVal RevokeDate As Date, _
       ByVal DatePhysicallyTagged As Date, _
       ByVal Prohibition As Boolean, _
       ByVal TightFillAdapters As Boolean, _
       ByVal DropTube As Boolean, _
       ByVal TankCPType As Integer, _
       ByVal PlacedInServiceDate As Date, _
       ByVal TankTypes As Integer, _
       ByVal TankLocationDesc As String, _
       ByVal TankManufacturer As Integer, _
       ByVal Deleted As Boolean, _
       ByVal CreatedBy As String, _
       ByVal CreatedOn As Date, _
       ByVal ModifiedBy As String, _
       ByVal LastEdited As Date)
            onTankId = TankID
            onTankIndex = TankIndex
            onFacilityId = FacilityID
            onTankStatus = TankStatus
            odtDateReceived = DateReceived.Date
            obolManifold = Manifold
            obolCompartment = Compartment
            onTankCapacity = TankCapacity
            onSubstance = Substance
            onCASNumber = CASNumber
            onSubstanceCommentsID = SubstanceCommentsID
            odtDateLastUsed = DateLastUsed.Date
            odtDateClosureReceived = DateClosureReceived.Date
            odtDateClosed = DateClosed.Date
            onClosureStatusDesc = ClosureStatusDesc
            onClosureType = ClosureType
            onInertMaterial = InertMaterial
            onTankMatDesc = TankMatDesc
            onTankModDesc = TankModDesc
            onTankOtherMaterial = TankOtherMaterial
            obolOverFillInstalled = OverFillInstalled
            obolSpillInstalled = SpillInstalled
            onLicenseeId = LicenseeID
            onContractorId = ContractorID
            odtDateSigned = DateSigned.Date
            odtDateInstalledTank = DateInstalledTank.Date
            odtDateSpillInstalled = DateSpillInstalled.Date
            odtDateSpillTested = DateSpillTested.Date
            odtDateOverfillInstalled = DateOverfillInstalled.Date
            odtDateOverfillTested = DateOverfillTested.Date
            '    odtDateTankSecInsp = DateTankSecInsp.Date
            odtDateTankElecInsp = DateTankElecInsp.Date
            odtDateATGInsp = DateATGInsp.Date
            obolSmallDelivery = SmallDelivery
            obolTankEmergen = TankEmergen
            odtPlannedInstDate = PlannedInstDate.Date
            odtLastTCPDate = LastTCPDate.Date
            odtLinedInteriorInstallDate = LinedInteriorInstallDate.Date
            odtLinedInteriorInspectDate = LinedInteriorInspectDate.Date
            odtTCPInstallDate = TCPInstallDate.Date
            odtTTTDate = TTTDate.Date
            onTankLD = TankLD
            onOverFillType = OverFillType
            onRevokeReason = RevokeReason
            onRevokeDate = RevokeDate
            onDatePhysicallyTagged = DatePhysicallyTagged
            obolProhibition = Prohibition
            obolTightFillAdapters = TightFillAdapters
            obolDropTube = DropTube
            onTankCPType = TankCPType
            odtPlacedInServiceDate = PlacedInServiceDate.Date
            onTankTypes = TankTypes
            ostrTankLocationDesc = TankLocationDesc
            onTankManufacturer = TankManufacturer
            obolDeleted = Deleted
            strCreatedBy = CreatedBy
            dtCreatedOn = CreatedOn
            strModifiedBy = ModifiedBy
            dtModifiedOn = LastEdited
            dtDataAge = Now()
            bolFacilityPowerOff = False
            'Added by kiran on 03/09/2005
            colCompartment = New MUSTER.Info.CompartmentCollection
            colComments = New MUSTER.Info.CommentsCollection
            colPipe = New MUSTER.Info.PipesCollection
            'end changes
            Me.Reset()
        End Sub
        Sub New(ByVal drTank As DataRow)
            Try
                'IIf(drTank.Item("") Is System.DBNull.Value, 0, drTank.Item(""))
                onTankId = drTank.Item("TANK_ID")
                onTankIndex = drTank.Item("TANK_INDEX")
                onFacilityId = drTank.Item("FACILITY_ID")
                onTankStatus = IIf(drTank.Item("TankStatus") Is System.DBNull.Value, 0, drTank.Item("TankStatus"))
                odtDateReceived = IIf(drTank.Item("DateReceived") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateReceived"))
                odtDateReceived = odtDateReceived.Date
                obolManifold = IIf(drTank.Item("Manifold") Is System.DBNull.Value, False, drTank.Item("Manifold"))
                obolCompartment = IIf(drTank.Item("Compartment") Is System.DBNull.Value, False, drTank.Item("Compartment"))
                onTankCapacity = IIf(drTank.Item("TankCapacity") Is System.DBNull.Value, 0, drTank.Item("TankCapacity"))
                onSubstance = IIf(drTank.Item("Substance") Is System.DBNull.Value, 0, drTank.Item("Substance"))
                onCASNumber = IIf(drTank.Item("CASNumber") Is System.DBNull.Value, 0, drTank.Item("CASNumber"))
                onSubstanceCommentsID = IIf(drTank.Item("SUBSTANCECOMMENTS_ID") Is System.DBNull.Value, 0, drTank.Item("SUBSTANCECOMMENTS_ID"))
                odtDateLastUsed = IIf(drTank.Item("DATELASTUSED") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DATELASTUSED")) 'drTank.Item("DateLastUsed")
                odtDateLastUsed = odtDateLastUsed.Date
                odtDateClosureReceived = IIf(drTank.Item("DateClosureReceived") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateClosureReceived")) 'drTank.Item("DateClosureReceived")
                odtDateClosureReceived = odtDateClosureReceived.Date
                odtDateClosed = IIf(drTank.Item("DateClosed") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateClosed")) 'drTank.Item("DateClosed")
                odtDateClosed = odtDateClosed.Date
                onClosureStatusDesc = IIf(drTank.Item("ClosureStatusDesc") Is System.DBNull.Value, 0, drTank.Item("ClosureStatusDesc"))
                onClosureType = IIf(drTank.Item("CLOSURETYPE") Is System.DBNull.Value, 0, drTank.Item("CLOSURETYPE"))
                onInertMaterial = IIf(drTank.Item("InertMaterial") Is System.DBNull.Value, 0, drTank.Item("InertMaterial"))
                onTankMatDesc = IIf(drTank.Item("TankMatDesc") Is System.DBNull.Value, 0, drTank.Item("TankMatDesc"))
                onTankModDesc = IIf(drTank.Item("TankModDesc") Is System.DBNull.Value, 0, drTank.Item("TankModDesc"))
                onTankOtherMaterial = IIf(drTank.Item("TankOtherMaterial") Is System.DBNull.Value, 0, drTank.Item("TankOtherMaterial"))
                obolOverFillInstalled = IIf(drTank.Item("OverFillInstalled") Is System.DBNull.Value, False, drTank.Item("OverFillInstalled"))
                obolSpillInstalled = IIf(drTank.Item("SpillInstalled") Is System.DBNull.Value, False, drTank.Item("SpillInstalled"))
                onLicenseeId = IIf(drTank.Item("LICENSEEID") Is System.DBNull.Value, 0, drTank.Item("LICENSEEID")) ' drTank.Item("LICENSEEID")
                onContractorId = IIf(drTank.Item("ContractorID") Is System.DBNull.Value, 0, drTank.Item("ContractorID")) ' drTank.Item("ContractorID")
                odtDateSigned = IIf(drTank.Item("DateSigned") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateSigned")) ' drTank.Item("DateSigned")
                odtDateSigned = odtDateSigned.Date
                odtDateInstalledTank = IIf(drTank.Item("DateInstalledTank") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateInstalledTank")) ' drTank.Item("DateInstalledTank")
                odtDateInstalledTank = odtDateInstalledTank.Date
                odtDateSpillInstalled = IIf(drTank.Item("DateSpillPreventionInstalled") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateSpillPreventionInstalled")) ' drTank.Item("DateSpillPreventionInstalled")
                odtDateSpillInstalled = odtDateSpillInstalled.Date
                odtDateSpillTested = IIf(drTank.Item("DateSpillPreventionLastTested") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateSpillPreventionLastTested")) ' drTank.Item("DateSpillPreventionLastTested")
                odtDateSpillTested = odtDateSpillTested.Date
                odtDateOverfillInstalled = IIf(drTank.Item("DateOverfillPreventionInstalled") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateOverfillPreventionInstalled")) ' drTank.Item("DateOverfillPreventionInstalled")
                odtDateOverfillInstalled = odtDateOverfillInstalled.Date
                odtDateOverfillTested = IIf(drTank.Item("DateOverfillPreventionLastInspected") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateOverfillPreventionLastInspected")) ' drTank.Item("DateOverfillPreventionLastInspected")
                odtDateOverfillTested = odtDateOverfillTested.Date
                'odtDateTankSecInsp = IIf(drTank.Item("DateSecondaryContainmentLastInspected") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateSecondaryContainmentLastInspected")) ' drTank.Item("DateSecondaryContainmentLastInspected")
                'odtDateTankSecInsp = odtDateTankSecInsp.Date
                odtDateTankElecInsp = IIf(drTank.Item("DateElectronicDeviceInspected") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateElectronicDeviceInspected")) ' drTank.Item("DateElectronicDeviceInspected")
                odtDateTankElecInsp = odtDateTankElecInsp.Date
                odtDateATGInsp = IIf(drTank.Item("DateATGLastInspected") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DateATGLastInspected")) ' drTank.Item("DateATGLastInspected")
                odtDateATGInsp = odtDateATGInsp.Date
                obolSmallDelivery = IIf(drTank.Item("SmallDelivery") Is System.DBNull.Value, False, drTank.Item("SmallDelivery"))
                obolTankEmergen = IIf(drTank.Item("TankEmergen") Is System.DBNull.Value, False, drTank.Item("TankEmergen"))
                odtPlannedInstDate = IIf(drTank.Item("PlannedInstDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("PlannedInstDate")) ' drTank.Item("PlannedInstDate")
                odtPlannedInstDate = odtPlannedInstDate.Date
                odtLastTCPDate = IIf(drTank.Item("LastTCPDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("LastTCPDate")) ' drTank.Item("LastTCPDate")
                odtLastTCPDate = odtLastTCPDate.Date
                odtLinedInteriorInstallDate = IIf(drTank.Item("LinedInteriorInstallDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("LinedInteriorInstallDate")) ' drTank.Item("LinedInteriorInstallDate")
                odtLinedInteriorInstallDate = odtLinedInteriorInstallDate.Date
                odtLinedInteriorInspectDate = IIf(drTank.Item("LinedInteriorInspectDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("LinedInteriorInspectDate")) ' drTank.Item("LinedInteriorInspectDate")
                odtLinedInteriorInspectDate = odtLinedInteriorInspectDate.Date
                odtTCPInstallDate = IIf(drTank.Item("TCPInstallDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("TCPInstallDate")) ' drTank.Item("TCPInstallDate")
                odtTCPInstallDate = odtTCPInstallDate.Date
                odtTTTDate = IIf(drTank.Item("TTTDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("TTTDate")) ' drTank.Item("TTTDate")
                odtTTTDate = odtTTTDate.Date
                onTankLD = IIf(drTank.Item("TankLD") Is System.DBNull.Value, 0, drTank.Item("TankLD"))
                onOverFillType = IIf(drTank.Item("OverFillType") Is System.DBNull.Value, 0, drTank.Item("OverFillType"))
                onRevokeReason = IIf(drTank.Item("RevokeReason") Is System.DBNull.Value, 0, drTank.Item("RevokeReason"))
                onRevokeDate = IIf(drTank.Item("RevokeDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("RevokeDate"))
                onDatePhysicallyTagged = IIf(drTank.Item("DatePhysicallyTagged") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DatePhysicallyTagged"))
                obolProhibition = IIf(drTank.Item("Prohibition") Is System.DBNull.Value, False, drTank.Item("Prohibition"))
                obolTightFillAdapters = IIf(drTank.Item("TightFillAdapters") Is System.DBNull.Value, False, drTank.Item("TightFillAdapters"))
                obolDropTube = IIf(drTank.Item("DropTube") Is System.DBNull.Value, False, drTank.Item("DropTube"))
                onTankCPType = IIf(drTank.Item("TankCPType") Is System.DBNull.Value, 0, drTank.Item("TankCPType"))
                odtPlacedInServiceDate = IIf(drTank.Item("PlacedInServiceDate") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("PlacedInServiceDate")) ' drTank.Item("PlacedInServiceDate")
                odtPlacedInServiceDate = odtPlacedInServiceDate.Date
                onTankTypes = IIf(drTank.Item("TankTypes") Is System.DBNull.Value, 0, drTank.Item("TankTypes"))
                ostrTankLocationDesc = IIf(drTank.Item("TANKLOCATION_DESCRIPTION") Is System.DBNull.Value, String.Empty, drTank.Item("TANKLOCATION_DESCRIPTION"))
                onTankManufacturer = IIf(drTank.Item("TankManufacturer") Is System.DBNull.Value, 0, drTank.Item("TankManufacturer"))
                '  obolDeleted = IIf(drTank.Item("Deleted") Is System.DBNull.Value, False, drTank.Item("Deleted"))
                If drTank.Item("Deleted") Is System.DBNull.Value Then
                    obolDeleted = False
                Else
                    obolDeleted = drTank.Item("Deleted")
                End If
                strCreatedBy = IIf(drTank.Item("CREATED_BY") Is System.DBNull.Value, String.Empty, drTank.Item("CREATED_BY")) ' drTank.Item("CREATED_BY")
                dtCreatedOn = IIf(drTank.Item("DATE_CREATED") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DATE_CREATED")) ' drTank.Item("DATE_CREATED")
                strModifiedBy = IIf(drTank.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drTank.Item("LAST_EDITED_BY")) ' drTank.Item("LAST_EDITED_BY")
                dtModifiedOn = IIf(drTank.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drTank.Item("DATE_LAST_EDITED")) ' drTank.Item("DATE_LAST_EDITED")
                dtDataAge = Now()
                bolFacilityPowerOff = False
                'Added by kiran on 03/09/2005
                colCompartment = New MUSTER.Info.CompartmentCollection
                colComments = New MUSTER.Info.CommentsCollection
                colPipe = New MUSTER.Info.PipesCollection
                'end changes
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nTankId >= 0 Then
                nTankId = onTankId
            End If
            nTankIndex = onTankIndex
            nFacilityId = onFacilityId
            nTankStatus = onTankStatus
            dtDateReceived = odtDateReceived
            bolManifold = obolManifold
            bolCompartment = obolCompartment
            nTankCapacity = onTankCapacity
            nSubstance = onSubstance
            nCASNumber = onCASNumber
            nSubstanceCommentsID = onSubstanceCommentsID
            dtDateLastUsed = odtDateLastUsed
            dtDateClosureReceived = odtDateClosureReceived
            dtDateClosed = odtDateClosed
            nClosureStatusDesc = onClosureStatusDesc
            nClosureType = onClosureType
            nInertMaterial = onInertMaterial
            nTankMatDesc = onTankMatDesc
            nTankModDesc = onTankModDesc
            nTankOtherMaterial = onTankOtherMaterial
            bolOverFillInstalled = obolOverFillInstalled
            bolSpillInstalled = obolSpillInstalled
            nLicenseeId = onLicenseeId
            nContractorId = onContractorId
            dtDateSigned = odtDateSigned
            dtDateInstalledTank = odtDateInstalledTank
            dtDateSpillInstalled = odtDateSpillInstalled
            dtDateSpillTested = odtDateSpillTested
            dtDateOverfillInstalled = odtDateOverfillInstalled
            dtDateOverfillTested = odtDateOverfillTested
            'dtDateTankSecInsp = odtDateTankSecInsp
            dtDateTankElecInsp = odtDateTankElecInsp
            dtDateATGInsp = odtDateATGInsp
            bolSmallDelivery = obolSmallDelivery
            bolTankEmergen = obolTankEmergen
            dtPlannedInstDate = odtPlannedInstDate
            dtLastTCPDate = odtLastTCPDate
            dtLinedInteriorInstallDate = odtLinedInteriorInstallDate
            dtLinedInteriorInspectDate = odtLinedInteriorInspectDate
            dtTCPInstallDate = odtTCPInstallDate
            dtTTTDate = odtTTTDate
            nTankLD = onTankLD
            nOverFillType = onOverFillType
            nRevokeReason = onRevokeReason
            nRevokeDate = onRevokeDate
            nDatePhysicallyTagged = onDatePhysicallyTagged
            bolProhibition = obolProhibition
            bolTightFillAdapters = obolTightFillAdapters
            bolDropTube = obolDropTube
            nTankCPType = onTankCPType
            dtPlacedInServiceDate = odtPlacedInServiceDate
            nTankTypes = onTankTypes
            strTankLocationDesc = ostrTankLocationDesc
            nTankManufacturer = onTankManufacturer
            bolDeleted = obolDeleted
            bolIsDirty = False
            RaiseEvent evtTankInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onTankId = nTankId
            onTankIndex = nTankIndex
            onFacilityId = nFacilityId
            onTankStatus = nTankStatus
            odtDateReceived = dtDateReceived
            obolManifold = bolManifold
            obolCompartment = bolCompartment
            onTankCapacity = nTankCapacity
            onSubstance = nSubstance
            onCASNumber = nCASNumber
            onSubstanceCommentsID = nSubstanceCommentsID
            odtDateLastUsed = dtDateLastUsed
            odtDateClosureReceived = dtDateClosureReceived
            odtDateClosed = dtDateClosed
            onClosureStatusDesc = nClosureStatusDesc
            onClosureType = nClosureType
            onInertMaterial = nInertMaterial
            onTankMatDesc = nTankMatDesc
            onTankModDesc = nTankModDesc
            onTankOtherMaterial = nTankOtherMaterial
            obolOverFillInstalled = bolOverFillInstalled
            obolSpillInstalled = bolSpillInstalled
            onLicenseeId = nLicenseeId
            onContractorId = nContractorId
            odtDateSigned = dtDateSigned
            odtDateInstalledTank = dtDateInstalledTank
            odtDateSpillInstalled = dtDateSpillInstalled
            odtDateSpillTested = dtDateSpillTested
            odtDateOverfillInstalled = dtDateOverfillInstalled
            odtDateOverfillTested = dtDateOverfillTested
            'odtDateTankSecInsp = dtDateTankSecInsp
            odtDateTankElecInsp = dtDateTankElecInsp
            odtDateATGInsp = dtDateATGInsp
            obolSmallDelivery = bolSmallDelivery
            obolTankEmergen = bolTankEmergen
            odtPlannedInstDate = dtPlannedInstDate
            odtLastTCPDate = dtLastTCPDate
            odtLinedInteriorInstallDate = dtLinedInteriorInstallDate
            odtLinedInteriorInspectDate = dtLinedInteriorInspectDate
            odtTCPInstallDate = dtTCPInstallDate
            odtTTTDate = dtTTTDate
            onTankLD = nTankLD
            onOverFillType = nOverFillType
            onRevokeReason = nRevokeReason
            onRevokeDate = nRevokeDate
            onDatePhysicallyTagged = nDatePhysicallyTagged
            obolProhibition = bolProhibition
            obolTightFillAdapters = bolTightFillAdapters
            obolDropTube = bolDropTube
            onTankCPType = nTankCPType
            odtPlacedInServiceDate = dtPlacedInServiceDate
            onTankTypes = nTankTypes
            ostrTankLocationDesc = strTankLocationDesc
            onTankManufacturer = nTankManufacturer
            obolDeleted = bolDeleted
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            'nTankId <> onTankId Or _
            'nFacilityId <> onFacilityId Or _
            'nTankCapacity <> onTankCapacity Or _

            'dtDateTankSecInsp <> odtDateTankSecInsp Or _

            bolIsDirty = (nTankIndex <> onTankIndex Or _
                        nTankStatus <> onTankStatus Or _
                        dtDateReceived <> odtDateReceived Or _
                        bolManifold <> obolManifold Or _
                        bolCompartment <> obolCompartment Or _
                        nSubstance <> onSubstance Or _
                        nCASNumber <> onCASNumber Or _
                        nSubstanceCommentsID <> onSubstanceCommentsID Or _
                        dtDateLastUsed <> odtDateLastUsed Or _
                        dtDateClosureReceived <> odtDateClosureReceived Or _
                        dtDateClosed <> odtDateClosed Or _
                        nClosureStatusDesc <> onClosureStatusDesc Or _
                        nClosureType <> onClosureType Or _
                        nInertMaterial <> onInertMaterial Or _
                        nTankMatDesc <> onTankMatDesc Or _
                        nTankModDesc <> onTankModDesc Or _
                        nTankOtherMaterial <> onTankOtherMaterial Or _
                        bolOverFillInstalled <> obolOverFillInstalled Or _
                        bolSpillInstalled <> obolSpillInstalled Or _
                        nLicenseeId <> onLicenseeId Or _
                        nContractorId <> onContractorId Or _
                        dtDateSigned <> odtDateSigned Or _
                        dtDateInstalledTank <> odtDateInstalledTank Or _
                        dtDateSpillInstalled <> odtDateSpillInstalled Or _
                        dtDateSpillTested <> odtDateSpillTested Or _
                        dtDateOverfillInstalled <> odtDateOverfillInstalled Or _
                        dtDateOverfillTested <> odtDateOverfillTested Or _
                        dtDateTankElecInsp <> odtDateTankElecInsp Or _
                        dtDateATGInsp <> odtDateATGInsp Or _
                        bolSmallDelivery <> obolSmallDelivery Or _
                        bolTankEmergen <> obolTankEmergen Or _
                        dtPlannedInstDate <> odtPlannedInstDate Or _
                        dtLastTCPDate <> odtLastTCPDate Or _
                        dtLinedInteriorInstallDate <> odtLinedInteriorInstallDate Or _
                        dtLinedInteriorInspectDate <> odtLinedInteriorInspectDate Or _
                        dtTCPInstallDate <> odtTCPInstallDate Or _
                        dtTTTDate <> odtTTTDate Or _
                        nTankLD <> onTankLD Or _
                        nOverFillType <> onOverFillType Or _
                        nRevokeReason <> onRevokeReason Or _
                        nRevokeDate <> onRevokeDate Or _
                        nDatePhysicallyTagged <> onDatePhysicallyTagged Or _
                        bolProhibition <> obolProhibition Or _
                        bolTightFillAdapters <> obolTightFillAdapters Or _
                        bolDropTube <> obolDropTube Or _
                        nTankCPType <> onTankCPType Or _
                        dtPlacedInServiceDate <> odtPlacedInServiceDate Or _
                        nTankTypes <> onTankTypes Or _
                        strTankLocationDesc <> ostrTankLocationDesc Or _
                        nTankManufacturer <> onTankManufacturer Or _
                        bolDeleted <> obolDeleted)
            If bolOldState <> bolIsDirty Then
                'MsgBox("Info F:" + FacilityId.ToString + " TI:" + TankIndex.ToString + " TID:" + TankId.ToString)
                RaiseEvent evtTankInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onTankId = 0
            onTankIndex = 0
            onFacilityId = 0
            onTankStatus = 0
            odtDateReceived = System.DateTime.Now
            obolManifold = False
            obolCompartment = False
            onTankCapacity = 0
            onSubstance = 0
            onCASNumber = 0
            onSubstanceCommentsID = 0
            odtDateLastUsed = CDate("01/01/0001")
            odtDateClosureReceived = CDate("01/01/0001")
            odtDateClosed = CDate("01/01/0001")
            onClosureStatusDesc = 0
            onClosureType = 0
            onInertMaterial = 0
            onTankMatDesc = 0
            onTankModDesc = 0
            onTankOtherMaterial = 0
            obolOverFillInstalled = False
            obolSpillInstalled = False
            onLicenseeId = 0
            onContractorId = 0
            odtDateSigned = CDate("01/01/0001")
            odtDateInstalledTank = CDate("01/01/0001")
            odtDateSpillInstalled = CDate("01/01/0001")
            odtDateSpillTested = CDate("01/01/0001")
            odtDateOverfillInstalled = CDate("01/01/0001")
            odtDateOverfillTested = CDate("01/01/0001")
            'odtDateTankSecInsp = CDate("01/01/0001")
            odtDateTankElecInsp = CDate("01/01/0001")
            odtDateATGInsp = CDate("01/01/0001")
            obolSmallDelivery = SmallDelivery
            obolTankEmergen = TankEmergen
            odtPlannedInstDate = CDate("01/01/0001")
            odtLastTCPDate = CDate("01/01/0001")
            odtLinedInteriorInstallDate = CDate("01/01/0001")
            odtLinedInteriorInspectDate = CDate("01/01/0001")
            odtTCPInstallDate = CDate("01/01/0001")
            odtTTTDate = CDate("01/01/0001")
            onTankLD = 0
            onOverFillType = 0
            onRevokeReason = 0
            onRevokeDate = CDate("01/01/0001")
            onDatePhysicallyTagged = CDate("01/01/0001")
            obolProhibition = False
            obolTightFillAdapters = False
            obolDropTube = False
            onTankCPType = 0
            'odtPlacedInServiceDate = Nothing 'CDate("01/01/0001")
            odtPlacedInServiceDate = CDate("01/01/0001")
            onTankTypes = 0
            ostrTankLocationDesc = String.Empty
            onTankManufacturer = 0
            obolDeleted = False
            dtCreatedOn = CDate("01/01/0001")
            dtModifiedOn = CDate("01/01/0001")
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            bolPOU = False
            bolNonPre88 = False
            bolFacilityPowerOff = False
            nFacCapStatus = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        'added by kumar on March 11th
        Public Property CompartmentCollection() As MUSTER.Info.CompartmentCollection
            Get
                Return colCompartment
            End Get
            Set(ByVal Value As MUSTER.Info.CompartmentCollection)
                colCompartment = Value
            End Set
        End Property

        Public Property commentsCollection() As MUSTER.Info.CommentsCollection
            Get
                Return colComments
            End Get
            Set(ByVal Value As MUSTER.Info.CommentsCollection)
                colComments = Value
            End Set
        End Property

        Public Property pipesCollection() As MUSTER.Info.PipesCollection
            Get
                Return colPipe
            End Get
            Set(ByVal Value As MUSTER.Info.PipesCollection)
                colPipe = Value
            End Set
        End Property
        'End Changes

        Public Property TankId() As Integer
            Get
                Return nTankId
            End Get

            Set(ByVal value As Integer)
                nTankId = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankIndex() As Integer
            Get
                Return nTankIndex
            End Get

            Set(ByVal value As Integer)
                nTankIndex = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FacilityId() As Integer
            Get
                Return nFacilityId
            End Get

            Set(ByVal value As Integer)
                nFacilityId = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankStatus() As Integer
            Get
                Return nTankStatus
            End Get

            Set(ByVal value As Integer)
                nTankStatus = value
                Me.CheckDirty()
                'RaiseEvent eInfoTankStatus(onTankStatus, nTankStatus)
            End Set
        End Property

        Public Property DateReceived() As Date
            Get
                Return dtDateReceived.Date
            End Get

            Set(ByVal value As Date)
                dtDateReceived = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property Manifold() As Boolean
            Get
                Return bolManifold
            End Get

            Set(ByVal value As Boolean)
                bolManifold = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Compartment() As Boolean
            Get
                Return bolCompartment
            End Get

            Set(ByVal value As Boolean)
                bolCompartment = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankCapacity() As Integer
            Get
                Return nTankCapacity
            End Get

            Set(ByVal value As Integer)
                nTankCapacity = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Substance() As Integer
            Get
                Return nSubstance
            End Get

            Set(ByVal value As Integer)
                nSubstance = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property CASNumber() As Integer
            Get
                Return nCASNumber
            End Get

            Set(ByVal value As Integer)
                nCASNumber = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property SubstanceCommentsID() As Integer
            Get
                Return nSubstanceCommentsID
            End Get

            Set(ByVal value As Integer)
                nSubstanceCommentsID = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateLastUsed() As Date
            Get
                Return dtDateLastUsed.Date
            End Get

            Set(ByVal value As Date)
                dtDateLastUsed = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateClosureReceived() As Date
            Get
                Return dtDateClosureReceived.Date
            End Get

            Set(ByVal value As Date)
                dtDateClosureReceived = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateClosed() As Date
            Get
                Return dtDateClosed.Date
            End Get

            Set(ByVal value As Date)
                dtDateClosed = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property ClosureStatusDesc() As Integer
            Get
                Return nClosureStatusDesc
            End Get

            Set(ByVal value As Integer)
                nClosureStatusDesc = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ClosureType() As Integer
            Get
                Return nClosureType
            End Get
            Set(ByVal Value As Integer)
                nClosureType = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property InertMaterial() As Integer
            Get
                Return nInertMaterial
            End Get

            Set(ByVal value As Integer)
                nInertMaterial = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankMatDesc() As Integer
            Get
                Return nTankMatDesc
            End Get

            Set(ByVal value As Integer)
                nTankMatDesc = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankModDesc() As Integer
            Get
                Return nTankModDesc
            End Get

            Set(ByVal value As Integer)
                nTankModDesc = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankOtherMaterial() As Integer
            Get
                Return nTankOtherMaterial
            End Get

            Set(ByVal value As Integer)
                nTankOtherMaterial = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OverFillInstalled() As Boolean
            Get
                Return bolOverFillInstalled
            End Get

            Set(ByVal value As Boolean)
                bolOverFillInstalled = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property SpillInstalled() As Boolean
            Get
                Return bolSpillInstalled
            End Get

            Set(ByVal value As Boolean)
                bolSpillInstalled = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property LicenseeID() As Integer
            Get
                Return nLicenseeId
            End Get

            Set(ByVal value As Integer)
                nLicenseeId = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ContractorID() As Integer
            Get
                Return nContractorId
            End Get

            Set(ByVal value As Integer)
                nContractorId = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateSigned() As Date
            Get
                Return dtDateSigned.Date
            End Get

            Set(ByVal value As Date)
                dtDateSigned = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateInstalledTank() As Date
            Get
                Return dtDateInstalledTank.Date
            End Get

            Set(ByVal value As Date)
                dtDateInstalledTank = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateSpillInstalled() As Date
            Get
                Return dtDateSpillInstalled.Date
            End Get

            Set(ByVal value As Date)
                dtDateSpillInstalled = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateSpillTested() As Date
            Get
                Return dtDateSpillTested.Date
            End Get

            Set(ByVal value As Date)
                dtDateSpillTested = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateOverfillInstalled() As Date
            Get
                Return dtDateOverfillInstalled.Date
            End Get

            Set(ByVal value As Date)
                dtDateOverfillInstalled = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateOverfillTested() As Date
            Get
                Return dtDateOverfillTested.Date
            End Get

            Set(ByVal value As Date)
                dtDateOverfillTested = value.Date
                Me.CheckDirty()
            End Set
        End Property
        '  Public Property DateTankSecInsp() As Date
        '     Get
        '        Return dtDateTankSecInsp.Date
        '   End Get
        '
        '           Set(ByVal value As Date)
        '              dtDateTankSecInsp = value.Date
        '             Me.CheckDirty()
        '        End Set
        '   End Property

        Public Property DateTankElecInsp() As Date
            Get
                Return dtDateTankElecInsp.Date
            End Get

            Set(ByVal value As Date)
                dtDateTankElecInsp = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateATGInsp() As Date
            Get
                Return dtDateATGInsp.Date
            End Get

            Set(ByVal value As Date)
                dtDateATGInsp = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property SmallDelivery() As Boolean
            Get
                Return bolSmallDelivery
            End Get

            Set(ByVal value As Boolean)
                bolSmallDelivery = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankEmergen() As Boolean
            Get
                Return bolTankEmergen
            End Get

            Set(ByVal value As Boolean)
                bolTankEmergen = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PlannedInstDate() As Date
            Get
                Return dtPlannedInstDate.Date
            End Get

            Set(ByVal value As Date)
                dtPlannedInstDate = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property LastTCPDate() As Date
            Get
                Return dtLastTCPDate.Date
            End Get

            Set(ByVal value As Date)
                dtLastTCPDate = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property LinedInteriorInstallDate() As Date
            Get
                Return dtLinedInteriorInstallDate.Date
            End Get

            Set(ByVal value As Date)
                dtLinedInteriorInstallDate = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property LinedInteriorInspectDate() As Date
            Get
                Return dtLinedInteriorInspectDate.Date
            End Get

            Set(ByVal value As Date)
                dtLinedInteriorInspectDate = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public ReadOnly Property LinedInteriorInspectDateOriginal() As Date
            Get
                Return odtLinedInteriorInspectDate.Date
            End Get
        End Property

        Public Property TCPInstallDate() As Date
            Get
                Return dtTCPInstallDate.Date
            End Get

            Set(ByVal value As Date)
                dtTCPInstallDate = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property TTTDate() As Date
            Get
                Return dtTTTDate.Date
            End Get

            Set(ByVal value As Date)
                dtTTTDate = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public ReadOnly Property TTTDateOriginal() As Date
            Get
                Return odtTTTDate.Date
            End Get
        End Property

        Public Property TankLD() As Integer
            Get
                Return nTankLD
            End Get

            Set(ByVal value As Integer)
                nTankLD = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OverFillType() As Integer
            Get
                Return nOverFillType
            End Get

            Set(ByVal value As Integer)
                nOverFillType = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RevokeReason() As Integer
            Get
                Return nRevokeReason
            End Get

            Set(ByVal value As Integer)
                nRevokeReason = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RevokeDate() As Date
            Get
                Return nRevokeDate.Date
            End Get

            Set(ByVal value As Date)
                nRevokeDate = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property DatePhysicallyTagged() As Date
            Get
                Return nDatePhysicallyTagged.Date
            End Get

            Set(ByVal value As Date)
                nDatePhysicallyTagged = value.Date
                Me.CheckDirty()
            End Set
        End Property
        Public Property Prohibition() As Boolean
            Get
                Return bolProhibition
            End Get

            Set(ByVal value As Boolean)
                bolProhibition = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TightFillAdapters() As Boolean
            Get
                Return bolTightFillAdapters
            End Get

            Set(ByVal value As Boolean)
                bolTightFillAdapters = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DropTube() As Boolean
            Get
                Return bolDropTube
            End Get

            Set(ByVal value As Boolean)
                bolDropTube = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankCPType() As Integer
            Get
                Return nTankCPType
            End Get

            Set(ByVal value As Integer)
                nTankCPType = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PlacedInServiceDate() As Date
            Get
                Return dtPlacedInServiceDate.Date
            End Get

            Set(ByVal value As Date)
                dtPlacedInServiceDate = value.Date
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankTypes() As Integer
            Get
                Return nTankTypes
            End Get

            Set(ByVal value As Integer)
                nTankTypes = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankLocationDescription() As String
            Get
                Return strTankLocationDesc
            End Get

            Set(ByVal value As String)
                strTankLocationDesc = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TankManufacturer() As Integer
            Get
                Return nTankManufacturer
            End Get

            Set(ByVal value As Integer)
                nTankManufacturer = value
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
                If bolIsDirty Then Return bolIsDirty
                For Each comp As MUSTER.Info.CompartmentInfo In CompartmentCollection.Values
                    If comp.IsDirty Then
                        Return True
                    End If
                Next
                For Each pipe As MUSTER.Info.PipeInfo In pipesCollection.Values
                    If pipe.IsDirty Then
                        Return True
                    End If
                Next
                Return False
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property
        Public Property POU() As Boolean
            Get
                Return bolPOU
            End Get
            Set(ByVal Value As Boolean)
                bolPOU = Value
            End Set
        End Property
        Public Property NonPre88() As Boolean
            Get
                Return bolNonPre88
            End Get
            Set(ByVal Value As Boolean)
                bolNonPre88 = Value
            End Set
        End Property
        Public Property FacilityPowerOff() As Boolean
            Get
                Return bolFacilityPowerOff
            End Get
            Set(ByVal Value As Boolean)
                bolFacilityPowerOff = Value
            End Set
        End Property
        Public ReadOnly Property OriginalTankStatus() As Integer
            Get
                Return onTankStatus
            End Get
        End Property

        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get

            Set(ByVal value As Int16)
                nAgeThreshold = value
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get

        End Property
        Public Property FacCapStatus() As Integer
            Get
                Return nFacCapStatus
            End Get

            Set(ByVal value As Integer)
                nFacCapStatus = value
            End Set
        End Property
        Public ReadOnly Property ChildrenDirty() As Boolean
            Get
                Return CompartmentsDirty Or PipesDirty
            End Get
        End Property
        Public ReadOnly Property CompartmentsDirty() As Boolean
            Get
                For Each comp As MUSTER.Info.CompartmentInfo In CompartmentCollection.Values
                    If comp.IsDirty Then Return True
                Next
                Return False
            End Get
        End Property
        Public ReadOnly Property PipesDirty() As Boolean
            Get
                For Each pipe As MUSTER.Info.PipeInfo In pipesCollection.Values
                    If pipe.IsDirty Then Return True
                Next
                Return False
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
                Return strModifiedBy
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


