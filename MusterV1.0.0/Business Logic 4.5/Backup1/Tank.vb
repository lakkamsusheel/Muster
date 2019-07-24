'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Tank
'   Provides the info and collection objects to the client for manipulating
'   an TankInfo object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         KJ      12/16/04    Original class definition.
'   1.1         KJ      12/27/04    Made changes to the ValidateData function for Tank.
'   1.2         KJ      12/28/04    Changes to the Look up functions. All point now to GetDataTable
'   1.3         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.4         KJ      01/03/05    Changed Exposed Attributes, Retrieve and Add Functions
'   1.5         KJ      01/05/05    Added Events for Enabling/Disbaling dates. Streamlined code for PopulateReleaseDetection.
'   1.6         EN      01/21/05    Added all new functions in RaisingEvents for "Enable/Disabling controls in the Form" Region
'                                   changed the events name  in external event handlers and add new events
'                                   Added MusterException line in all raisingEvents Function.
'   1.7         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.8         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Tank" type.
'                                       Also added attribute EntityType to expose Type ID.
'   1.9         AN      02/02/05    Added Comments object
'   2.0         EN      02/10/05    Added TankCAPTable and CheckCAPStatus  and added  in all date properties to check the CAPSTATUS..
'   2.1         JVCII   02/14/05    Added data types to columns in EntityTable
'   2.2         MR      03/07/05    Added iAccessor Attributes to make them exposed to UI.
'   2.3         MNR     03/09/05    Implemented TOS/TOSI/POU rules
'   2.4         MNR     03/10/05    Implemented CIU to TOS/TOSI and TOS/TOSI to CIU rules
'   2.5         MNR     03/15/05    Added Load Sub
'   2.6         AB      03/16/05    Added DataAge check to the Retrieve function
'   2.7         MNR     03/16/05    Removed strSrc from events
'   2.8         MNR     03/18/05    Added Property FacilityPowerOff
'   3.1         KKM     03/18/05    Events for handling local CompartmentCollection, pipesCollection and CommentsCollection are added
'
' Function                          Description
' New()                 Initializes the TankCollection and TankInfo objects.
' Retrieve(ID)          Sets the internal TankInfo to the TankInfo matching the 
'                       supplied key.  
' GetTank(ID)           Returns the Tank requested by the int arg ID
' GetAllInfo()          Returns the Entire TankCollection from the repository           
' GetAllByFacilityID(nFacilityId)   Returns a TankCollection of only those Tank objects 
'                                       corresponding to the FacilityID
' Add(ID)               Adds the Tank identified by arg ID to the 
'                           internal TankCollection
' Add(TankInfo)         Adds the Tank passed as the argument to the internal 
'                           TankCollection
' Remove(ID)            Removes the Tank identified by arg ID from the internal 
'                           TankCollection
' TankTable()           Returns a datatable containing all columns for the Tank
'                           objects in the internal TankCollection.
' colIsDirty()          Returns a boolean indicating whether any of the TankInfo
'                           objects in the TankCollection has been modified since the
'                           last time it was retrieved from/saved to the repository.
' Flush()               Marshalls all modified/added TankInfo objects in the 
'                            TankCollection to the repository.
' Save()                Marshalls the internal TankInfo object to the repository.
' Save(TankInfo)        Marshalls the TankInfo object to the repository as supplied.
' ValidateData()        Validates the TankInfo Object for the Business Rules
' TankTable()           Returns a datatable containing rows corresponding to the 
'                           TankInfo objects in the internal TankCollection.
' GridTankTable         Returns the datatable to be shown in the DataGrid using the 
'                           v_TANK_DISPLAY_DATA view
'-------------------------------------------------------------------------------
'
' TODO - check properties and operations against list.
'

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pTank
#Region "Private Member Variables"
        Private oFacilityInfo As MUSTER.Info.FacilityInfo
        Private WithEvents oTankCompartment As MUSTER.BusinessLogic.pCompartment
        Public passYesToTOSTOSIToPipes As Boolean = False
        Private WithEvents oTankInfo As MUSTER.Info.TankInfo
        'Private WithEvents colTank As MUSTER.Info.TankCollection
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
        Private WithEvents oProperty As MUSTER.BusinessLogic.pProperty
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private oTankDB As New MUSTER.DataAccess.TankDB
        Private nNewIndex As Integer = 0
        Private blnShowDeleted As Boolean = False
        Private nID As Int64 = -1
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Tank").ID
        Private bolValidationErrorOccurred As Boolean = False
        'Added By Elango on FEB 9 2005 
        'Public Event evtCAPStatusfromPipe(ByVal nVal As Integer)
        Private bolCheckCAPSTATE As Boolean = False
#End Region
#Region "Public Events"
        Public Event evtTankErr(ByVal strMessage As String)
        Public Event evtTanksChanged(ByVal bolValue As Boolean)
        Public Event evtTankChanged(ByVal bolValue As Boolean)
        'Public Event evtTankCommentsChanged(ByVal bolValue As Boolean)
        Public Event evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String)
        Public Event evtTankSaved(ByVal bolStatus As Boolean)
        'Public Event evtTankStatusChanged(ByVal oldStat As Integer, ByVal newStat As Integer, ByVal facID As Integer)

        'Public Event eTankStatus(ByRef dtTankStatus As DataTable)
        'Public Event eTankSecondaryOption(ByRef dtTankSecondaryOption As DataTable)
        'Public Event eTankReleaseDetection(ByRef dtTankReleaseDetection As DataTable)
        'Public Event eTankLDDisabled(ByVal isEnabled As Boolean)   'Recheck We do not need this.
        Public Event eDtPickEnableDisable(ByVal dtPickCP As Boolean, ByVal dtPickIntLining As Boolean)
        'Added By Elango 
        'Public Event ecmbTankCPType(ByVal BolState As Boolean)
        'Public Event edtPickCPInstalled(ByVal BolState As Boolean)
        'Public Event edtPickCPLastTested(ByVal BolState As Boolean)
        'Public Event edtPickInteriorLiningInstalled(ByVal BolState As Boolean)
        'Public Event edtPickLastInteriorLinningInspection(ByVal BolState As Boolean)
        'Public Event edtPickCPLastTestedMessage(ByVal strMessage As String)
        'Public Event edtPickLastInteriorLinningInspectionMessage(ByVal strMessage As String)
        'Public Event edtPickTankTightnessTestMessage(ByVal strMessage As String)
        'Public Event edtPickTankTightnessTest(ByVal BolState As Boolean)
        'Public Event echkTankDrpTubeInvControl(ByVal BolState As Boolean)
        'Public Event ecmbTankStatustext(ByVal strtext As String)
        'Public Event edtPickDatePlacedInServiceFocus()
        'Public Event edtPickPlannedInstallation(ByVal BolState As Boolean)
        'Public Event edtPickLastUsed(ByVal BolState As Boolean)
        'Public Event edGridCompartmentsStatus(ByVal BolState As Boolean)
        'Public Event ecmbTankInertFill(ByVal BolState As Boolean)
        'Public Event ecmbTankClosureType(ByVal BolState As Boolean)
        'Public Event edtPickLastUsedFocus()
        'Public Event evtCAPStatusfromTank(ByVal facID As Integer)
        'Public Event evtPipeCommentsChanged(ByVal bolValue As Boolean)
        'added by kiran
        'Public Event evtTankColFac(ByVal facID As Integer, ByVal tankCol As MUSTER.Info.TankCollection)
        'Public Event evtCompartmentCol(ByVal TankID As Integer, ByVal CompartmentCol As MUSTER.Info.CompartmentCollection, ByVal FacId As Integer)
        'Public Event evtCommentsCol(ByVal TankID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection, ByVal FacId As Integer)
        'Public Event evtPipeCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal tankID As Integer, ByVal facID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        'end changes
        'Public Event evtTankInfoFac(ByVal tankInfo As MUSTER.Info.TankInfo, ByVal strDesc As String)
        'Public Event evtTankInfoTankID(ByVal tnkID As Integer)
        'Public Event evtCompInfoTank(ByVal compartmentInfo As MUSTER.Info.CompartmentInfo, ByVal strDesc As String)
        'Public Event evtPipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String)
        'Public Event evtFacilityInfoTankCol(ByRef colTnk As MUSTER.Info.TankCollection)
        'Public Event evtFacilityInfoTankColByFacilityID(ByVal facID As Integer, ByRef colTnk As MUSTER.Info.TankCollection)
        'Public Event evtTankChangeKey(ByVal oldID As Integer, ByVal newID As Integer)
#End Region
#Region "Constructors"
        Public Sub New(Optional ByRef facInfo As MUSTER.Info.FacilityInfo = Nothing)
            If facInfo Is Nothing Then
                oFacilityInfo = New MUSTER.Info.FacilityInfo
            Else
                oFacilityInfo = facInfo
            End If
            oTankInfo = New MUSTER.Info.TankInfo
            'colTank = New MUSTER.Info.TankCollection
            oTankCompartment = New MUSTER.BusinessLogic.pCompartment(oTankInfo)
            oComments = New MUSTER.BusinessLogic.pComments
            oProperty = New MUSTER.BusinessLogic.pProperty
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property TankInfo() As MUSTER.Info.TankInfo
            Get
                Return oTankInfo
            End Get
            Set(ByVal Value As MUSTER.Info.TankInfo)
                oTankInfo = Value
            End Set
        End Property
        Public Property TankId() As Integer
            Get
                Return oTankInfo.TankId
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankId = value
            End Set
        End Property
        Public Property TankIndex() As Integer
            Get
                Return oTankInfo.TankIndex
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankIndex = value
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property FacilityId() As Integer
            Get
                Return oTankInfo.FacilityId
            End Get

            Set(ByVal value As Integer)
                oTankInfo.FacilityId = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TankStatus() As Integer
            Get
                Return oTankInfo.TankStatus
            End Get

            Set(ByVal value As Integer)
                Dim oldTnkStat As Integer = oTankInfo.TankStatus
                oTankInfo.TankStatus = value
                CheckTankStatus(oldTnkStat, oTankInfo.TankStatus)
                'RaiseEvent evtTankStatusChanged(oldTnkStat, oTankInfo.TankStatus, oTankInfo.FacilityId)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property TankStatusOriginal() As Integer
            Get
                Return oTankInfo.OriginalTankStatus
            End Get
        End Property
        Public Property DateReceived() As Date
            Get
                Return oTankInfo.DateReceived
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateReceived = value
            End Set
        End Property
        Public Property Manifold() As Boolean
            Get
                Return oTankInfo.Manifold
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.Manifold = value
            End Set
        End Property
        Public Property Compartment() As Boolean
            Get
                Return oTankInfo.Compartment
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.Compartment = value
            End Set
        End Property
        Public Property TankCapacity() As Integer
            Get
                Return oTankInfo.TankCapacity
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankCapacity = value
            End Set
        End Property
        Public Property Substance() As Integer
            Get
                Return oTankInfo.Substance
            End Get

            Set(ByVal value As Integer)
                oTankInfo.Substance = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property CASNumber() As Integer
            Get
                Return oTankInfo.CASNumber
            End Get

            Set(ByVal value As Integer)
                oTankInfo.CASNumber = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property SubstanceCommentsID() As Integer
            Get
                Return oTankInfo.SubstanceCommentsID
            End Get

            Set(ByVal value As Integer)
                oTankInfo.SubstanceCommentsID = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property DateLastUsed() As Date
            Get
                Return oTankInfo.DateLastUsed
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateLastUsed = value
                'colTank(oTankInfo.TankId) = oTankInfo
                'CheckDateLastUsed(oTankInfo.DateLastUsed)
            End Set
        End Property
        Public Property DateClosureReceived() As Date
            Get
                Return oTankInfo.DateClosureReceived
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateClosureReceived = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property DateClosed() As Date
            Get
                Return oTankInfo.DateClosed
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateClosed = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property ClosureStatusDesc() As Integer
            Get
                Return oTankInfo.ClosureStatusDesc
            End Get

            Set(ByVal value As Integer)
                oTankInfo.ClosureStatusDesc = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property ClosureType() As Integer
            Get
                Return oTankInfo.ClosureType
            End Get
            Set(ByVal Value As Integer)
                oTankInfo.ClosureType = Value
            End Set
        End Property
        Public Property InertMaterial() As Integer
            Get
                Return oTankInfo.InertMaterial
            End Get

            Set(ByVal value As Integer)
                oTankInfo.InertMaterial = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TankMatDesc() As Integer
            Get
                Return oTankInfo.TankMatDesc
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankMatDesc = value
                'CheckTankMatDesc(oTankInfo.TankMatDesc)
            End Set
        End Property
        Public Property TankModDesc() As Integer
            Get
                Return oTankInfo.TankModDesc
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankModDesc = value
                'colTank(oTankInfo.TankId) = oTankInfo
                'CheckTankModDesc(oTankInfo.TankModDesc)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property TankOtherMaterial() As Integer
            Get
                Return oTankInfo.TankOtherMaterial
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankOtherMaterial = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property OverFillInstalled() As Boolean
            Get
                Return oTankInfo.OverFillInstalled
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.OverFillInstalled = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property SpillInstalled() As Boolean
            Get
                Return oTankInfo.SpillInstalled
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.SpillInstalled = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property DateSpillInstalled() As Date
            Get
                Return oTankInfo.DateSpillInstalled
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateSpillInstalled = value
            End Set
        End Property
        Public Property DateSpillTested() As Date
            Get
                Return oTankInfo.DateSpillTested
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateSpillTested = value
            End Set
        End Property
        Public Property DateOverfillInstalled() As Date
            Get
                Return oTankInfo.DateOverfillInstalled
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateOverfillInstalled = value
            End Set
        End Property
        Public Property DateOverfillTested() As Date
            Get
                Return oTankInfo.DateOverfillTested
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateOverfillTested = value
            End Set
        End Property

        'Public Property DateTankSecInsp() As Date
        '   Get
        '      Return oTankInfo.DateTankSecInsp
        ' End Get

        'Set(ByVal value As Date)
        '   oTankInfo.DateTankSecInsp = value
        ' End Set
        ' End Property

        Public Property DateTankElecInsp() As Date
            Get
                Return oTankInfo.DateTankElecInsp
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateTankElecInsp = value
            End Set
        End Property
        Public Property DateATGInsp() As Date
            Get
                Return oTankInfo.DateATGInsp
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateATGInsp = value
            End Set
        End Property
        Public Property LicenseedID() As Integer
            Get
                Return oTankInfo.LicenseeID
            End Get

            Set(ByVal value As Integer)
                oTankInfo.LicenseeID = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property ContractorID() As Integer
            Get
                Return oTankInfo.ContractorID
            End Get

            Set(ByVal value As Integer)
                oTankInfo.ContractorID = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property DateSigned() As Date
            Get
                Return oTankInfo.DateSigned
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateSigned = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property DateInstalledTank() As Date
            Get
                Return oTankInfo.DateInstalledTank
            End Get

            Set(ByVal value As Date)
                oTankInfo.DateInstalledTank = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property SmallDelivery() As Boolean
            Get
                Return oTankInfo.SmallDelivery
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.SmallDelivery = value
            End Set
        End Property
        Public Property TankEmergen() As Boolean
            Get
                Return oTankInfo.TankEmergen
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.TankEmergen = value
                Me.PopulateTankReleaseDetection(oTankInfo.TankModDesc, 0)
                If oTankInfo.TankEmergen Then
                    'set pipe release detection group 1 to deferred (248)
                    ' group 2 blank (0) and clear alld date if pipe type not safe suction
                    For Each pipe As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                        If pipe.PipeTypeDesc <> 267 Then
                            pipe.PipeLD = 248
                            pipe.ALLDType = 0
                            pipe.ALLDTestDate = CDate("01/01/0001")
                        End If
                    Next
                End If
            End Set
        End Property
        Public Property PlannedInstDate() As Date
            Get
                Return oTankInfo.PlannedInstDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.PlannedInstDate = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property LastTCPDate() As Date
            Get
                Return oTankInfo.LastTCPDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.LastTCPDate = value
                'CheckLastTCPDate(oTankInfo.LastTCPDate)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property LinedInteriorInstallDate() As Date
            Get
                Return oTankInfo.LinedInteriorInstallDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.LinedInteriorInstallDate = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property LinedInteriorInspectDate() As Date
            Get
                Return oTankInfo.LinedInteriorInspectDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.LinedInteriorInspectDate = value
                'CheckLinedInteriorInspectDate(oTankInfo.LinedInteriorInspectDate)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property LinedInteriorInspectDateOriginal() As Date
            Get
                Return oTankInfo.LinedInteriorInspectDateOriginal
            End Get
        End Property
        Public Property TCPInstallDate() As Date
            Get
                Return oTankInfo.TCPInstallDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.TCPInstallDate = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TTTDate() As Date
            Get
                Return oTankInfo.TTTDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.TTTDate = value
                'CheckTTTDate(oTankInfo.TTTDate)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property TTTDateOriginal() As Date
            Get
                Return oTankInfo.TTTDateOriginal
            End Get
        End Property
        Public Property TankLD() As Integer
            Get
                Return oTankInfo.TankLD
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankLD = value
                'colTank(oTankInfo.TankId) = oTankInfo
                'CheckTankLD(oTankInfo.TankLD)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property OverFillType() As Integer
            Get
                Return oTankInfo.OverFillType
            End Get

            Set(ByVal value As Integer)
                oTankInfo.OverFillType = value

            End Set
        End Property
        Public Property RevokeReason() As Integer
            Get
                Return oTankInfo.RevokeReason
            End Get

            Set(ByVal value As Integer)
                oTankInfo.RevokeReason = value

            End Set
        End Property
        Public Property RevokeDate() As Date
            Get
                Return oTankInfo.RevokeDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.RevokeDate = value
            End Set
        End Property
        Public Property DatePhysicallyTagged() As Date
            Get
                Return oTankInfo.DatePhysicallyTagged
            End Get

            Set(ByVal value As Date)
                oTankInfo.DatePhysicallyTagged = value
            End Set
        End Property
        Public Property Prohibition() As Boolean
            Get
                Return oTankInfo.Prohibition
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.Prohibition = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TightFillAdapters() As Boolean
            Get
                Return oTankInfo.TightFillAdapters
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.TightFillAdapters = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property DropTube() As Boolean
            Get
                Return oTankInfo.DropTube
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.DropTube = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TankCPType() As Integer
            Get
                Return oTankInfo.TankCPType
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankCPType = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property PlacedInServiceDate() As Date
            Get
                Return oTankInfo.PlacedInServiceDate
            End Get

            Set(ByVal value As Date)
                oTankInfo.PlacedInServiceDate = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TankTypes() As Integer
            Get
                Return oTankInfo.TankTypes
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankTypes = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TankLocationDescription() As String
            Get
                Return oTankInfo.TankLocationDescription
            End Get

            Set(ByVal value As String)
                oTankInfo.TankLocationDescription = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property TankManufacturer() As Integer
            Get
                Return oTankInfo.TankManufacturer
            End Get

            Set(ByVal value As Integer)
                oTankInfo.TankManufacturer = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property Deleted() As Boolean
            Get
                Return oTankInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.Deleted = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oTankInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oTankInfo.IsDirty = value
                'colTank(oTankInfo.TankId) = oTankInfo
            End Set
        End Property
        'Public Property CheckCAPSTATE() As Boolean

        '    Get
        '        Return bolCheckCAPSTATE
        '    End Get

        '    Set(ByVal value As Boolean)
        '        bolCheckCAPSTATE = value
        '    End Set
        'End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xTankInfo As MUSTER.Info.TankInfo
                For Each xTankInfo In oFacilityInfo.TankCollection.Values
                    If xTankInfo.IsDirty Then
                        'MsgBox("BLL F:" + oFacilityInfo.ID.ToString + " TI:" + xTankInfo.TankIndex.ToString + " TID:" + xTankInfo.TankId.ToString)
                        Return True
                        Exit Property
                    End If
                Next
                If oTankCompartment.colIsDirty Then
                    Return True
                    Exit Property
                End If
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property ShowDeleted() As Boolean
            Get
                Return blnShowDeleted
            End Get

            Set(ByVal value As Boolean)
                blnShowDeleted = value
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
        Public Property FacilityInfo() As MUSTER.Info.FacilityInfo
            Get
                Return oFacilityInfo
            End Get
            Set(ByVal Value As MUSTER.Info.FacilityInfo)
                oFacilityInfo = Value
            End Set
        End Property
        Public Property FacCapStatus() As Integer
            Get
                Return oTankInfo.FacCapStatus
            End Get
            Set(ByVal Value As Integer)
                oTankInfo.FacCapStatus = Value
            End Set
        End Property
        ' Collections
        Public ReadOnly Property TankCollection() As MUSTER.Info.TankCollection
            Get
                Return oFacilityInfo.TankCollection
            End Get
        End Property
        Public Property Pipes() As MUSTER.BusinessLogic.pPipe
            Get
                Return oTankCompartment.Pipes
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pPipe)
                oTankCompartment.Pipes = Value
            End Set
        End Property
        Public Property Compartments() As MUSTER.BusinessLogic.pCompartment
            Get
                Return oTankCompartment
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pCompartment)
                oTankCompartment = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oTankInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oTankInfo.CreatedBy = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oTankInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oTankInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oTankInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oTankInfo.ModifiedOn
            End Get
        End Property
        Public ReadOnly Property FacilityPowerOff() As Boolean
            Get
                Return oTankInfo.FacilityPowerOff
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Function MessageBox(ByVal str As String)
            Dim a As MsgBoxResult = MsgBox(str, MsgBoxStyle.YesNo)

            If a = MsgBoxResult.No Then
                Me.passYesToTOSTOSIToPipes = True
            End If

            Return a
        End Function

        Public Sub RetrieveAll(ByVal ownerID As Integer, Optional ByVal [Module] As String = "", Optional ByVal showDeleted As Boolean = False, Optional ByVal facID As Int64 = 0, Optional ByVal tankID As Int64 = 0)
            Try
                Dim ds As New DataSet
                ds = oTankDB.DBGetDS(ownerID, [Module], showDeleted, facID, tankID)
                If tankID = 0 Then
                    ' 0 - Facilities + addresses
                    ' 1 - Tanks
                    ' 2 - Compartments
                    ' 3 - Pipes
                    ds.Tables(0).TableName = "Facilities"
                    ds.Tables(1).TableName = "Tanks"
                    ds.Tables(2).TableName = "Compartments"
                    ds.Tables(3).TableName = "Pipes"
                Else
                    ' 0 - Tanks
                    ' 1 - Compartments
                    ' 2 - Pipes
                    ds.Tables(0).TableName = "Tanks"
                    ds.Tables(1).TableName = "Compartments"
                    ds.Tables(2).TableName = "Pipes"
                End If
                ' Tanks
                Load(oFacilityInfo, ds, [Module])
                ds = Nothing
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Load(ByRef FacInfo As MUSTER.Info.FacilityInfo, ByRef ds As DataSet, ByVal [Module] As String)
            Dim dr As DataRow
            oFacilityInfo = FacInfo
            Try
                If ds.Tables("Tanks").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Tanks").Select("FACILITY_ID = " + oFacilityInfo.ID.ToString)
                        oTankInfo = New MUSTER.Info.TankInfo(dr)
                        'RaiseEvent evtTankInfoFac(oTankInfo, "ADD")
                        oTankInfo.FacCapStatus = oFacilityInfo.CapStatus
                        If oTankInfo.TankStatus = 426 Then
                            oTankInfo.POU = True
                            If Date.Compare(oTankInfo.DateLastUsed, CDate("12/22/1988")) >= 0 Then
                                oTankInfo.NonPre88 = True
                            End If
                        End If
                        If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                            oTankInfo.FacilityPowerOff = True
                        Else
                            oTankInfo.FacilityPowerOff = False
                        End If
                        oFacilityInfo.TankCollection.Add(oTankInfo)
                        'oComments.Load(ds, [Module], nEntityTypeID, oTankInfo.TankId)
                        ds.Tables("Tanks").Rows.Remove(dr)
                        oTankCompartment.Load(oTankInfo, ds, [Module])
                        'For Each oPipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                        '    If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                        '        oPipeInfoLocal.FacilityPowerOff = True
                        '    Else
                        '        oPipeInfoLocal.FacilityPowerOff = False
                        '    End If
                        'Next
                        ' if module is inspection and tank status = pou and
                        ' has no pipes, remove tank from collection
                        ' we do not need to show pou tanks unless they have pipes
                        If [Module].Trim.ToUpper = "INSPECTION" And _
                                oTankInfo.TankStatus = 426 And _
                                oTankInfo.pipesCollection.Count = 0 Then
                            oFacilityInfo.TankCollection.Remove(oTankInfo)
                        End If
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function RetrieveTank(ByVal tnkId As Integer) As MUSTER.Info.TankInfo
            Try
                oTankInfo = oFacilityInfo.TankCollection.Item(tnkId)
                ' test for dataage here...

                If oTankInfo Is Nothing Then
                    Add(tnkId)
                Else
                    If oTankInfo.IsDirty = False And oTankInfo.IsAgedData = True Then
                        oFacilityInfo.TankCollection.Remove(oTankInfo)
                        Add(tnkId)
                    End If
                End If
                SetInfoInChild()
                Return oTankInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Retrieve(ByRef FacInfo As MUSTER.Info.FacilityInfo, ByVal facID As Int64, Optional ByVal tnkID As Int64 = 0, Optional ByVal ShowDeleted As Boolean = False, Optional ByVal tnkIDDesc As String = "NONE") As MUSTER.Info.TankInfo
            Try
                oFacilityInfo = FacInfo
                Dim oTankInfoLocal As MUSTER.Info.TankInfo
                If tnkID = 0 Then
                    ' retrieve tanks related to facilityId and tankId
                    If tnkIDDesc = "NEW" Then
                        ' Check in Tank Collection
                        oTankInfo = oFacilityInfo.TankCollection.Item(tnkID)
                        ' zxc test for dataage here...

                        If oTankInfo Is Nothing Then
                            Add(tnkID)
                        Else
                            If oTankInfo.IsDirty = False And oTankInfo.IsAgedData = True Then
                                oFacilityInfo.TankCollection.Remove(oTankInfo)
                                'RaiseEvent evtTankInfoFac(oTankInfo, "REMOVE")
                                'Add(tnkID, ShowDeleted)
                                RetrieveAll(oFacilityInfo.OwnerID, , ShowDeleted, oFacilityInfo.ID, tnkID)
                            End If
                        End If
                        oTankCompartment.Retrieve(oTankInfo, oTankInfo.TankId, ShowDeleted)
                        'oTankCompartment.TankSiteID = oTankInfo.TankIndex
                        'If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                        '    oTankInfo.FacilityPowerOff = True
                        'Else
                        '    oTankInfo.FacilityPowerOff = False
                        'End If
                        'For Each oPipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                        '    If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                        '        oPipeInfoLocal.FacilityPowerOff = True
                        '    Else
                        '        oPipeInfoLocal.FacilityPowerOff = False
                        '    End If
                        'Next
                    Else
                        ' ------------------------------------------------------------
                        ' No data aging is done here at this point.  It should be done when
                        ' this process is re-written to return all tanks for a facility
                        ' rather than just one.
                        ' ------------------------------------------------------------
                        '
                        ' retrieve all tanks related to facilityId
                        Dim BolCompRetrieved As Boolean = False
                        ' check in collection
                        For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                            If oTankInfoLocal.TankId = tnkID Then
                                BolCompRetrieved = True
                                oTankInfo = oTankInfoLocal
                                oTankCompartment.Retrieve(oTankInfo, oTankInfoLocal.TankId, ShowDeleted)
                            End If
                        Next
                        'Get From DB
                        If Not BolCompRetrieved Then
                            RetrieveAll(oFacilityInfo.OwnerID, , ShowDeleted, facID, )
                            'oFacilityInfo.TankCollection = oTankDB.DBGetByFacilityID(facID, ShowDeleted)
                            'If oFacilityInfo.TankCollection.Count > 0 Then
                            'Dim dtTempDate As Date = "12-22-88"
                            'For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                            'oTankInfo = oTankInfoLocal
                            'If oTankInfo.TankStatus = 426 Then
                            '    oTankInfo.POU = True
                            '    If Date.Compare(oTankInfo.DateLastUsed, dtTempDate) >= 0 Then
                            '        oTankInfo.NonPre88 = True
                            '    End If
                            'End If
                            'oTankCompartment.Retrieve(oTankInfo, oTankInfo.TankId, ShowDeleted)
                            'oTankCompartment.FacilityId = oTankInfo.FacilityId
                            'oTankCompartment.TankSiteID = oTankInfo.TankIndex
                            'If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                            '    oTankInfo.FacilityPowerOff = True
                            'Else
                            '    oTankInfo.FacilityPowerOff = False
                            'End If
                            'For Each oPipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                            '    If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                            '        oPipeInfoLocal.FacilityPowerOff = True
                            '    Else
                            '        oPipeInfoLocal.FacilityPowerOff = False
                            '    End If
                            'Next
                            'Next
                            'Else
                            'If we didn't get tanks related to this facility id then pass Nothing... 
                            'Need to check again ... 
                            'End If
                        End If
                    End If
                Else
                    ' retrieve tank related to tankId
                    oTankInfo = oFacilityInfo.TankCollection.Item(tnkID)
                    ' zxc test for dataage here...

                    If oTankInfo Is Nothing Then
                        RetrieveAll(oFacilityInfo.OwnerID, , ShowDeleted, oFacilityInfo.ID, tnkID)
                        'Add(tnkID)
                    End If
                    oTankCompartment.Retrieve(oTankInfo, oTankInfo.TankId, ShowDeleted)
                    'oTankCompartment.TankSiteID = oTankInfo.TankIndex
                    'If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                    '    oTankInfo.FacilityPowerOff = True
                    'Else
                    '    oTankInfo.FacilityPowerOff = False
                    'End If
                    'For Each oPipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                    '    If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                    '        oPipeInfoLocal.FacilityPowerOff = True
                    '    Else
                    '        oPipeInfoLocal.FacilityPowerOff = False
                    '    End If
                    'Next
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'oComments.Clear()
            'If Not oTankInfo Is Nothing Then
            'oComments.GetByModule("", nEntityTypeID, oTankInfo.TankId)
            'End If
            SetInfoInChild()
            Return oTankInfo
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oTankDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal validateCapDates As Boolean = True, Optional ByVal strExcludeOCE As String = "", Optional ByVal bolReplacementTank As Boolean = False, Optional ByVal bolSaveToInspectionMirror As Boolean = False) As Boolean
            Dim oldID As Integer
            Dim oldStat As Integer
            Dim CompInfoLocal As MUSTER.Info.CompartmentInfo
            Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Try
                Dim strModule As String = String.Empty

                If moduleID = 612 Then
                    strModule = "REGISTRATION"
                ElseIf moduleID = 613 Then
                    strModule = "CAE"
                ElseIf moduleID = 891 Then
                    strModule = "CLOSURE"
                ElseIf moduleID = 615 Then
                    strModule = "INSPECTION"
                End If

                If Not bolValidated And Not oTankInfo.Deleted And Not bolDelete Then
                    If Not Me.ValidateData(validateCapDates, strModule) Then
                        Exit Function
                    End If
                End If

                If Not (oTankInfo.TankId < 0 And oTankInfo.Deleted) Then
                    oldID = oTankInfo.TankId
                    oldStat = oTankInfo.OriginalTankStatus
                    ' TODO - do i need to check for tos/tosi rules even when deleting the tank
                    CheckTOSTOSIRules(oTankInfo, moduleID, staffID, returnVal, strExcludeOCE)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                    ' flags
                    If strModule.ToUpper = "REGISTRATION" Or strModule.ToUpper = "CLOSURE" Or strModule.ToUpper = "INSPECTION" Or strModule.ToUpper = "CAE" Then
                        ' if tank status was changed to tos
                        Dim flags As New MUSTER.BusinessLogic.pFlag
                        If oTankInfo.TankStatus = 425 Then ' TOS
                            If oTankInfo.OriginalTankStatus <> 425 Then ' NOT TOS
                                flags.RetrieveFlags(oTankInfo.FacilityId, 6, , , , , "SYSTEM", "TOS Tank")
                                If flags.FlagsCol.Count <= 0 Then
                                    flags.Add(New MUSTER.Info.FlagInfo(0, _
                                        oTankInfo.FacilityId, _
                                        6, _
                                        "TOS Tanks for Facility - " + oTankInfo.FacilityId.ToString, _
                                        False, _
                                        CDate("01/01/0001"), _
                                        "REGISTRATION", _
                                        0, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        "SYSTEM", _
                                        "RED"))
                                    flags.Save()
                                End If
                            Else
                                flags.RetrieveFlags(oTankInfo.FacilityId, 6, , , , , "SYSTEM", "TOS Tank")
                                If flags.FlagsCol.Count <= 0 Then
                                    flags.Add(New MUSTER.Info.FlagInfo(0, _
                                        oTankInfo.FacilityId, _
                                        6, _
                                        "TOS Tanks for Facility - " + oTankInfo.FacilityId.ToString, _
                                        False, _
                                        CDate("01/01/0001"), _
                                        "REGISTRATION", _
                                        0, _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        String.Empty, _
                                        CDate("01/01/0001"), _
                                        CDate("01/01/0001"), _
                                        "SYSTEM", _
                                        "RED"))
                                    flags.Save()
                                End If
                            End If
                        Else
                            Dim bolNoTOSTanks As Boolean = True
                            Dim ds As DataSet = oTankDB.DBGetDS("SELECT COUNT(TANK_ID) AS TOSTANKS FROM TBLREG_TANK WHERE TANKSTATUS = 425 AND DELETED = 0 AND FACILITY_ID = " + oTankInfo.FacilityId.ToString)
                            If ds.Tables(0).Rows(0)(0) > 0 Then
                                bolNoTOSTanks = False
                            Else
                                bolNoTOSTanks = True
                            End If
                            'For Each tnk As MUSTER.Info.TankInfo In FacilityInfo.TankCollection.Values
                            '    If tnk.TankStatus = 425 Then
                            '        bolNoTOSTanks = False
                            '        Exit For
                            '    End If
                            'Next
                            If bolNoTOSTanks Then
                                flags.RetrieveFlags(oTankInfo.FacilityId, 6, , , , , "SYSTEM", "TOS Tank")
                                For Each flagInfo As MUSTER.Info.FlagInfo In flags.FlagsCol.Values
                                    flagInfo.Deleted = True
                                Next
                                If flags.FlagsCol.Count > 0 Then
                                    flags.Flush()
                                End If
                            End If
                        End If
                    End If
                    ' save tank
                    oTankDB.Put(oTankInfo, oFacilityInfo.CapStatus, moduleID, staffID, returnVal, strUser, bolReplacementTank, bolSaveToInspectionMirror)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    ' #2753
                    If TankInfo.OriginalTankStatus = 429 And TankInfo.TankStatus = 424 Then
                        Dim dsPipeIDs As DataSet = oTankDB.DBGetDS("SELECT [PIPE ID], [COMPARTMENT NUMBER] FROM V_PIPES_DISPLAY_DATA WHERE FACILITY_ID = " + TankInfo.FacilityId.ToString + " and [TANK ID] = " + TankInfo.TankId.ToString)
                        If dsPipeIDs.Tables.Count > 0 Then
                            Dim compInfo As MUSTER.Info.CompartmentInfo
                            Dim pipeInfo As MUSTER.Info.PipeInfo

                            ' if tank has pipes
                            If dsPipeIDs.Tables(0).Rows.Count > 0 Then
                                ' if there are pipes in collection
                                If TankInfo.pipesCollection.Count <> 0 Then
                                    'For Each pipeInfo In TankInfo.pipesCollection.Values
                                    'If dsPipeIDs.Tables(0).Select("[PIPE ID] = " + pipeInfo.PipeID.ToString + " AND [COMPARTMENT NUMBER] = " + pipeInfo.CompartmentNumber.ToString).Length > 0 Then
                                    'If pipeInfo.PipeStatusDesc <> 426 Then
                                    '  pipeInfo.PipeStatusDesc = 424 ' ciu
                                    'End If

                                    ' dsPipeIDs.Tables(0).Rows.Remove(dsPipeIDs.Tables(0).Select("[PIPE ID] = " + pipeInfo.PipeID.ToString + " AND [COMPARTMENT NUMBER] = " + pipeInfo.CompartmentNumber.ToString)(0))
                                    'End If
                                    '      Next
                                End If


                                ' load pipes into collection
                                If dsPipeIDs.Tables(0).Rows.Count > 0 Then
                                    Dim ds As DataSet = oTankDB.DBGetDS(FacilityInfo.OwnerID, strModule, ShowDeleted, TankInfo.FacilityId, TankInfo.TankId)
                                    ' 0 - Tanks
                                    ' 1 - Compartments
                                    ' 2 - Pipes
                                    ds.Tables(0).TableName = "Tanks"
                                    ds.Tables(1).TableName = "Compartments"
                                    ds.Tables(2).TableName = "Pipes"

                                    ds.Tables.Remove("Tanks")
                                    ' if compartment exists in collection, delete the row from dataset
                                    For Each compInfo In TankInfo.CompartmentCollection.Values
                                        If ds.Tables("Compartments").Select("COMPARTMENT_NUMBER = " + compInfo.COMPARTMENTNumber.ToString).Length > 0 Then
                                            ds.Tables("Compartments").Rows.Remove(ds.Tables("Compartments").Select("COMPARTMENT_NUMBER = " + compInfo.COMPARTMENTNumber.ToString)(0))
                                        End If
                                    Next
                                    ' if compartment exists in collection, delete the row from dataset
                                    For Each pipeInfo In TankInfo.pipesCollection.Values
                                        If ds.Tables("Pipes").Select("COMPARTMENT_NUMBER = " + pipeInfo.CompartmentNumber.ToString + " AND PIPE_ID = " + pipeInfo.PipeID.ToString).Length > 0 Then
                                            ds.Tables("Pipes").Rows.Remove(ds.Tables("Pipes").Select("COMPARTMENT_NUMBER = " + pipeInfo.CompartmentNumber.ToString + " AND PIPE_ID = " + pipeInfo.PipeID.ToString)(0))
                                        End If
                                    Next
                                    ' load compartments
                                    If ds.Tables("Compartments").Rows.Count > 0 Then
                                        Compartments.Load(TankInfo, ds, strModule)
                                    Else
                                        Pipes.Load(TankInfo, ds, strModule)
                                    End If
                                    ds = Nothing
                                End If

                                ' For Each pipeInfo In TankInfo.pipesCollection.Values

                                'If pipeInfo.PipeStatusDesc <> 426 Then
                                'pipeInfo.PipeStatusDesc = 424 ' ciu
                                ' End If

                                '  Next

                            End If
                        End If
                    End If

                    If oTankInfo.TankStatus = 426 Then
                        oTankInfo.POU = True
                        If Date.Compare(oTankInfo.DateLastUsed, CDate("12/22/1988")) >= 0 Then
                            oTankInfo.NonPre88 = True
                        End If
                    End If
                    If oFacilityInfo.CapStatusOriginal <> oFacilityInfo.CapStatus Then
                        oFacilityInfo.CapStatusOriginal = oFacilityInfo.CapStatus
                        'RaiseEvent evtCAPStatusfromTank(oFacilityInfo.ID)
                    End If
                    oTankInfo.FacCapStatus = oFacilityInfo.CapStatus
                    ' #2753 tosi to ciu change pipe status to ciu
                    'If oTankInfo.OriginalTankStatus = 429 And oTankInfo.TankStatus = 424 Then
                    '    For Each pipeInfo As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                    '        pipeInfo.PipeStatusDesc = 424
                    '    Next
                    'End If
                    oTankInfo.Archive()
                    RaiseEvent evtTankSaved(True)
                    If oldID <> oTankInfo.TankId Then
                        If Not bolValidated Then
                            oFacilityInfo.TankCollection.ChangeKey(oldID, oTankInfo.TankId.ToString)
                            'RaiseEvent evtTankChangeKey(oldID, oTankInfo.TankId)
                        End If
                        Dim IDs As New Collection
                        Dim delIDs As New Collection
                        Dim index As Integer
                        For Each CompInfoLocal In oTankInfo.CompartmentCollection.Values
                            If CompInfoLocal.TankId = oldID Then
                                IDs.Add(CompInfoLocal.ID)
                            End If
                        Next
                        If Not (IDs Is Nothing) Then
                            Dim colkey As String = String.Empty
                            For index = 1 To IDs.Count
                                colkey = CType(IDs.Item(index), String)
                                CompInfoLocal = oTankInfo.CompartmentCollection.Item(colkey)
                                If CompInfoLocal.Capacity = 0 And _
                                    CompInfoLocal.Substance = 0 And _
                                    CompInfoLocal.CCERCLA = 0 And _
                                    CompInfoLocal.FuelTypeId = 0 And _
                                    oTankInfo.Compartment And _
                                    IDs.Count > 1 Then
                                    oTankCompartment.Remove(colkey)
                                Else
                                    oTankCompartment.ChangeCompartmentNumberKey(, oTankInfo.TankId, CompInfoLocal)
                                End If
                            Next
                        End If
                    Else

                    End If
                    If bolDelete Then
                        Dim oCompLocal As MUSTER.Info.CompartmentInfo
                        For Each oCompLocal In oTankInfo.CompartmentCollection.Values
                            oCompLocal.Deleted = True
                            oCompLocal.ModifiedBy = oTankInfo.ModifiedBy
                        Next
                    End If
                    oTankInfo.Archive()
                    oTankInfo.IsDirty = False
                    'If Not bolValidated And bolDelete Then
                    'SetInfoInChild()
                    'oTankCompartment.Flush()
                    'End If
                    SetInfoInChild()
                    Dim userID As String = String.Empty
                    If oTankInfo.TankId <= 0 Then
                        userID = oTankInfo.CreatedBy
                    Else
                        userID = oTankInfo.ModifiedBy
                    End If
                    oTankCompartment.Flush(moduleID, staffID, returnVal, userID, bolSaveToInspectionMirror)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                End If
                If Not bolValidated And bolDelete Then
                    If oTankInfo.Deleted Then
                        ' check if other tanks are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oTankInfo.TankId Then
                            If strPrev = oTankInfo.TankId Then
                                RaiseEvent evtTankErr("Tank " + oTankInfo.TankIndex.ToString + " deleted")
                                oFacilityInfo.TankCollection.Remove(oTankInfo)
                                'RaiseEvent evtTankInfoFac(oTankInfo, "REMOVE")
                                If bolDelete Then
                                    oTankInfo = New MUSTER.Info.TankInfo
                                Else
                                    oTankInfo = Me.Retrieve(oFacilityInfo, 0, oTankInfo.FacilityId)
                                End If
                            Else
                                RaiseEvent evtTankErr("Tank " + oTankInfo.TankIndex.ToString + " deleted")
                                oFacilityInfo.TankCollection.Remove(oTankInfo)
                                'RaiseEvent evtTankInfoFac(oTankInfo, "REMOVE")
                                oTankInfo = Me.Retrieve(oFacilityInfo, strPrev, oTankInfo.FacilityId)
                            End If
                        Else
                            RaiseEvent evtTankErr("Tank " + oTankInfo.TankIndex.ToString + " deleted")
                            oFacilityInfo.TankCollection.Remove(oTankInfo)
                            'RaiseEvent evtTankInfoFac(oTankInfo, "REMOVE")
                            oTankInfo = Me.Retrieve(oFacilityInfo, strNext, oTankInfo.FacilityId)
                        End If
                    End If
                End If
                RaiseEvent evtTankChanged(Me.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
                Return False
            End Try
        End Function
        Public Function DeleteTank(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String) As Boolean
            ' check if tank has any pipes in coll
            Try
                'oTankDB.DeleteTank(oTankInfo.TankId)
                'Dim oCompLocal As MUSTER.Info.CompartmentInfo
                'Dim oPipeLocal As MUSTER.Info.PipeInfo
                'For Each oCompLocal In oTankInfo.CompartmentCollection.Values
                '    For Each oPipeLocal In oTankInfo.pipesCollection.Values
                '        RaiseEvent evtTankErr("The Specified tank has associated Pipe(s). Delete Pipe(s) before deleting the tank")
                '        Return False
                '    Next
                'Next
                'For Each oPipeLocal In oTankCompartment.PipeCollection.Values
                '    If oPipeLocal.TankID = oTankInfo.TankId Then
                '        RaiseEvent evtTankErr("The Specified tank has associated Pipe(s). Delete Pipe(s) before deleting the tank")
                '        Return False
                '    End If
                'Next
                If oTankInfo.TankId > 0 Then
                    Dim ds As DataSet
                    ds = oTankDB.DBGetDS("EXEC spCheckDependancy NULL,NULL," + oTankInfo.TankId.ToString + ",0,NULL")
                    If ds.Tables(0).Rows(0)(0) Then
                        RaiseEvent evtTankErr(IIf(ds.Tables(0).Rows(0)("MSG") Is DBNull.Value, "Tank has dependants", ds.Tables(0).Rows(0)("MSG")))
                        Return False
                    End If
                End If

                ' tank does not have pipe(s), delete tank
                oTankInfo.Deleted = True
                Return Me.Save(moduleID, staffID, returnVal, "", True, True)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
        Public Function ValidateData(ByVal validateCapDates As Boolean, Optional ByVal [module] As String = "Registration", Optional ByRef strError As String = "", Optional ByVal returnString As Boolean = False) As Boolean
            'This function Self Validates the Object.
            ' TankStatus -> Currently in Use (CIU) = 424; 
            ' TankStatus -> Temporarily Out of Service (TOS) = 425;
            ' TankStatus -> Permanently Out of Use (POS) = 426;
            ' TankStatus -> Temporarily Out of Service Indefinitely (TOSI) = 429;
            ' TankStatus -> Unregulated = 430;
            ' TankStatus -> Permanent Closure Pending = 431;
            Try
                Dim errStr As String = ""
                Dim msgStr As String = ""
                Dim validateSuccess As Boolean = True
                Dim dtNullDate As Date
                Dim dttemp, dtValidDate As Date
                Dim dtTodayPlus90Days As Date = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 4, CDate(Today.Month.ToString + "/1/" + Today.Year.ToString)))
                Select Case [module].Trim.ToUpper
                    Case "REGISTRATION", "INSPECTION"
                        If oTankInfo.TankId <> 0 Then
                            If oTankInfo.TankId < 0 Then
                                If oTankInfo.TankStatus = 0 Then
                                    errStr += "Tank Status cannot be empty" + vbCrLf
                                    validateSuccess = False
                                    Exit Select
                                End If
                            End If ' end If oTankInfo.TankId < 0
                            'status
                            'status = CIU - material, tank option, release detection
                            'status = POU - date last used < 12/22/1988
                            'status = TOS - date last used
                            If oTankInfo.TankStatus = 424 Or oTankInfo.TankStatus = 429 Then 'CIU OR TOSI
                                If oTankInfo.TankMatDesc = 0 Then
                                    errStr += "Tank Material cannot be empty" + vbCrLf
                                    validateSuccess = False
                                ElseIf oTankInfo.TankMatDesc <> 350 And oTankInfo.TankModDesc = 0 Then
                                    errStr += "Tank Secondary Information cannot be empty" + vbCrLf
                                    validateSuccess = False
                                ElseIf oTankInfo.TankModDesc <> 0 And oTankInfo.TankLD = 0 Then
                                    errStr += "Tank Release Detection cannot be empty" + vbCrLf
                                    validateSuccess = False
                                Else
                                    validateSuccess = True
                                End If

                                If oTankInfo.TankStatus = 429 Then 'TOSI
                                    If Date.Compare(oTankInfo.DateLastUsed, CDate("01/01/0001")) = 0 Then
                                        errStr += "Date Last Used cannot be empty" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If

                                If Date.Compare(oTankInfo.DateInstalledTank, CDate("01/01/0001")) = 0 Then
                                    errStr += "Date Installed cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If

                                If Date.Compare(oTankInfo.PlacedInServiceDate, CDate("01/01/0001")) = 0 Then
                                    errStr += "Date Placed in Service cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If

                                ' Tank Placed in Service On >= Tank Installed On
                                If Date.Compare(oTankInfo.PlacedInServiceDate, dtNullDate) <> 0 And Date.Compare(oTankInfo.DateInstalledTank, dtNullDate) <> 0 Then
                                    If Date.Compare(oTankInfo.PlacedInServiceDate, oTankInfo.DateInstalledTank) < 0 Then
                                        errStr += "Tank Placed in Service On must be greater or equal to Tank Installed On" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If

                                ' TankModDesc like cathodically protected
                                If validateCapDates Then
                                    If oTankInfo.TankModDesc = 412 Or oTankInfo.TankModDesc = 415 Or oTankInfo.TankModDesc = 475 Then
                                        If Date.Compare(oTankInfo.TCPInstallDate, dtNullDate) = 0 Then
                                            msgStr += "Provide Tank CP Install Date" + vbCrLf
                                        End If

                                        dttemp = oTankInfo.LastTCPDate
                                        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                                        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                            msgStr += "Valid date for Tank CP Last Tested : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                                        End If
                                    End If
                                End If

                                ' Tank CP Last Tested On >= Tank CP Installed On
                                If Date.Compare(oTankInfo.LastTCPDate, dtNullDate) <> 0 And Date.Compare(oTankInfo.TCPInstallDate, dtNullDate) <> 0 Then
                                    If Date.Compare(oTankInfo.LastTCPDate, oTankInfo.TCPInstallDate) < 0 Then
                                        errStr += "Tank CP Last Tested must be greater or equal to Tank CP Installed" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If

                                ' TankModDesc = Lined Interior

                                If validateCapDates Then
                                    If oTankInfo.TankModDesc = 476 Then
                                        dttemp = oTankInfo.LinedInteriorInspectDate
                                        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        ' install = null 10 yrs
                                        If Date.Compare(oTankInfo.LinedInteriorInstallDate, dtNullDate) = 0 Then
                                            msgStr += "Provide Tank InteriorLinning Install Date" + vbCrLf
                                            dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                                        Else ' if install is more than 15 yrs old, 5 yrs
                                            ' first inspection = 10yrs, second and onwards = 5yrs
                                            If Date.Compare(oTankInfo.LinedInteriorInstallDate, DateAdd(DateInterval.Year, -15, Today.Date)) <= 0 Then
                                                dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                                            Else
                                                dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                                            End If
                                            'If Date.Compare(oTankInfo.LinedInteriorInspectDate, dtNullDate) = 0 Or _
                                            '    Date.Compare(oTankInfo.LinedInteriorInstallDate, oTankInfo.LinedInteriorInspectDate) > 0 Then
                                            '    dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                                            'Else
                                            '    dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                                            'End If
                                        End If
                                        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                            msgStr += "Valid date for Tank InteriorLinning Inspection : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                                        End If
                                    End If
                                End If

                                ' Tank LastInterior Lining Inspection On >= Tank Interior Lining Installed On
                                If Date.Compare(oTankInfo.LinedInteriorInspectDate, dtNullDate) <> 0 And Date.Compare(oTankInfo.LinedInteriorInstallDate, dtNullDate) <> 0 Then
                                    If Date.Compare(oTankInfo.LinedInteriorInspectDate, oTankInfo.LinedInteriorInstallDate) < 0 Then
                                        errStr += "Tank LastInterior Lining Inspection must be greater or equal to Tank Interior Lining Installed" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If

                                '****************************************************************************'
                                '*                                                                          *'
                                '****************************************************************************'
                                If oTankInfo.TankStatus = 424 Then 'CIU

                                    ' TankLD = Inventory Control/Prevision Tightness Testing
                                    If validateCapDates Then
                                        If oTankInfo.TankLD = 338 Then
                                            dttemp = oTankInfo.TTTDate
                                            dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                            dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                                            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                                msgStr += "Valid date for Tank Tightness Test : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                                            End If
                                        End If
                                    End If

                                End If

                            ElseIf oTankInfo.TankStatus = 425 Then 'TOS
                                If Date.Compare(oTankInfo.DateInstalledTank, CDate("01/01/0001")) = 0 Then
                                    errStr += "Date Installed cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If

                                If Date.Compare(oTankInfo.PlacedInServiceDate, CDate("01/01/0001")) = 0 Then
                                    errStr += "Date Placed in Service cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If

                                If Date.Compare(oTankInfo.DateLastUsed, CDate("01/01/0001")) = 0 Then
                                    errStr += "Date Last Used cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If
                            ElseIf oTankInfo.TankStatus = 426 Then 'POU
                                Dim dtTempDate As Date = "12-22-1988"
                                If Date.Compare(oTankInfo.DateLastUsed, CDate("01/01/0001")) = 0 Then
                                    errStr += "Date Last Used cannot be empty" + vbCrLf
                                    validateSuccess = False
                                Else
                                    If Date.Compare(oTankInfo.DateLastUsed, dtTempDate) >= 0 Then
                                        If Not oTankInfo.NonPre88 Then
                                            errStr += "Date Last Used must be < '12-22-1988' for POU Tanks" + vbCrLf
                                            validateSuccess = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                End Select
                If errStr.Length > 0 And Not validateSuccess Then
                    If msgStr.Length > 0 Then errStr += "Optional:" + vbCrLf + msgStr
                    If returnString Then
                        strError = errStr
                    Else
                        RaiseEvent evtTankValidationErr(oTankInfo.TankId, errStr)
                    End If
                ElseIf msgStr.Length > 0 Then
                    If returnString Then
                        strError = "Optional:" + vbCrLf + msgStr
                    Else
                        MsgBox("Optional:" + vbCrLf + msgStr, MsgBoxStyle.OKOnly, "Tank Validation")
                    End If
                End If
                Return validateSuccess
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub SetInfoInChild()
            oTankCompartment.TankInfo = oTankInfo
            oTankCompartment.Pipes.TankInfo = oTankInfo
        End Sub
        Private Sub CheckTOSTOSIRules(ByRef tankInfo As MUSTER.Info.TankInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strExcludeOCE As String)
            Dim citation As New MUSTER.BusinessLogic.pInspectionCitation
            Dim citationExists As Boolean = False

            passYesToTOSTOSIToPipes = False

            Try
                ' check for CIU <=> TOS/TOSI rules for existing tank (not new tank)
                If tankInfo.TankId > 0 Then
                    ' CIU TO TOS/TOSI
                    If tankInfo.OriginalTankStatus = 424 And _
                        (tankInfo.TankStatus = 429 Or tankInfo.TankStatus = 425) Then
                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), tankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)
                        ' if there is a citation for corrosion protection not maintained
                        If citationExists Then
                            ' notify user that selected status was changed to TOS if selected status was not TOS
                            If tankInfo.TankStatus <> 425 AndAlso MessageBox("Tank: " + tankInfo.TankIndex.ToString + "'s business rule states that the status should change to TOS. Do you wish to change to TOSI?") = MsgBoxResult.No Then
                                ' tank status is TOS
                                tankInfo.TankStatus = 425
                            End If

                        ElseIf tankInfo.TankStatus = 425 Then ' TOS
                            ProcessSubstandard(tankInfo, citationExists, "")
                        End If

                        ' TOSI TO CIU/TOS
                    ElseIf tankInfo.OriginalTankStatus = 429 And _
                        (tankInfo.TankStatus = 424 Or tankInfo.TankStatus = 425) Then
                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), tankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)
                        ' if status is tos, process substandard
                        ' if status is ciu, allow change - no need to check for any conditions
                        If tankInfo.TankStatus = 425 Then ' TOS
                            ProcessSubstandard(tankInfo, citationExists, "")
                            'If tankInfo.TankStatus = 425 Then ' TOS
                            '    ' create a citation for corrosion protection not maintained
                            '    ' create a manual fce with citation 10
                            '    CreateCPNotMaintainedCitation(tankInfo, moduleID, staffID, returnVal)
                            '    If Not returnVal = String.Empty Then
                            '        Exit Sub
                            '    End If
                            'End If
                        End If

                        ' TOS TO CIU/TOSI
                    ElseIf tankInfo.OriginalTankStatus = 425 And _
                            (tankInfo.TankStatus = 424 Or tankInfo.TankStatus = 429) Then
                        If tankInfo.TankStatus = 429 Then ' TOSI
                            citationExists = citation.CheckCitationExists(CDate("01/01/0001"), tankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)
                            ' if status is tosi, process substandard
                            ProcessSubstandard(tankInfo, citationExists, "")
                            'DeleteCPNotMaintainedCitation(tankInfo, moduleID, staffID, returnVal)
                            'If Not returnVal = String.Empty Then
                            '    Exit Sub
                            'End If
                        ElseIf tankInfo.TankStatus = 424 Then
                            citationExists = citation.CheckCitationExists(CDate("01/01/0001"), tankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)
                            ' if status is ciu
                            ' if there is a citation for corrosion protection not maintained
                            ' tank status is tos
                            If citationExists AndAlso MessageBox("Tank: " + tankInfo.TankIndex.ToString + "'s business rule states that the status should change to TOS. Do you wish to change to TOSI?") = MsgBoxResult.No Then

                                ' notify user that selected status was changed to TOS
                                tankInfo.TankStatus = 425
                                'DeleteCPNotMaintainedCitation(tankInfo, moduleID, staffID, returnVal)
                                'If Not returnVal = String.Empty Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    ElseIf tankInfo.OriginalTankStatus = tankInfo.TankStatus And _
                            (tankInfo.TankStatus = 425 Or tankInfo.TankStatus = 429) Then ' TOS / TOSI

                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), tankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)

                        If (tankInfo.TankStatus = 429 And citationExists) AndAlso MessageBox("Tank: " + tankInfo.TankIndex.ToString + "'s business rule states that the status should change to TOS. Do you wish to change to TOSI?") = MsgBoxResult.No Then
                            ' change to tos
                            tankInfo.TankStatus = 425
                            'CreateCPNotMaintainedCitation(tankInfo, moduleID, staffID, returnVal)
                            'If Not returnVal = String.Empty Then
                            '    Exit Sub
                            'End If
                        ElseIf tankInfo.TankStatus = 425 And Not citationExists Then
                            ProcessSubstandard(tankInfo, citationExists, "")
                            'ElseIf tankInfo.TankStatus = 425 And citationExists Then
                            '    DeleteCPNotMaintainedCitation(tankInfo, moduleID, staffID, returnVal)
                            '    If Not returnVal = String.Empty Then
                            '        Exit Sub
                            '    End If
                        End If

                    End If
                Else
                    ' if it is a new tank
                    ' if tank status is TOS, process substandard
                    If tankInfo.TankStatus = 425 Then
                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), tankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)
                        ProcessSubstandard(tankInfo, citationExists, "")
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub ProcessSubstandard(ByRef tankInfo As MUSTER.Info.TankInfo, ByVal citationExists As Boolean, Optional ByVal msg2 As String = "")

            Dim p As BusinessLogic.pProperty

            Try
                p = New BusinessLogic.pProperty
                ' if ((tank material is "asphalt coated or bare steel" or "unknown" or "other") and
                '       tank secondary option is "none" or "not specified" (null or empty))
                '                                    OR
                '    there is a citation for corrosion protection not maintained
                '                                    OR
                '   (tank materisl is "epoxy" and tank secondary option is "none" and tank status is "tos")
                '       tank status = tos
                ' else
                '       tank status = tosi

                Dim msg As String = String.Empty

                If ((tankInfo.TankMatDesc = 344 Or tankInfo.TankMatDesc = 350 Or tankInfo.TankMatDesc = 351) And _
                    (tankInfo.TankModDesc = 414 Or tankInfo.TankModDesc = 0)) Or _
                    citationExists Or _
                    (tankInfo.TankMatDesc = 347 And tankInfo.TankModDesc = 414 And tankInfo.TankStatus = 425) Then


                    ' notify user that selected status was changed to TOS if selected status was not TOS
                    If tankInfo.TankStatus <> 425 AndAlso MessageBox("Tank: " + tankInfo.TankIndex.ToString + String.Format("'s business rule states that the status should change to TOS{0}. Do you wish to change to TOSI?", msg)) = MsgBoxResult.No Then
                        tankInfo.TankStatus = 425
                    End If
                    ' tank status is TOS
                Else

                    ' notify user that selected status was changed to TOSI if selected status was not TOSI
                    If tankInfo.TankStatus <> 429 AndAlso MessageBox("Tank: " + tankInfo.TankIndex.ToString + String.Format("'s business rule states that the status should change to TOSI{0}. Do you wish to change to TOS?", msg)) = MsgBoxResult.No Then

                        ' tank status is TOSI
                        tankInfo.TankStatus = 429

                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                p = Nothing
            End Try
        End Sub
        'Private Sub CreateCPNotMaintainedCitation(ByVal tankInfo As MUSTER.Info.TankInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
        '    Dim oInspection As New MUSTER.BusinessLogic.pInspection
        '    Dim oFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
        '    Dim oInspectionCitation As New MUSTER.BusinessLogic.pInspectionCitation
        '    Dim nOwnerID As Integer = oFacilityInfo.OwnerID
        '    Try
        '        ' check rights
        '        CheckRightsToSaveCPNotMaintainedCitation(moduleID, staffID, returnVal)
        '        If Not returnVal = String.Empty Then
        '            Exit Sub
        '        End If
        '        ' create inspection
        '        oInspection.Retrieve(0)
        '        oInspection.FacilityID = tankInfo.FacilityId
        '        If tankInfo.FacilityId <> oFacilityInfo.ID Then
        '            Dim ds As DataSet = oTankDB.DBGetDS("SELECT OWNER_ID FROM TBLREG_FACILITY WHERE FACILITY_ID = " + tankInfo.FacilityId.ToString)
        '            If ds.Tables(0).Rows(0)("OWNER_ID") Is DBNull.Value Then
        '                returnVal = "Invalid Owner ID for facility " + tankInfo.FacilityId.ToString
        '                Exit Sub
        '            Else
        '                nOwnerID = ds.Tables(0).Rows(0)("OWNER_ID")
        '            End If
        '        End If
        '        oInspection.OwnerID = nOwnerID
        '        oInspection.InspectionType = 1132
        '        oInspection.LetterGenerated = False
        '        oInspection.CreatedBy = IIf(tankInfo.ModifiedBy = String.Empty, tankInfo.CreatedBy, tankInfo.ModifiedBy)
        '        oInspection.Save(moduleID, staffID, returnVal, , , , True)
        '        ' create manual fce
        '        oFCE.Retrieve(0)
        '        oFCE.InspectionID = oInspection.ID
        '        oFCE.OwnerID = oInspection.OwnerID
        '        oFCE.FacilityID = oInspection.FacilityID
        '        oFCE.Source = "ADMIN"
        '        oFCE.FCEDate = Today.Date
        '        oFCE.CreatedBy = oInspection.CreatedBy
        '        oFCE.Save(moduleID, staffID, returnVal, , , True)
        '        ' create citation
        '        oInspectionCitation.Retrieve(oInspection.InspectionInfo, 0)
        '        oInspectionCitation.FacilityID = oFCE.FacilityID
        '        oInspectionCitation.FCEID = oFCE.ID
        '        oInspectionCitation.InspectionID = oInspection.ID
        '        oInspectionCitation.CitationID = 10
        '        oInspectionCitation.QuestionID = oInspection.CheckListMaster.RetrieveByCheckListItemNum("99998").ID
        '        oInspectionCitation.CreatedBy = oFCE.CreatedBy
        '        oInspectionCitation.Save(moduleID, staffID, returnVal, , , True)
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub DeleteCPNotMaintainedCitation(ByVal tankInfo As MUSTER.Info.TankInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
        '    Dim nFCEID, nInspectionID, nCitationID As Integer
        '    Try
        '        ' check rights
        '        CheckRightsToSaveCPNotMaintainedCitation(moduleID, staffID, returnVal)
        '        If Not returnVal = String.Empty Then
        '            Exit Sub
        '        End If
        '        ' delete only if the citation is linked to manual fce
        '        Dim strSQL As String = "SELECT FCE.FCE_ID, FCE.INSPECTION_ID, C.INS_CIT_ID " + _
        '                                "FROM tblCAE_FACILITIY_COMPLIANCE_EVENT FCE " + _
        '                                "INNER JOIN tblINS_INSPECTION_CITATION C ON C.INSPECTION_ID = FCE.INSPECTION_ID AND C.FCE_ID = FCE.FCE_ID " + _
        '                                "WHERE FCE.FACILITY_ID = " + tankInfo.FacilityId.ToString + _
        '                                "AND FCE.SOURCE = 'ADMIN' " + _
        '                                "AND FCE.DELETED = 0 " + _
        '                                "AND FCE.OCE_GENERATED = 0 " + _
        '                                "AND C.RESCINDED = 0 " + _
        '                                "AND C.NFA_DATE IS NULL"
        '        Dim ds As DataSet = oTankDB.DBGetDS(strSQL)
        '        If ds.Tables(0).Rows.Count > 0 Then
        '            nFCEID = IIf(ds.Tables(0).Rows(0)("FCE_ID") Is DBNull.Value, 0, ds.Tables(0).Rows(0)("FCE_ID"))
        '            nInspectionID = IIf(ds.Tables(0).Rows(0)("INSPECTION_ID") Is DBNull.Value, 0, ds.Tables(0).Rows(0)("INSPECTION_ID"))
        '            nCitationID = IIf(ds.Tables(0).Rows(0)("INS_CIT_ID") Is DBNull.Value, 0, ds.Tables(0).Rows(0)("INS_CIT_ID"))
        '            If nFCEID <> 0 And nInspectionID <> 0 And nCitationID <> 0 Then
        '                Dim oInspection As New MUSTER.BusinessLogic.pInspection
        '                Dim oFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
        '                Dim oInspectionCitation As New MUSTER.BusinessLogic.pInspectionCitation
        '                oInspection.Retrieve(nInspectionID)
        '                oFCE.Retrieve(nFCEID)
        '                oInspectionCitation.Retrieve(oInspection.InspectionInfo, nCitationID)
        '                oInspection.Deleted = True
        '                oFCE.Deleted = True
        '                oInspectionCitation.Deleted = True
        '                oInspection.Save(moduleID, staffID, returnVal, True, , , True)
        '                If Not returnVal = String.Empty Then
        '                    Exit Sub
        '                End If
        '                oFCE.Save(moduleID, staffID, returnVal, True, , True)
        '                If Not returnVal = String.Empty Then
        '                    Exit Sub
        '                End If
        '                oInspectionCitation.Save(moduleID, staffID, returnVal, True, , True)
        '                If Not returnVal = String.Empty Then
        '                    Exit Sub
        '                End If
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckRightsToSaveCPNotMaintainedCitation(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
        '    Try
        '        If Not oTankDB.SqlHelperProperty.HasWriteAccess(moduleID, staffID, oTankDB.SqlHelperProperty.EntityTypes.Inspection) Then
        '            returnVal = "Inspection, "
        '        End If
        '        If Not oTankDB.SqlHelperProperty.HasWriteAccess(moduleID, staffID, oTankDB.SqlHelperProperty.EntityTypes.CAEFacilityCompliantEvent) Then
        '            returnVal += "FCE, "
        '        End If
        '        If Not oTankDB.SqlHelperProperty.HasWriteAccess(moduleID, staffID, oTankDB.SqlHelperProperty.EntityTypes.Citation) Then
        '            returnVal += "Citation, "
        '        End If
        '        If returnVal <> String.Empty Then
        '            returnVal = "You do not have rights to save " + returnVal.Trim.TrimEnd(",")
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        Public Function GetAttachedPipeIDs(ByVal tankID As Integer) As String
            Try
                Return oTankDB.DBGetAttachedPipeIDs(tankID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetAttachedPipes(ByVal pipeIDs As String) As DataSet
            Try
                Dim strSQL As String
                strSQL = "SELECT * FROM v_DETACH_PIPES_DISPLAY_DATA WHERE PIPE_ID IN (" + pipeIDs + ") " + _
                            "ORDER BY [TANK SITE ID], [COMPARTMENT NUMBER], [PIPE SITE ID]"
                Return oTankDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function GetParentablePipes(ByVal tankID As Integer) As DataSet

            Try
                Dim strSQL As String
                strSQL = "SELECT [PIPE ID], [PIPE SITE ID],[COMPARTMENT NUMBER],[MATERIAL],[TYPE],[DESC]  FROM v_PIPES_DISPLAY_DATA WHERE [TANK ID] = '" + tankID.ToString + "' " + _
                            " AND [HAS PARENT] = 'No' Order by [PIPE SITE ID] "
                Return oTankDB.DBGetDS(strSQL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function DetachPipe(ByVal pipeID As Integer, ByVal tankID As Integer, ByVal compNum As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String) As Boolean
            Try
                oTankDB.DetachPipe(pipeID, tankID, compNum, moduleID, staffID, returnVal, strUser)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return False
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAllInfo() As MUSTER.Info.TankCollection
            Try
                oFacilityInfo.TankCollection.Clear()
                oFacilityInfo.TankCollection = oTankDB.GetAllInfo()
                Return oFacilityInfo.TankCollection
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetAllByFacilityID(ByVal nFacilityId As Integer) As MUSTER.Info.TankCollection
            Try
                oFacilityInfo.TankCollection.Clear()
                oFacilityInfo.TankCollection = oTankDB.DBGetByFacilityID(nFacilityId)
                Return oFacilityInfo.TankCollection
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Add(ByVal ID As Int64, Optional ByVal ShowDeleted As Boolean = False)
            Try
                Dim dtTempDate As Date = "12-22-1988"
                oTankInfo = oTankDB.DBGetByID(ID, ShowDeleted)
                If oTankInfo.TankId = 0 Then
                    oTankInfo.TankId = nID
                    oTankInfo.FacilityId = oFacilityInfo.ID
                    nID -= 1
                End If
                If oTankInfo.TankStatus = 426 Then
                    oTankInfo.POU = True
                    If Date.Compare(oTankInfo.DateLastUsed, CDate("12/22/1988")) >= 0 Then
                        oTankInfo.NonPre88 = True
                    End If
                End If
                If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                    oTankInfo.FacilityPowerOff = True
                Else
                    oTankInfo.FacilityPowerOff = False
                End If
                oTankInfo.FacCapStatus = oFacilityInfo.CapStatus
                oFacilityInfo.TankCollection.Add(oTankInfo)
                SetInfoInChild()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Add(ByRef oTank As MUSTER.Info.TankInfo)
            Try
                Dim dtTempDate As Date = "12-22-1988"
                oTankInfo = oTank
                If oTankInfo.TankId = 0 Then
                    oTankInfo.TankId = nID
                    oTankInfo.FacilityId = oFacilityInfo.ID
                    nID -= 1
                End If
                If oTankInfo.TankStatus = 426 Then
                    oTankInfo.POU = True
                    If Date.Compare(oTankInfo.DateLastUsed, CDate("12/22/1988")) >= 0 Then
                        oTankInfo.NonPre88 = True
                    End If
                End If
                If Date.Compare(oFacilityInfo.DatePowerOff, CDate("01/01/0001")) <> 0 Then
                    oTankInfo.FacilityPowerOff = True
                Else
                    oTankInfo.FacilityPowerOff = False
                End If
                oTankInfo.FacCapStatus = oFacilityInfo.CapStatus
                oFacilityInfo.TankCollection.Add(oTankInfo)
                SetInfoInChild()
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Remove(ByVal ID As Int64)
            Try
                Dim oTankInfoLocal As MUSTER.Info.TankInfo
                oTankInfoLocal = oFacilityInfo.TankCollection.Item(ID)
                If Not (oTankInfoLocal Is Nothing) Then
                    oFacilityInfo.TankCollection.Remove(oTankInfoLocal)
                    ' RaiseEvent evtTankInfoFac(oTankInfo, "REMOVE")
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Tank " & ID.ToString & " is not in the collection of Tank.")
        End Sub
        Public Sub Remove(ByVal oTankInfoLocal As MUSTER.Info.TankInfo)
            Try
                If oFacilityInfo.TankCollection.Contains(oTankInfoLocal.TankId) Then
                    oFacilityInfo.TankCollection.Remove(oTankInfoLocal)
                End If
                ' RaiseEvent evtTankInfoFac(oTankInfoLocal, "REMOVE")
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Tank " & oTankInfoLocal.TankId & " is not in the collection of Tank.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolSaveAsInspection As Boolean = False)
            Dim IDs As New Collection
            Dim delIDs As New Collection
            Dim index As Integer
            Dim oTempInfo As MUSTER.Info.TankInfo
            Try
                'Dim colTankContained As MUSTER.Info.TankCollection
                'RaiseEvent evtFacilityInfoTankCol(colTankContained)
                For Each oTempInfo In oFacilityInfo.TankCollection.Values
                    If oTempInfo.IsDirty Then
                        oTankInfo = oTempInfo
                        If oTankInfo.Deleted Then
                            If oTankInfo.TankId < 0 Then
                                delIDs.Add(oTankInfo.TankId)
                            Else
                                Me.Save(moduleID, staffID, returnVal, strUser, True, , False, , , bolSaveAsInspection)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Me.ValidateData(False) Then
                                If oTankInfo.TankId < 0 Then
                                    IDs.Add(oTankInfo.TankId)
                                End If
                                Me.Save(moduleID, staffID, returnVal, strUser, True, , False, , , bolSaveAsInspection)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            Else : Exit For
                            End If
                        End If
                        'If oTankInfo.TankId < 0 Then
                        '    If oTankInfo.Deleted Then
                        '        delIDs.Add(oTankInfo.TankId)
                        '    Else
                        '        If Me.ValidateData(False) Then
                        '            IDs.Add(oTankInfo.TankId)
                        '            Me.Save(moduleID, staffID, returnVal, strUser, True)
                        '        Else : Exit For
                        '        End If
                        '    End If
                        'Else
                        '    If oTankInfo.Deleted Then
                        '        delIDs.Add(oTankInfo.TankId)
                        '    End If
                        '    Me.Save(moduleID, staffID, returnVal, strUser, True)
                        'End If
                    ElseIf oTempInfo.TankId > 0 And oTempInfo.ChildrenDirty Then
                        oTankInfo = oTempInfo
                        SetInfoInChild()
                        oTankCompartment.Flush(moduleID, staffID, returnVal, strUser, bolSaveAsInspection)
                        If Not returnVal = String.Empty Then
                            Exit Sub
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        oTempInfo = oFacilityInfo.TankCollection.Item(CType(delIDs.Item(index), String))
                        oFacilityInfo.TankCollection.Remove(oTempInfo)
                        'RaiseEvent evtTankInfoFac(oTempInfo, "REMOVE")
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        oTempInfo = oFacilityInfo.TankCollection.Item(colKey)
                        oFacilityInfo.TankCollection.ChangeKey(colKey, oTempInfo.TankId.ToString)
                        'oFacilityInfo.TankCollection.ChangeKey(colKey, oTempInfo.TankId.ToString)
                    Next
                End If
                'oTankCompartment.Flush()
                'oComments.Flush()
                RaiseEvent evtTanksChanged(oTankInfo.IsDirty)
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
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            Dim colstrArr As New Collection
            Dim i As Integer
            'Dim colTankContained As MUSTER.Info.TankCollection
            'RaiseEvent evtFacilityInfoTankCol(colTankContained)
            For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                colstrArr.Add(oTankInfoLocal.TankId)
            Next
            Dim strArr(colstrArr.Count - 1) As String
            For i = 0 To colstrArr.Count - 1
                strArr(i) = CType(colstrArr(i + 1), String)
            Next
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.TankId.ToString))
            If colIndex + direction > -1 Then
                If colIndex + direction <= nArr.GetUpperBound(0) Then
                    Return oFacilityInfo.TankCollection.Item(nArr.GetValue(colIndex + direction)).TankId.ToString
                Else
                    Return oFacilityInfo.TankCollection.Item(nArr.GetValue(0)).TankId.ToString
                End If
            Else
                Return oFacilityInfo.TankCollection.Item(nArr.GetValue(nArr.GetUpperBound(0))).TankId.ToString
            End If
        End Function
        Public Function GetTanksStatus(ByRef FacInfo As MUSTER.Info.FacilityInfo) As Integer
            oFacilityInfo = FacInfo
            ' To be stored globally as enum
            ' 510	Confirmed Release
            ' 511   Unconfirmed(Release)
            ' 512   Closed(Release)
            ' 513   Unregulated()
            ' 514   Active()
            ' 515   Closed()
            ' 516   Pre(88)

            ' 424   Currently In User (CIU)
            ' 425   Temporarily Out of Service (TOS)
            ' 426   Permanently Out of Use (POU)
            ' 429   Temporarily Out of Service Indefinitely (TOSI)
            Dim status As Integer = 515 'Closed
            Dim dtLastUsed As Date = "12/22/1988"
            Dim i As Integer = 0
            Try
                'Dim colTankContained As New MUSTER.Info.TankCollection
                'RaiseEvent evtFacilityInfoTankColByFacilityID(oFacilityInfo.ID, colTankContained)
                If oFacilityInfo.TankCollection.Count > 0 Then
                    Dim oTankInfoLocal As MUSTER.Info.TankInfo
                    For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                    Next
                    status = GetALLTanksStatusFromCollection(513)
                    If status = 0 Then
                        status = GetOneTanksStatusFromCollection(424, 425, 429)
                        If status = 0 Then
                            status = GetALLTanksStatusFromCollection(426)
                            If status = 0 Then
                                status = 515
                            Else
                                status = GetAllLastUsedDateFromCollection(dtLastUsed)
                                If status = 0 Then
                                    status = 515
                                End If
                            End If
                        Else : Exit Try
                        End If
                    Else : Exit Try
                    End If
                Else : Exit Try 'status = 515
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Return status
        End Function
        Private Function GetALLTanksStatusFromCollection(ByVal Status As Integer) As Integer
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            Dim nCount As Integer = 0
            Dim bolUnregulated As Boolean = False
            For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                If oTankInfoLocal.TankStatus <> Status Then
                    Return 0
                Else
                    nCount += 1
                End If
            Next
            If nCount = oFacilityInfo.TankCollection.Count Then
                Return Status
            Else
                Return 0
            End If
        End Function
        Private Function GetOneTanksStatusFromCollection(ByVal Status1 As Integer, ByVal Status2 As Integer, ByVal Status3 As Integer) As Integer
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                If oTankInfoLocal.TankStatus = Status1 Or _
                oTankInfoLocal.TankStatus = Status2 Or _
                oTankInfoLocal.TankStatus = Status3 Then
                    Return 514
                End If
            Next
            Return 0
        End Function
        Private Function GetAllLastUsedDateFromCollection(ByVal dtLastUsed As Date) As Integer
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            'Dim colTankContained As MUSTER.Info.TankCollection
            'RaiseEvent evtFacilityInfoTankColByFacilityID(facID, colTankContained)
            For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                If Date.Compare(oTankInfoLocal.DateLastUsed, dtLastUsed) >= 0 Then
                    Return 0
                End If
            Next
            Return 516
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oTankInfo = New MUSTER.Info.TankInfo
            oFacilityInfo.TankCollection.Clear()
            oTankCompartment.Clear()
        End Sub
        Public Sub Reset()
            Dim oldTnkStat As Integer = oTankInfo.TankStatus
            oTankInfo.Reset()
            'If oTankInfo.DateLastUsed <> CDate("01/01/0001") Then
            '    CheckDateLastUsed(oTankInfo.DateLastUsed)
            'End If
            'If oTankInfo.TankLD <> 0 Then
            '    CheckTankLD(oTankInfo.TankLD)
            'End If
            'If oTankInfo.TankMatDesc <> 0 Then
            '    CheckTankMatDesc(oTankInfo.TankMatDesc)
            'End If
            'If oTankInfo.LinedInteriorInspectDate <> CDate("01/01/0001") Then
            'CheckLinedInteriorInspectDate(oTankInfo.LinedInteriorInspectDate)
            'End If
            'If oTankInfo.TankModDesc <> 0 Then
            '    CheckTankModDesc(oTankInfo.TankModDesc)
            'End If
            'If oTankInfo.TankStatus <> 0 Then
            '    CheckTankStatus(oldTnkStat, oTankInfo.TankStatus)
            'End If
            'If oTankInfo.LastTCPDate <> CDate("01/01/0001") Then
            'CheckLastTCPDate(oTankInfo.LastTCPDate)
            'End If
            'If oTankInfo.TTTDate <> CDate("01/01/0001") Then
            '    CheckTTTDate(oTankInfo.TTTDate)
            'End If
            oTankCompartment.Reset()
            'oComments.Reset()
        End Sub
        Public Sub ResetCollection()
            Dim xTankLocalInfo As MUSTER.Info.TankInfo
            If Not oFacilityInfo.TankCollection.Values Is Nothing Then
                For Each xTankLocalInfo In oFacilityInfo.TankCollection.Values
                    If xTankLocalInfo.IsDirty Then
                        xTankLocalInfo.Reset()
                    End If
                Next
            End If

            'Need to check with JAY/ADAM
            'oComments.Reset()

        End Sub
#End Region
#Region "Look Up Operations"
        Public Function PopulateTankStatus(Optional ByVal nMode As String = "", Optional ByVal DateLastUsed As Object = Nothing) As DataTable
            Try
                If UCase(nMode).Trim = "ADD" Then
                    Dim dtReturn As DataTable = GetDataTable("vTANKSTATUSTYPE WHERE PROPERTY_ID = 424 OR PROPERTY_ID = 425 OR PROPERTY_ID = 429 OR PROPERTY_ID = 426")
                    Return dtReturn
                ElseIf UCase(nMode).Trim = "EDIT" Then
                    Dim dtTempDate As Date = "12-22-1988"
                    If oTankInfo.POU Then
                        Dim dtreturn As DataTable = GetDataTable("vTANKSTATUSTYPE WHERE PROPERTY_ID = 426")
                        Return dtreturn
                    Else
                        Dim str As String = ""
                        If Date.Compare(DateLastUsed, CDate("01/01/0001")) <> 0 Then
                            If Date.Compare(DateLastUsed, dtTempDate) >= 0 Then
                                str = " WHERE PROPERTY_ID <> 426"
                            End If
                        End If
                        If oTankInfo.TankStatus = 426 Then
                            Dim dtreturn As DataTable = GetDataTable("vTANKSTATUSTYPE WHERE PROPERTY_ID = 426")
                            Return dtReturn
                        Else
                            Dim dtreturn As DataTable = GetDataTable("vTANKSTATUSTYPE" + str)
                            Return dtreturn
                        End If
                    End If
                Else
                    Dim dtReturn As DataTable = GetDataTable("vTANKSTATUSTYPE")
                    Return dtReturn
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateTankStatusShort(Optional ByVal nMode As String = "", Optional ByVal DateLastUsed As Object = Nothing) As DataTable
            Try
                If UCase(nMode).Trim = "ADD" Then
                    Dim dtReturn As DataTable = GetDataTable("vPIPETATUSTYPESHORT WHERE PROPERTY_ID = 424 OR PROPERTY_ID = 425 OR PROPERTY_ID = 429 OR PROPERTY_ID = 426")
                    Return dtReturn
                ElseIf UCase(nMode).Trim = "EDIT" Then
                    Dim dtTempDate As Date = "12-22-1988"
                    If oTankInfo.POU Then
                        Dim dtreturn As DataTable = GetDataTable("vPIPESTATUSTYPESHORT WHERE PROPERTY_ID = 426")
                        Return dtreturn
                    Else
                        Dim str As String = ""
                        If Date.Compare(DateLastUsed, CDate("01/01/0001")) <> 0 Then
                            If Date.Compare(DateLastUsed, dtTempDate) >= 0 Then
                                str = " WHERE PROPERTY_ID <> 426"
                            End If
                        End If
                        If oTankInfo.TankStatus = 426 Then
                            Dim dtreturn As DataTable = GetDataTable("vPIPESTATUSTYPESHORT WHERE PROPERTY_ID = 426")
                            Return dtReturn
                        Else
                            Dim dtreturn As DataTable = GetDataTable("vPIPESTATUSTYPESHORT" + str)
                            Return dtreturn
                        End If
                    End If
                Else
                    Dim dtReturn As DataTable = GetDataTable("vPIPESTATUSTYPESHORT")
                    Return dtReturn
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function PopulateTankSecondaryOption(Optional ByVal nVal As Int64 = 0) As DataTable
            Dim dtReturn As DataTable
            Try
                If nVal = 0 Then
                    dtReturn = GetDataTable("vSECONDARYTANKOPTIONSTYPE")
                    If oTankInfo.TankStatus = 424 Or oTankInfo.TankStatus = 429 Then
                        Return dtReturn
                    Else
                        dtReturn = Me.GetDistinctDataTableListItems(dtReturn)
                        'RaiseEvent eTankSecondaryOption(dtReturn)
                        Return dtReturn
                    End If
                Else
                    If oTankInfo.TankStatus = 424 Or oTankInfo.TankStatus = 429 Then
                        dtReturn = GetDataTable("vSECONDARYTANKOPTIONSTYPE", nVal)
                        'RaiseEvent eTankSecondaryOption(dtReturn)
                        Return dtReturn
                    Else
                        Return PopulateTankSecondaryOption()
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Different one because raises a Event
        Public Function PopulateTankReleaseDetection(Optional ByVal nVal As Int64 = 0, Optional ByVal nCap As Int64 = 0, Optional ByVal AfterInstallDate As Boolean = False) As DataTable
            Try
                Dim dtReturn As DataTable
                Dim str As String
                Dim FinalStr As String
                If nVal = 0 Then
                    dtReturn = Nothing
                    Return dtReturn
                Else
                    'If nVal = 412 Or nVal = 415 Then          'Cathodically Protected
                    '    RaiseEvent eDtPickEnableDisable(True, False)
                    'ElseIf nVal = 475 Then      'Cathodically Protected/Lined Interior
                    '    RaiseEvent eDtPickEnableDisable(True, True)
                    'ElseIf nVal = 476 Then      'Lined Interior
                    '    RaiseEvent eDtPickEnableDisable(False, True)
                    'Else                        'Rest of the Values
                    '    RaiseEvent eDtPickEnableDisable(False, False)
                    'End If

                    'Filter According to Page 58 of Documentation
                    '-------------------------------------------------------------------------
                    If nCap = 0 Then   '-- just check for changes in Tank Emergency Off Switch
                        If oTankInfo.TankEmergen Then
                            str += " AND PROPERTY_ID <> 337" ' Manual Tank Gauging
                        End If
                    Else
                        If nCap > 2000 Or oTankInfo.TankEmergen Then
                            str += " AND PROPERTY_ID <> 337" ' Manual Tank Gauging
                        End If
                    End If
                    '--------------------------------------------------------------------------

                    If oTankInfo.TankModDesc <> 413 And oTankInfo.TankModDesc <> 415 Then
                        str += " AND PROPERTY_ID <> 339 AND PROPERTY_ID <> 343"
                    End If

                    ' to find greatest of the dateinstalled, cpdateinstalled, linedinstalled dates
                    Dim dt As Date
                    ' if installed date is greater than cp installed date, assign installed date to dt
                    If Date.Compare(oTankInfo.DateInstalledTank, oTankInfo.TCPInstallDate) > 0 Then
                        dt = oTankInfo.DateInstalledTank
                    Else
                        dt = oTankInfo.TCPInstallDate
                    End If
                    ' if lined interior install date is greater than dt, set dt to lined interior install date
                    If Date.Compare(oTankInfo.LinedInteriorInstallDate, dt) > 0 Then
                        dt = oTankInfo.LinedInteriorInstallDate
                    End If

                    If Date.Compare(DateAdd(DateInterval.Year, 10, dt), Now.Date) < 0 Then
                        str += " AND PROPERTY_ID <> 338"
                    End If

                    If AfterInstallDate Then
                        str += " AND property_name like '%Interstitial Monitoring%' "
                    End If



                    ' deferred is unavailable by default and
                    ' only available if Tank is an emergency power generator tank
                    If Not oTankInfo.TankEmergen Then
                        str += " AND PROPERTY_ID <> 341"
                    ElseIf Not AfterInstallDate Then
                        str = " AND PROPERTY_ID = 341"
                    End If

                    If str.Length > 0 Then
                        FinalStr = "VRELEASEDETECTIONTYPE WHERE PROPERTY_ID_PARENT =" + nVal.ToString() + str
                    End If


                    dtReturn = GetDataTable(FinalStr)
                    'RaiseEvent eTankReleaseDetection(dtReturn)
                    Return dtReturn
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankManufacturer() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTANKMANUFACTURER")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankSubstance() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vSUBSTANCEDESCTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankMaterialOfConstruction() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTANKMATERIALOFCONSTRUCTIONTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankOverFillProtectionType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTANKOVERFILLPROTECTIONTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankCPType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VTANKCATHODICPROTECTIONTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTANKTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateCompartmentFuelTypes(Optional ByVal compSubstance As Int64 = 0) As DataTable
            Dim dtReturn As DataTable
            Try
                If compSubstance = 0 Then
                    dtReturn = GetDataTable("vFUELTYPES")
                    dtReturn = GetDistinctDataTableListItems(dtReturn)
                Else
                    dtReturn = GetDataTable("vFUELTYPES", compSubstance)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateProhibition() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTankProhibtion")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateCompartmentSubstance() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vSUBSTANCEDESCTYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateCERCLA() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCERLATYPE")
                dtReturn.Columns(0).ColumnName = "PROPERTY_ID"
                dtReturn.Columns(1).ColumnName = "PROPERTY_NAME"
                dtReturn.Columns(2).ColumnName = "PROPERTY_DESC"
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankPipeClosureStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTANKPIPECLOSURESTATUS")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateTankPipeInertFill() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vINNERTMATERIAL")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function populateClosureStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCLOSURESTATUS")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateClosureType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCLOSURETYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal nVal As Int64 = 0) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            strSQL = "SELECT * FROM " & strProperty
            If nVal <> 0 Then
                strSQL = strSQL + " WHERE PROPERTY_ID_PARENT = " + nVal.ToString()
            End If
            Try
                dsReturn = oTankDB.DBGetDS(strSQL)
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
        Public Function PopulateTankContractor() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable_Company("vCOM_LICENSEENAME_REG")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCompanyName(Optional ByVal LicenseeID As Integer = 0) As DataSet
            Dim dsReturn As New DataSet

            Try
                dsReturn = oTankDB.DBGetCompanyDetails(LicenseeID)
                If Not dsReturn.Tables(0).Rows.Count > 0 Then
                    dsReturn = Nothing
                End If
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetDataTable_Company(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName

                dsReturn = oTankDB.DBGetDS(strSQL)
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
#End Region
#Region "RaisingEvents for Enable/Disabling controls in the Form"
        'Private Sub CheckDateLastUsed(ByVal dtDateLastUsed As Date)
        '    Try
        '        If Me.TankStatus = 426 Then 'permanently out of use
        '            Dim dtTempDate As Date = "12-22-88"
        '            If Date.Compare(dtDateLastUsed, dtTempDate) < 0 Then
        '                RaiseEvent ecmbTankClosureType(True)
        '                RaiseEvent ecmbTankInertFill(True)
        '            Else
        '                RaiseEvent ecmbTankClosureType(False)
        '                RaiseEvent ecmbTankInertFill(False)
        '            End If
        '        Else
        '            RaiseEvent ecmbTankClosureType(False)
        '            RaiseEvent ecmbTankInertFill(False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try

        'End Sub
        'Private Sub CheckTankLD(ByVal nTankLd As Integer)
        '    Try
        '        If nTankLd = 338 Then
        '            RaiseEvent edtPickTankTightnessTest(True)
        '            RaiseEvent echkTankDrpTubeInvControl(True)
        '        Else
        '            RaiseEvent edtPickTankTightnessTest(False)
        '            RaiseEvent echkTankDrpTubeInvControl(False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckTankMatDesc(ByVal nVal As Integer)
        '    Try
        '        PopulateTankSecondaryOption(nVal)
        '        If nVal = 350 Then
        '            RaiseEvent eDtPickEnableDisable(False, False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub

        'Private Sub CheckTankModDesc(ByVal TankModDesc As Integer)
        '    Try
        '        PopulateTankReleaseDetection(TankModDesc, 0)  '---- check for changes in Tank Mod Desc 
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub

        Private Sub CheckTankStatus(ByVal oOldStatus As Integer, ByVal nNewStatus As Integer) ' Handles oTankInfo.eInfoTankStatus
            Dim dtTempDate As Date = "12-22-1988"
            Try
                '424 - CIU
                '425 - TOS
                '426 - POU
                '428 - RegPending
                '429 - TOSI
                '430 - Unregulated
                ''' - commented by manju
                If nNewStatus = 425 Or nNewStatus = 429 Then ' TOS OR TOSI
                    '''RaiseEvent edtPickPlannedInstallation(False)
                    '''If oTankInfo.POU And oTankInfo.NonPre88 Then
                    '''    RaiseEvent edtPickLastUsed(False)
                    '''Else
                    '''    RaiseEvent edtPickLastUsed(True)
                    '''End If
                    '''RaiseEvent ecmbTankCPType(False)
                    '''RaiseEvent edtPickCPInstalled(False)
                    '''RaiseEvent edtPickCPLastTested(False)
                    '''RaiseEvent edtPickInteriorLiningInstalled(False)
                    '''RaiseEvent edtPickLastInteriorLinningInspection(False)
                    '''RaiseEvent edtPickTankTightnessTest(False)
                    '''RaiseEvent echkTankDrpTubeInvControl(False)
                    '''RaiseEvent edGridCompartmentsStatus(True)

                    If oOldStatus = 424 Then ' CIU
                        '''If Me.TankModDesc = 412 Then 'Cathodically Protected
                        '''RaiseEvent ecmbTankCPType(True)
                        '''RaiseEvent edtPickCPInstalled(True)
                        '  If chkCAPCandidate.Checked Then  'To check -Elango 
                        '''RaiseEvent edtPickCPLastTested(True)
                        'End If
                        '''ElseIf Me.TankModDesc = 476 Then 'Lined Interior
                        '''RaiseEvent edtPickInteriorLiningInstalled(True)
                        'If chkCAPCandidate.Checked Then 'To check -Elango 
                        '''RaiseEvent edtPickLastInteriorLinningInspection(True)
                        ' End If
                        '''End If
                        If nNewStatus = 429 Then '"Temporarily Out of Service Indefinitely".ToLower 
                            'Asphalt Coated or Bare Steel -344 ,'unknown  - 350,'other - 351
                            ' If ((cmbTankMaterial.Text.Trim.ToLower.IndexOf("Bare Steel".ToLower) >= 0 Or cmbTankMaterial.Text.Trim.ToLower.IndexOf("unknown".ToLower) >= 0 Or cmbTankMaterial.Text.Trim.ToLower.IndexOf("other".ToLower) >= 0) _
                            'And (cmbTankOptions.Text.Trim.ToLower.IndexOf("None".ToLower) >= 0 Or cmbTankOptions.SelectedIndex = -1)) _
                            ' Or ((cmbTankCPType.Text.Trim.ToLower.IndexOf("Impressed Current".ToLower) >= 0 And dtFacilityPowerOff.Checked)) Then

                            If ((Me.TankMatDesc = 344 Or Me.TankMatDesc = 350 Or Me.TankMatDesc = 351) _
                                And (Me.TankModDesc = 414 Or Me.TankModDesc = 0)) Or _
                                (Me.TankCPType = 418 And Me.FacilityPowerOff) Then
                                'PopulateTankSecondaryOption()
                                '''RaiseEvent ecmbTankStatustext("Temporarily Out of Service")
                                oTankInfo.TankStatus = 425
                                'Throw New Exception("Selected Index was Changed to TOS")
                                RaiseEvent evtTankErr("Selected Index was Changed to TOS")
                            End If
                            '''RaiseEvent edtPickLastUsedFocus()
                            Exit Sub
                        End If
                    End If
                    '''RaiseEvent edGridCompartmentsStatus(True)
                    '''RaiseEvent edtPickLastUsedFocus()
                    '''If nNewStatus = 425 Then  ' TOS
                    '''    PopulateTankSecondaryOption()
                    '''ElseIf nNewStatus = 429 Then ' TOSI
                    '''    PopulateTankSecondaryOption(oTankInfo.TankMatDesc)
                    '''End If
                    '''RaiseEvent ecmbTankInertFill(False)
                    '''RaiseEvent ecmbTankClosureType(False)

                    '''ElseIf nNewStatus = 424 Then ' CIU
                    '''RaiseEvent edtPickPlannedInstallation(False)
                    '''RaiseEvent edtPickLastUsed(False)
                    '''RaiseEvent ecmbTankCPType(False)
                    '''RaiseEvent edtPickCPInstalled(False)
                    '''RaiseEvent edtPickCPLastTested(False)
                    '''RaiseEvent edtPickInteriorLiningInstalled(False)
                    '''RaiseEvent edtPickLastInteriorLinningInspection(False)
                    '''RaiseEvent edtPickTankTightnessTest(False)
                    '''RaiseEvent echkTankDrpTubeInvControl(False)
                    '''RaiseEvent ecmbTankClosureType(False)
                    '''RaiseEvent ecmbTankInertFill(False)
                    '''RaiseEvent edGridCompartmentsStatus(True)

                    '''If oOldStatus = 429 Or oOldStatus = 425 Then
                    '''RaiseEvent edtPickDatePlacedInServiceFocus()
                    '''If Me.TankModDesc = 412 Then ' cathodically protected
                    '''    RaiseEvent ecmbTankCPType(True)
                    '''    RaiseEvent edtPickCPInstalled(True)
                    '''    'If chkCAPCandidate.Checked Then  'To Check - Elango 
                    '''    RaiseEvent edtPickCPLastTested(True)
                    '''    'End If
                    '''ElseIf Me.TankModDesc = 476 Then 'Lined Interior
                    '''    RaiseEvent edtPickInteriorLiningInstalled(True)
                    '''    ' If chkCAPCandidate.Checked Then
                    '''    RaiseEvent edtPickLastInteriorLinningInspection(True)
                    '''    'End If
                    '''End If
                    '''PopulateTankSecondaryOption(oTankInfo.TankMatDesc)
                    '''RaiseEvent ecmbTankInertFill(False)
                    '''RaiseEvent ecmbTankClosureType(False)
                    '''Else
                    '''RaiseEvent ecmbTankInertFill(False)
                    '''RaiseEvent ecmbTankClosureType(False)
                    '''If nNewStatus = 426 Then
                    '''    RaiseEvent edtPickLastUsedFocus()
                    '''End If
                    '''PopulateTankSecondaryOption()
                    '''End If
                    'ElseIf nNewStatus = 428 Then 'Registration Pending
                    '    RaiseEvent edtPickPlannedInstallation(True)
                    '    RaiseEvent edGridCompartmentsStatus(False)
                    '''ElseIf nNewStatus = 426 Then ' POU
                    'If oOldStatus = 424 Then
                    ' change all pipes status to pou
                    'Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
                    'For Each oPipeInfoLocal In oTankCompartment.PipeCollection.Values
                    '    If oPipeInfoLocal.TankID = oTankInfo.TankId Then
                    '        oPipeInfoLocal.PipeStatusDesc = 426 ' pou
                    '    End If
                    'Next
                    'End If
                    '''If oTankInfo.POU And oTankInfo.NonPre88 Then
                    '''    RaiseEvent edtPickLastUsed(False)
                    '''    RaiseEvent ecmbTankInertFill(True)
                    '''    RaiseEvent ecmbTankClosureType(True)
                    '''Else
                    '''    RaiseEvent edtPickLastUsed(True)
                    '''    If Date.Compare(DateLastUsed, dtTempDate) < 0 Then  ' Status POU
                    '''        RaiseEvent ecmbTankInertFill(True)
                    '''        RaiseEvent ecmbTankClosureType(True)
                    '''    Else
                    '''        RaiseEvent ecmbTankInertFill(False)
                    '''        RaiseEvent ecmbTankClosureType(False)
                    '''    End If
                    '''End If
                    '''RaiseEvent edGridCompartmentsStatus(False)
                    '''RaiseEvent echkTankDrpTubeInvControl(False)

                    '''RaiseEvent edtPickLastUsedFocus()
                    'Else ' Unregulated
                    '    RaiseEvent edGridCompartmentsStatus(False)
                    '    RaiseEvent echkTankDrpTubeInvControl(False)
                    '    RaiseEvent ecmbTankInertFill(False)
                    '    RaiseEvent ecmbTankClosureStatus(False)
                    '    If oTankInfo.POU And oTankInfo.NonPre88 Then
                    '        RaiseEvent edtPickLastUsed(False)
                    '    Else
                    '        RaiseEvent edtPickLastUsed(True)
                    '    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        'Private Sub CheckLastTCPDate(ByVal dtPickCPLastTested As Date)
        '    Try
        '        Dim dtTemp As Date
        '        If (Date.Compare(dtPickCPLastTested, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, -3, Today()), dtPickCPLastTested) < 0) Then
        '            RaiseEvent edtPickCPLastTestedMessage("Tank CP Last Tested Date cannot be more than 3 years old")
        '            'If Date.Compare(olddtPickCPLastTested, CDate("01/01/0001")) = 0 Then
        '            '    oTankInfo.LastTCPDate = System.DateTime.Now
        '            'Else
        '            '    oTankInfo.LastTCPDate = olddtPickCPLastTested
        '            'End If
        '            Exit Sub
        '            'RaiseEvent edtPickCPLastTested(False)
        '            'ElseIf DateDiff(DateInterval.Day, Today(), dtPickCPLastTested) > 0 Then
        '            '    RaiseEvent edtPickCPLastTestedMessage("Tank CP Last Tested Date cannot be greater than Today")
        '            '    'RaiseEvent edtPickCPLastTested(False)
        '            '    Exit Sub
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckTTTDate(ByVal dtPickTankTightnessTest As Date)
        '    Try
        '        ' UIUtilsGen.ToggleDateFormat(dtPickTankTightnessTest)  'TO check - Elango
        '        If Me.TankModDesc = 412 Then
        '            If (Date.Compare(TCPInstallDate, CDate("01/01/0001")) <> 0) And Not DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 10, dtPickTankTightnessTest), dtPickTankTightnessTest) < 0 Then
        '                If (Date.Compare(TCPInstallDate, CDate("01/01/0001")) <> 0) And Not DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 10, dtPickTankTightnessTest), dtPickTankTightnessTest) < 0 Then
        '                    RaiseEvent edtPickTankTightnessTestMessage("Selected Tank Tightness Test Date should be less than or equal to")
        '                    'If Date.Compare(oTankInfo.TTTDate, CDate("01/01/0001")) = 0 Then
        '                    '    oTankInfo.TTTDate = System.DateTime.Now
        '                    'Else
        '                    '    oTankInfo.TTTDate = olddtPickTankTightnessTest
        '                    'End If
        '                    Exit Sub
        '                End If
        '            End If
        '        Else
        '            If (Date.Compare(TCPInstallDate, CDate("01/01/0001")) <> 0) And Not DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 10, TCPInstallDate), dtPickTankTightnessTest) < 0 Then
        '                RaiseEvent edtPickTankTightnessTestMessage("Selected Tank Tightness Test Date should be less than or equal to")
        '                'If Date.Compare(oTankInfo.TTTDate, CDate("01/01/0001")) = 0 Then
        '                '    oTankInfo.TTTDate = System.DateTime.Now
        '                'Else
        '                '    oTankInfo.TTTDate = olddtPickTankTightnessTest
        '                'End If
        '                Exit Sub
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            Dim dr As DataRow
            Dim tbTankTable As New DataTable
            Try
                tbTankTable.Columns.Add("TANK_ID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANK_INDEX", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("FACILITY_ID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANKSTATUS", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("DATERECEIVED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("MANIFOLD", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("COMPARTMENT", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("TANKCAPACITY", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("SUBSTANCE", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("CASNUMBER", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("SUBSTANCECOMMENTS_ID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("DATELASTUSED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("DATECLOSURERECEIVED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("DATECLOSED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("CLOSURESTATUSDESC", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("CLOSURETYPE", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("INERTMATERIAL", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANKMATDESC", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANKMODDESC", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANKOTHERMATERIAL", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("OVERFILLINSTALLED", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("SPILLINSTALLED", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("LICENSEEID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("CONTRACTORID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("DATESIGNED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("DATEINSTALLEDTANK", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("SMALLDELIVERY", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("TANKEMERGEN", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("PLANNEDINSTDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("LastTCPDate", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("LINEDINTERIORINSTALLDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("LINEDINTERIORINSPECTDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("TCPINSTALLDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("TTTDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("TANKLD", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("OVERFILLTYPE", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("REVOKEREASON", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("REVOKEDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("DatePhysicallyTagged", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("PROHIBITION", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("TIGHTFILLADAPTERS", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("DROPTUBE", Type.GetType("System.Boolean"))
                tbTankTable.Columns.Add("TANKCPTYPE", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("PLACEDINSERVICEDATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("TANKTYPES", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANKLOCATION_DESCRIPTION", Type.GetType("System.String"))
                tbTankTable.Columns.Add("TANKMANUFACTURER", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("DELETED", Type.GetType("System.Boolean"))

                'Dim colTankContained As MUSTER.Info.TankCollection
                'RaiseEvent evtFacilityInfoTankCol(colTankContained)
                For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                    dr = tbTankTable.NewRow()
                    dr("TANK_ID") = oTankInfoLocal.TankId
                    dr("TANK_INDEX") = oTankInfoLocal.TankIndex
                    dr("FACILITY_ID") = oTankInfoLocal.FacilityId
                    dr("TANKSTATUS") = oTankInfoLocal.TankStatus
                    dr("DATERECEIVED") = oTankInfoLocal.DateReceived
                    dr("MANIFOLD") = oTankInfoLocal.Manifold
                    dr("COMPARTMENT") = oTankInfoLocal.Compartment
                    dr("TANKCAPACITY") = oTankInfoLocal.TankCapacity
                    dr("SUBSTANCE") = oTankInfoLocal.Substance
                    dr("CASNUMBER") = oTankInfoLocal.CASNumber
                    dr("DATELASTUSED") = oTankInfoLocal.DateLastUsed
                    dr("DATECLOSURERECEIVED") = oTankInfoLocal.DateClosureReceived
                    dr("DATECLOSED") = oTankInfoLocal.DateClosed
                    dr("CLOSURESTATUSDESC") = oTankInfoLocal.ClosureStatusDesc
                    dr("CLOSURETYPE") = oTankInfoLocal.ClosureType
                    dr("INERTMATERIAL") = oTankInfoLocal.InertMaterial
                    dr("TANKMATDESC") = oTankInfoLocal.TankMatDesc
                    dr("TANKMODDESC") = oTankInfoLocal.TankModDesc
                    dr("TANKOTHERMATERIAL") = oTankInfoLocal.TankOtherMaterial
                    dr("OVERFILLINSTALLED") = oTankInfoLocal.OverFillInstalled
                    dr("SPILLINSTALLED") = oTankInfoLocal.SpillInstalled
                    dr("LICENSEEID") = oTankInfoLocal.LicenseeID
                    dr("CONTRACTORID") = oTankInfoLocal.ContractorID
                    dr("DATESIGNED") = oTankInfoLocal.DateSigned
                    dr("SMALLDELIVERY") = oTankInfoLocal.SmallDelivery
                    dr("TANKEMERGEN") = oTankInfoLocal.TankEmergen
                    dr("PLANNEDINSTDATE") = oTankInfoLocal.PlannedInstDate
                    dr("LASTTCPDATE") = oTankInfoLocal.LastTCPDate
                    dr("LINEDINTERIORINSTALLDATE") = oTankInfoLocal.LinedInteriorInstallDate
                    dr("LINEDINTERIORINSPECTDATE") = oTankInfoLocal.LinedInteriorInspectDate
                    dr("TCPINSTALLDATE") = oTankInfoLocal.TCPInstallDate
                    dr("TTTDATE") = oTankInfoLocal.TTTDate
                    dr("TANKLD") = oTankInfoLocal.TankLD
                    dr("OVERFILLTYPE") = oTankInfoLocal.OverFillType
                    dr("PROHIBITION") = oTankInfoLocal.Prohibition
                    dr("REVOKEREASON") = oTankInfoLocal.RevokeReason
                    dr("REVOKEDATE") = oTankInfoLocal.RevokeDate
                    dr("DatePhysicallyTagged") = oTankInfoLocal.DatePhysicallyTagged
                    dr("DROPTUBE") = oTankInfoLocal.DropTube
                    dr("TANKCPTYPE") = oTankInfoLocal.TankCPType
                    dr("PLACEDINSERVICEDATE") = oTankInfoLocal.PlacedInServiceDate
                    dr("TANKTYPES") = oTankInfoLocal.TankTypes
                    dr("TANKLOCATION_DESCRIPTION") = oTankInfoLocal.TankLocationDescription
                    dr("TANKMANUFACTURER") = oTankInfoLocal.TankManufacturer
                    dr("DELETED") = oTankInfoLocal.Deleted
                    tbTankTable.Rows.Add(dr)
                Next
                Return tbTankTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetDistinctDataTableListItems(ByVal MasterDtTable As DataTable) As DataTable
            Dim FilteredItems As New DataTable
            Dim strRepeatedText As String = String.Empty
            Dim drRow As DataRow
            Dim drNewRow As DataRow

            GetDistinctDataTableListItems = Nothing
            FilteredItems.Columns.Add("PROPERTY_NAME")
            FilteredItems.Columns.Add("PROPERTY_ID")
            Try
                For Each drRow In MasterDtTable.Rows
                    If drRow.Item("PROPERTY_NAME").ToString <> strRepeatedText Then
                        drNewRow = FilteredItems.NewRow
                        drNewRow("PROPERTY_NAME") = drRow.Item("PROPERTY_NAME")
                        drNewRow("PROPERTY_ID") = drRow.Item("PROPERTY_ID")
                        FilteredItems.Rows.Add(drNewRow)
                        strRepeatedText = drRow.Item("PROPERTY_NAME").ToString
                    End If
                Next
                GetDistinctDataTableListItems = FilteredItems
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        Public Function TankCAPTable(ByVal nFACILITYID As Integer) As DataTable
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            Dim dr As DataRow
            Dim tbTankTable As New DataTable
            Try
                tbTankTable.Columns.Add("FACILITY ID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANK ID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("TANK SITE ID", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("STATUS", Type.GetType("System.String"))
                tbTankTable.Columns.Add("INSTALLED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("CAPACITY", Type.GetType("System.Int64"))
                tbTankTable.Columns.Add("SUBSTANCE", Type.GetType("System.String"))
                tbTankTable.Columns.Add("CP DATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("TT DATE", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("LI INSTALL", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("LI INSPECTED", Type.GetType("System.DateTime"))
                tbTankTable.Columns.Add("TANKLD", Type.GetType("System.String"))
                tbTankTable.Columns.Add("TANKMODDESC", Type.GetType("System.String"))
                tbTankTable.Columns.Add("TCPINSTALLDATE", Type.GetType("System.DateTime"))

                'Dim colTankContained As MUSTER.Info.TankCollection
                'RaiseEvent evtFacilityInfoTankColByFacilityID(nFACILITYID, colTankContained)
                For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                    dr = tbTankTable.NewRow()
                    If Not (oTankInfoLocal.Deleted) And nFACILITYID = oTankInfoLocal.FacilityId Then
                        oTankInfo = oTankInfoLocal
                        dr("FACILITY ID") = oTankInfoLocal.FacilityId
                        dr("TANK ID") = oTankInfoLocal.TankId
                        dr("TANK SITE ID") = oTankInfoLocal.TankIndex
                        dr("STATUS") = IIf(oProperty.Retrieve(oTankInfoLocal.TankStatus).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankStatus).Name)
                        dr("INSTALLED") = IIf(Date.Compare(oTankInfoLocal.DateInstalledTank, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.DateInstalledTank)
                        dr("CAPACITY") = oTankInfoLocal.TankCapacity
                        dr("SUBSTANCE") = IIf(oProperty.Retrieve(oTankInfoLocal.Substance).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.Substance).Name)
                        dr("CP DATE") = IIf(Date.Compare(oTankInfoLocal.LastTCPDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.LastTCPDate)
                        dr("TT DATE") = IIf(Date.Compare(oTankInfoLocal.TTTDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.TTTDate)
                        dr("LI INSTALL") = IIf(Date.Compare(oTankInfoLocal.LinedInteriorInstallDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.LinedInteriorInstallDate)
                        dr("LI INSPECTED") = IIf(Date.Compare(oTankInfoLocal.LinedInteriorInspectDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.LinedInteriorInspectDate)
                        dr("TANKLD") = IIf(oProperty.Retrieve(oTankInfoLocal.TankLD).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankLD).Name)
                        dr("TANKMODDESC") = IIf(oProperty.Retrieve(oTankInfoLocal.TankModDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankModDesc).Name)
                        dr("TCPINSTALLDATE") = IIf(Date.Compare(oTankInfoLocal.TCPInstallDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.TCPInstallDate)
                        tbTankTable.Rows.Add(dr)
                    End If
                Next
                tbTankTable.DefaultView.Sort = "TANK SITE ID"
                Return tbTankTable

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function CheckCAPStatus()
        '    Try
        '        RaiseEvent evtCAPStatusfromTank(Me.FacilityId)
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Public Function TanksTable(ByVal facID As Integer) As DataTable
            Dim oTankInfoLocal As MUSTER.Info.TankInfo
            Dim dr As DataRow
            Dim index As Integer
            Dim capacity As Int64 = 0
            Dim substance As String = ""
            Dim dtTankTable As New DataTable
            Dim dtCompTable As New DataTable
            Try
                'dtTankTable.Columns.Add("Facility_ID", Type.GetType("System.Int64"))
                dtTankTable.Columns.Add("Tank_ID", Type.GetType("System.Int64"))
                dtTankTable.Columns.Add("Tank Site ID", Type.GetType("System.Int64"))
                dtTankTable.Columns.Add("Status", Type.GetType("System.String"))
                dtTankTable.Columns.Add("Installed", Type.GetType("System.DateTime"))
                dtTankTable.Columns.Add("Last Used", Type.GetType("System.DateTime"))
                dtTankTable.Columns.Add("Substance", Type.GetType("System.String"))
                dtTankTable.Columns.Add("Capacity", Type.GetType("System.Int64"))
                dtTankTable.Columns.Add("Material", Type.GetType("System.String"))
                dtTankTable.Columns.Add("Sec Option", Type.GetType("System.String"))
                dtTankTable.Columns.Add("CP Type", Type.GetType("System.String"))
                dtTankTable.Columns.Add("LD", Type.GetType("System.String"))

                'Dim colTankContained As MUSTER.Info.TankCollection
                'RaiseEvent evtFacilityInfoTankColByFacilityID(facID, colTankContained)
                For Each oTankInfoLocal In oFacilityInfo.TankCollection.Values
                    If oTankInfoLocal.FacilityId = facID And Not (oTankInfoLocal.Deleted) Then
                        dr = dtTankTable.NewRow()
                        'dr("Facility_ID") = oTankInfoLocal.FacilityId
                        dr("Tank_ID") = oTankInfoLocal.TankId
                        dr("Tank Site ID") = oTankInfoLocal.TankIndex
                        dr("Status") = IIf(oProperty.Retrieve(oTankInfoLocal.TankStatus).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankStatus).Name)
                        dr("Installed") = IIf(Date.Compare(oTankInfoLocal.DateInstalledTank, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.DateInstalledTank)
                        dr("Last Used") = IIf(Date.Compare(oTankInfoLocal.DateLastUsed, CDate("01/01/0001")) = 0, System.DBNull.Value, oTankInfoLocal.DateLastUsed)
                        dtCompTable = oTankCompartment.CompartmentsTable(oTankInfoLocal.TankId)
                        Dim drComp As DataRow
                        For Each drComp In dtCompTable.Rows
                            substance += IIf(CType(drComp.Item("Substance"), String) Is Nothing, String.Empty, CType(drComp.Item("Substance"), String)) + ","
                            capacity += CType(drComp.Item("Capacity"), Int64)
                        Next
                        substance = substance.TrimStart(",")
                        substance = substance.TrimEnd(",")
                        dr("Substance") = substance
                        dr("Capacity") = capacity
                        dr("Material") = IIf(oProperty.Retrieve(oTankInfoLocal.TankMatDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankMatDesc).Name)
                        dr("Sec Option") = IIf(oProperty.Retrieve(oTankInfoLocal.TankModDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankModDesc).Name)
                        dr("CP Type") = IIf(oProperty.Retrieve(oTankInfoLocal.TankCPType).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankCPType).Name)
                        dr("LD") = IIf(oProperty.Retrieve(oTankInfoLocal.TankLD).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oTankInfoLocal.TankLD).Name)
                        dtTankTable.Rows.Add(dr)
                        substance = ""
                        capacity = 0
                    End If
                Next
                Return dtTankTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub CopyTankProfile(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try
                'Warn about the consequences of deleting/deactivating a tank
                If MsgBox("Copying the Tank Profile will copy pipe/s associated with the tank." + vbCrLf + vbCrLf + "Do you want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim newID As Integer
                    newID = oTankDB.CopyTankProfile(oTankInfo.TankId, moduleID, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If

                    If newID <> oTankInfo.TankId Then
                        'RaiseEvent evtTankErr("The tank profile was copied successfully")
                        MsgBox("The tank profile was copied successfully")
                        Retrieve(oFacilityInfo, oTankInfo.FacilityId, newID)
                    Else
                        'RaiseEvent evtTankErr("Error copying tank profile")
                        MsgBox("Error copying tank profile")
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#End Region
#Region "Event Handlers"
        'Private Sub TanksChanged() Handles colTank.TankColChanged
        '    RaiseEvent evtTanksChanged(Me.colIsDirty)
        'End Sub
        Private Sub TankChanged(ByVal bolValue As Boolean) Handles oTankInfo.evtTankInfoChanged
            RaiseEvent evtTankChanged(bolValue)
        End Sub
        'Private Sub TankCommentsChanged(ByVal bolValue As Boolean) Handles oComments.InfoBecameDirty
        '    RaiseEvent evtTankCommentsChanged(bolValue)
        'End Sub
        'Compartment events
        'Private Sub CAPStatusfromPipe(ByVal nfacID As Integer) Handles oTankCompartment.evtCAPStatusPipe
        '    RaiseEvent evtCAPStatusfromPipe(nfacID)
        'End Sub
        'Private Sub PipeCommentsChanged(ByVal bolValue As Boolean) Handles oTankCompartment.evtPipeCommentsChanged
        '    RaiseEvent evtPipeCommentsChanged(bolValue)
        'End Sub
        Private Sub CompartmentChanged(ByVal bolValue As Boolean) Handles oTankCompartment.evtCompartmentChanged
            RaiseEvent evtTankChanged(bolValue)
        End Sub
        Private Sub CompartmentErr(ByVal MsgStr As String) Handles oTankCompartment.evtCompartmentErr
            RaiseEvent evtTankErr(MsgStr)
        End Sub
        'Added by kiran
        'Private Sub TankCompartmentCol(ByVal TankID As Integer, ByVal CompartmentCol As MUSTER.Info.CompartmentCollection) Handles oTankCompartment.evtCompColTank
        '    'Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    'Try
        '    '    oTankInfoLocal = colTank.Item(TankID)
        '    '    If Not (oTankInfoLocal Is Nothing) Then
        '    '        oTankInfoLocal.CompartmentCollection = CompartmentCol
        '    '    End If
        '    'Catch ex As Exception
        '    '    If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '    '    Throw ex
        '    'End Try
        '    RaiseEvent evtCompartmentCol(TankID, CompartmentCol, oTankInfo.FacilityId)
        'End Sub
        'Private Sub TankCommentsCol(ByVal TankID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oComments.evtCommentColTanks
        '    'Dim oTankInfoLocal As MUSTER.Info.TankInfo
        '    'Try
        '    '    oTankInfoLocal = colTank.Item(TankID)
        '    '    If Not (oTankInfoLocal Is Nothing) Then
        '    '        oTankInfoLocal.commentsCollection = commentsCol
        '    '    End If
        '    'Catch ex As Exception
        '    '    If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '    '    Throw ex
        '    'End Try
        '    RaiseEvent evtCommentsCol(TankID, commentsCol, oTankInfo.FacilityId)
        'End Sub
        'Private Sub pipeCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal tankID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oTankCompartment.evtPipeCommentsCol
        '    RaiseEvent evtPipeCommentsCol(pipeID, compID, tankID, oTankInfo.FacilityId, commentsCol)
        'End Sub
        'Private Sub PipesColTank(ByVal TnkID As Integer, ByVal pipesCol As MUSTER.Info.PipesCollection) Handles oTankCompartment.evtPipeColTank
        '    'Dim oCompartmentInfoLocal As MUSTER.Info.CompartmentInfo
        '    'Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
        '    Try
        '        '    For Each oCompartmentInfoLocal In oTankInfo.CompartmentCollection.Values 'oTankCompartment.CompartmentCollection.Values

        '        '        For Each oPipeInfoLocal In pipesCol.Values
        '        '            If oPipeInfoLocal.CompartmentNumber = oCompartmentInfoLocal.COMPARTMENTNumber Then
        '        '                oPipeInfoLocal.TankSiteID = oCompartmentInfoLocal.TankSiteID
        '        '                If oPipeInfoLocal.FacilityID = 0 Then
        '        '                    oPipeInfoLocal.FacilityID = oCompartmentInfoLocal.FacilityId
        '        '                End If
        '        '                'oPipeInfoLocal.CompartmentID = oCompartmentInfoLocal.ID
        '        '                'oPipeInfoLocal.CompartmentNumber = oCompartmentInfo.COMPARTMENTNumber
        '        '                oPipeInfoLocal.CompartmentSubstance = oCompartmentInfoLocal.Substance
        '        '                oPipeInfoLocal.CompartmentCERCLA = oCompartmentInfoLocal.CCERCLA
        '        '                oPipeInfoLocal.CompartmentFuelType = oCompartmentInfoLocal.FuelTypeId
        '        '                oCompartmentInfoLocal.pipesCollection.Add(oPipeInfoLocal)
        '        '            End If
        '        '        Next
        '        '    Next
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'end changes
        'Private Sub CompInfoTank(ByVal compartmentInfo As MUSTER.Info.CompartmentInfo, ByVal strDesc As String) Handles oTankCompartment.evtCompInfoTank
        '    RaiseEvent evtCompInfoTank(compartmentInfo, strDesc)
        'End Sub
        'Private Sub PipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String) Handles oTankCompartment.evtPipeInfoCompartment
        '    RaiseEvent evtPipeInfoCompartment(pipeInfo, strDesc)
        'End Sub
        'Private Sub CompInfoCompID(ByVal cmpID As String) Handles oTankCompartment.evtCompInfoCompID
        '    Try
        '        oTankInfo.CompartmentCollection.Remove(cmpID)
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub TankInfoCompCol(ByRef colComp As MUSTER.Info.CompartmentCollection) Handles oTankCompartment.evtTankInfoCompCol
        '    colComp = oTankInfo.CompartmentCollection
        'End Sub
        'Private Sub CompartmentChangeKey(ByVal oldID As String, ByVal newID As String) Handles oTankCompartment.evtCompartmentChangeKey
        '    Try
        '        If oTankInfo.CompartmentCollection.Contains(oldID) Then
        '            oTankInfo.CompartmentCollection.ChangeKey(oldID, newID)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub SyncPipeInCol(ByVal pipeInfo As MUSTER.Info.PipeInfo) Handles oTankCompartment.evtSyncPipeInCol
        '    Try
        '        For Each tank As MUSTER.Info.TankInfo In oFacilityInfo.TankCollection.Values
        '            For Each pipe As MUSTER.Info.PipeInfo In tank.pipesCollection.Values
        '                If pipe.PipeID = pipeInfo.PipeID Then
        '                    pipe.ALLDTest = pipeInfo.ALLDTest
        '                    pipe.ALLDTestDate = pipeInfo.ALLDTestDate
        '                    pipe.ALLDType = pipeInfo.ALLDType
        '                    pipe.CASNumber = pipeInfo.CASNumber
        '                    pipe.ClosureStatusDesc = pipeInfo.ClosureStatusDesc
        '                    pipe.ClosureType = pipeInfo.ClosureType
        '                    pipe.CompPrimary = pipeInfo.CompPrimary
        '                    pipe.CompSecondary = pipeInfo.CompSecondary
        '                    pipe.ContainSumpDisp = pipeInfo.ContainSumpDisp
        '                    pipe.ContainSumpTank = pipeInfo.ContainSumpTank
        '                    pipe.ContractorID = pipeInfo.ContractorID
        '                    pipe.DateClosed = pipeInfo.DateClosed
        '                    pipe.DateClosureRecd = pipeInfo.DateClosureRecd
        '                    pipe.DateLastUsed = pipeInfo.DateLastUsed
        '                    pipe.DateRecd = pipeInfo.DateRecd
        '                    pipe.DateSigned = pipeInfo.DateSigned
        '                    pipe.Deleted = pipeInfo.Deleted
        '                    pipe.FacilityPowerOff = pipeInfo.FacilityPowerOff
        '                    pipe.InertMaterial = pipeInfo.InertMaterial
        '                    pipe.LCPInstallDate = pipeInfo.LCPInstallDate
        '                    pipe.LicenseeID = pipeInfo.LicenseeID
        '                    pipe.LTTDate = pipeInfo.LTTDate
        '                    pipe.ModifiedBy = pipeInfo.ModifiedBy
        '                    pipe.ModifiedOn = pipeInfo.ModifiedOn
        '                    pipe.PipeCPInstalledDate = pipeInfo.PipeCPInstalledDate
        '                    pipe.PipeCPTest = pipeInfo.PipeCPTest
        '                    pipe.PipeCPType = pipeInfo.PipeCPType
        '                    pipe.PipeInstallationPlannedFor = pipeInfo.PipeInstallationPlannedFor
        '                    pipe.PipeInstallDate = pipeInfo.PipeInstallDate
        '                    pipe.PipeLD = pipeInfo.PipeLD
        '                    pipe.PipeManufacturer = pipeInfo.PipeManufacturer
        '                    pipe.PipeMatDesc = pipeInfo.PipeMatDesc
        '                    pipe.PipeModDesc = pipeInfo.PipeModDesc
        '                    pipe.PipeOtherMaterial = pipeInfo.PipeOtherMaterial
        '                    pipe.PipeStatusDesc = pipeInfo.PipeStatusDesc
        '                    pipe.PipeTypeDesc = pipeInfo.PipeTypeDesc
        '                    pipe.PipingComments = pipeInfo.PipingComments
        '                    pipe.PlacedInServiceDate = pipeInfo.PlacedInServiceDate
        '                    pipe.SubstanceComments = pipeInfo.SubstanceComments
        '                    pipe.SubstanceDesc = pipeInfo.SubstanceDesc
        '                    pipe.TermCPInstalledDate = pipeInfo.TermCPInstalledDate
        '                    pipe.TermCPLastTested = pipeInfo.TermCPLastTested
        '                    pipe.TermCPTypeDisp = pipeInfo.TermCPTypeDisp
        '                    pipe.TermCPTypeTank = pipeInfo.TermCPTypeTank
        '                    pipe.TermTypeDisp = pipeInfo.TermTypeDisp
        '                    pipe.TermTypeTank = pipeInfo.TermTypeTank
        '                    pipe.Archive()
        '                End If
        '            Next
        '        Next
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub FacCapStatusFromPipe(ByVal facID As Integer) Handles oTankCompartment.evtFacCapStatus
        '    If oTankInfo.FacCapStatus <> oFacilityInfo.CapStatus Then
        '        oFacilityInfo.CapStatusOriginal = oTankInfo.FacCapStatus
        '        oFacilityInfo.CapStatus = oTankInfo.FacCapStatus
        '        RaiseEvent evtCAPStatusfromTank(facID)
        '    End If
        'End Sub
#End Region
    End Class
End Namespace
