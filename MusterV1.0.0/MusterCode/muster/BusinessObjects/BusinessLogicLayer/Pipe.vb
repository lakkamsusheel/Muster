'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Pipe
'   Provides the operations required to manipulate an Pipe object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MNR     12/15/04    Original class definition
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         EN      01/03/05    Added colIsDirty Property
'   1.3         MNR     01/04/05    Added function ValidateData(Optional ByVal [module] As String = "Registration")
'   1.4         MNR     01/05/05    Added set value for colIsDirty property
'   1.5         MNR     01/05/05    updated ValidateData(...) to handle UI validations
'   1.6         EN      01/06/05    Added Lookup Operation/ExtenalEventHandler/RaisingEvents and Events.
'                                   Added Properties to Collection in Set method.
'                                   Modified Reset Method.
'   1.7         MNR     01/06/05    Deleted adding Properties to collection in set method
'                                   (changes to the info object reflects in the collection
'                                    so, there is no need to update the collection everytime)
'   1.8         MNR     01/12/05    Modified Retrieve function to handle hirearchy
'   1.9         EN      01/20/05    Added new Events,new methods in RaisingEvents region and check the properties and raise the event back.
'   2.0         EN      01/21/05    Modified Reterieve,Remove,save,Flush methods.Added new events. 
'   2.1         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   2.2         AN      01/31/05    Adding new Property object 
'   2.3         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Pipe" type.
'                                       Also added EntityType attribute to expose Type ID.
'   2.4         AN      02/02/05    Added Comments object
'   2.5         JVC2    02/16/05    Modified CHECKPIPESTATUS() per Padmaja's direction.
'                                   Added new function GetDataTable_Status_Based to return the
'                                       pipe secondary options based on the MOC and Status.
'                                   Modified PopulatePipeSecondaryOption to take an optional
'                                       secondary parameter (pipe status) to control
'                                       production of return set for secondary option.
'   2.6          MR      03/07/05    Added iAccessor Attributes to make them exposed to UI.
'   2.7          MNR     03/09/05    Implemented TOS/TOSI/POU rules
'   2.8          MNR     03/15/05    Added Load Sub
'   2.9          AB      03/16/05    Added DataAge check to the Retrieve function
'   3.0          MNR     03/16/05    Removed strSrc from events
'   3.1          KKM     03/18/05    Event for handling local CommentsCollection is added
'   3.2          MR      04/13/05    Modified Validation Error Msg for ALLD Test Date.
'   3.3   Thomas Franey  02/23/09    Added parent pipe Logic.
'   3.3   Thomas Franey  02/24/09    Added Properties HasParent & has extension to extend parent pipe logic
'   3.4   Thomas Franey  03/05/09    Added a Get Pipe Extensions function to list a datstable of extended pipes by parent id
'
' Function          Description
' Retrieve(ID)      Returns an Info Object requested by the int arg ID  
' Save()            Saves the Info Object
' Validate(Optional ByVal [module] As String = "Registration")
'                   Validates the object data according to the data validation rules
'                   respective to the module and returns true if success else false
' GetAll()          Returns a collection with all the relevant information
' Add(ID)           Adds an Info Object identified by the int arg ID
'                   to the Pipes Collection
' Add(Entity)       Adds the Entity passed as an argument
'                   to the Pipes Collection
' Remove(ID)        Removes an Info Object identified by the int arg ID
'                   from the Pipes Collection
' Remove(Entity)    Removes the Entity passed as an argument
'                   from the Pipes Collection
' Flush()           Marshalls all modified/new Onwer Info objects in the
'                   Pipe Collection to the repository
' EntityTable()     Returns a datatable containing all columns for the Entity
'                   objects in the Pipes Collection
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pPipe
#Region "Public Events"
        Public Event evtFacCapStatus(ByVal facID As Integer)
        'Public Event evtPipeCommentsChanged(ByVal bolValue As Boolean)
        'Public Event evtPipeSaved(ByVal bolStatus As Boolean)
        Public Event evtPipeErr(ByVal MsgStr As String)
        'Public Event eSecondaryOption(ByRef dtPipeSecondaryOption As DataTable)
        'Public Event ePipeReleaseDetection1(ByRef dtPipeReleaseDetection1 As DataTable)
        'Public Event ePipeReleaseDetection2(ByRef dtPipeReleaseDetection2 As DataTable)
        Public Event evtPipesChanged(ByVal BolState As Boolean)
        Public Event evtPipeChanged(ByVal BolState As Boolean)
        'Public Event ePipeReleaseDetectionstatus1(ByVal nIndex1 As Integer, ByVal BolReleaseDetection1 As Boolean)
        'Public Event ePipeReleaseDetectionstatus2(ByVal nIndex2 As Integer, ByVal BolReleaseDetection2 As Boolean)
        'Public Event ePickPipeLeakDetectorTeststatus(ByVal bolPickPipeLeakDetectorTest As Boolean)
        'Public Event ePipeReleaseDetectiontext1(ByVal strText As String)
        'Public Event ePipeReleaseDetectiontext2(ByVal strText As String)
        'Public Event ePipeCPTypeEnable(ByVal BolState As Boolean, ByVal nIndex As Integer)
        'Public Event edtPickPipeCPLastTest(ByVal BolState As Boolean)
        'Public Event ePickPipeCPInstalled(ByVal BolState As Boolean)
        'Public Event ePickPipeTightnessTest(ByVal BolState As Boolean)
        'Public Event ePickPipePlannedInstallation(ByVal BolState As Boolean)
        'Public Event ePickPipeInstalled(ByVal BolState As Boolean)
        'Public Event ePickDatePipePlacedInService(ByVal BolState As Boolean)
        'Public Event ecmbPipeMaterial(ByVal BolState As Boolean)
        'Public Event echkPipeTankUsedInEmergency(ByVal BolState As Boolean)
        'Public Event ecmbPipeOptions(ByVal BolState As Boolean)
        'Public Event ecmbPipeCPType(ByVal BolState As Boolean)
        'Public Event echkPipeSumpsAtDispenser(ByVal BolState As Boolean)
        'Public Event echkPipeSumpsAtTank(ByVal BolState As Boolean)
        'Public Event ecmbPipeTerminationDispenserType(ByVal BolState As Boolean)
        'Public Event ecmbPipeTerminationTankType(ByVal BolState As Boolean)
        'Public Event ecmbPipeType(ByVal BolState As Boolean)
        'Public Event ecmbPipeReleaseDetection1(ByVal BolState As Boolean)
        'Public Event ecmbPipeReleaseDetection2(ByVal BolState As Boolean)
        'Public Event edtPickPipeSigned(ByVal BolState As Boolean)
        'Public Event ecmbPipeManufacturerID(ByVal BolState As Boolean)
        'Public Event edtPickPipeTerminationCPInstalled(ByVal BolState As Boolean)
        'Public Event edtPickPipeTerminationCPLastTested(ByVal BolState As Boolean)
        'Public Event edtPickPipeLeakDetectorTest(ByVal BolState As Boolean)
        'Public Event edtPickPipeLastUsed(ByVal BolState As Boolean)
        'Public Event edtPickPipeLastUsedFocus(ByVal strSrc As String)
        'Public Event ecmbPipeTerminationDispenserCPType(ByVal BolState As Boolean)
        'Public Event ecmbPipeTerminationTankCPType(ByVal Bolstate As Boolean)
        'Public Event enabledisablepipecontrols(ByVal Bolstate As Boolean)
        'Public Event ecmbPipeClosureType(ByVal BolState As Boolean)
        'Public Event ecmbPipeInertFill(ByVal BolState As Boolean)
        'Public Event ePipeCPTestMessage(ByVal str As String)
        'Public Event edtPickPipeTerminationCPLastTestedMessage(ByVal str As String)
        'Public Event ePickPipeTightnessTestMessage(ByVal str As String)
        'Public Event ePickPipeCPLastTest(ByVal BolState As Boolean)
        'Public Event edtPickPipeLeakDetectorTestMessage(ByVal str As String)
        'Public Event evtPipeCAPChanged(ByVal nVal As Integer)
        'added by kiran
        'Public Event evtPipeColCompartment(ByVal CompartmentID As String, ByVal pipeCol As MUSTER.Info.PipesCollection)
        'Public Event evtPipesCommentsCol(ByVal pipeID As Integer, ByVal compID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        'end changes
        'Public Event evtPipeInfoCompartment(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal strDesc As String)
        'Public Event evtCompInfoPipeCol(ByRef colPipe As MUSTER.Info.PipesCollection)
        'Public Event evtPipeChangeKey(ByVal oldID As String, ByVal newID As String)
        'Public Event evtSyncPipeInCol(ByVal pipeInfo As MUSTER.Info.PipeInfo)
#End Region
#Region "Private Member Variables"
        'Private oCompartmentInfo As MUSTER.Info.CompartmentInfo
        Private oTankInfo As MUSTER.Info.TankInfo
        Private oCompInfo As MUSTER.Info.CompartmentInfo
        Private WithEvents oPipeInfo As MUSTER.Info.PipeInfo
        'Private WithEvents colPipes As MUSTER.Info.PipesCollection
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
        Private WithEvents oProperty As MUSTER.BusinessLogic.pProperty
        Private oPipeDB As New MUSTER.DataAccess.PipeDB
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Pipe").ID
        Private bolCheckCAPSTATE As Boolean = False
#End Region
#Region "Constructors"
        Public Sub New(Optional ByRef TankInfo As MUSTER.Info.TankInfo = Nothing)
            If TankInfo Is Nothing Then
                oTankInfo = New MUSTER.Info.TankInfo
            Else
                oTankInfo = TankInfo
            End If
            oCompInfo = New MUSTER.Info.CompartmentInfo
            oPipeInfo = New MUSTER.Info.PipeInfo
            'colPipes = New MUSTER.Info.PipesCollection
            oProperty = New MUSTER.BusinessLogic.pProperty
            oComments = New MUSTER.BusinessLogic.pComments
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return oPipeInfo.ID
            End Get
            Set(ByVal Value As String)
                oPipeInfo.ID = Value
            End Set
        End Property

        Public Property ParentPipeID() As Integer
            Get
                Return oPipeInfo.ParentPipeID
            End Get

            Set(ByVal Value As Integer)
                oPipeInfo.ParentPipeID = Value
            End Set
        End Property

        Public ReadOnly Property HasExtensions() As Boolean
            Get
                Return oPipeInfo.HasExtensions
            End Get
        End Property

        Public ReadOnly Property HasParent() As Boolean
            Get
                Return oPipeInfo.HasParent
            End Get
        End Property

        Public Property PipeID() As Integer
            Get
                Return oPipeInfo.PipeID
            End Get
            Set(ByVal Value As Integer)
                'Dim oldID As Integer = oPipeInfo.ID
                oPipeInfo.PipeID = Value
                'colPipes.ChangeKey(oldID, oPipeInfo.ID)
            End Set
        End Property
        Public Property Index() As Integer
            Get
                Return oPipeInfo.Index
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.Index = Value
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property FacilityID() As Integer
            Get
                Return oPipeInfo.FacilityID
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.FacilityID = Value
            End Set
        End Property
        Public Property TankID() As Integer
            Get
                Return oPipeInfo.TankID
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.TankID = Value
            End Set
        End Property
        Public Property ALLDTest() As String
            Get
                Return oPipeInfo.ALLDTest
            End Get
            Set(ByVal Value As String)
                oPipeInfo.ALLDTest = Value
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property ALLDTestDate() As Date
            Get
                Return oPipeInfo.ALLDTestDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.ALLDTestDate = Value
                'CheckALLDTestDate(oPipeInfo.ALLDTestDate)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property CASNumber() As Integer
            Get
                Return oPipeInfo.CASNumber
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CASNumber = Value
            End Set
        End Property
        Public Property ClosureStatusDesc() As Integer
            Get
                Return oPipeInfo.ClosureStatusDesc
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.ClosureStatusDesc = Value
            End Set
        End Property
        Public ReadOnly Property ClosureStatusDesc_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.ClosureStatusDesc)
            End Get
        End Property
        Public Property ClosureType() As Integer
            Get
                Return oPipeInfo.ClosureType
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.ClosureType = Value
            End Set
        End Property
        Public ReadOnly Property ClosureType_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(oPipeInfo.ClosureType)
            End Get
        End Property
        Public Property CompPrimary() As Integer
            Get
                Return oPipeInfo.CompPrimary
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CompPrimary = Value
            End Set
        End Property
        Public Property CompSecondary() As Integer
            Get
                Return oPipeInfo.CompSecondary
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CompSecondary = Value
            End Set
        End Property
        Public Property ContainSumpDisp() As Boolean
            Get
                Return oPipeInfo.ContainSumpDisp
            End Get
            Set(ByVal Value As Boolean)
                oPipeInfo.ContainSumpDisp = Value
            End Set
        End Property
        Public Property ContainSumpTank() As Boolean
            Get
                Return oPipeInfo.ContainSumpTank
            End Get
            Set(ByVal Value As Boolean)
                oPipeInfo.ContainSumpTank = Value
            End Set
        End Property
        Public Property DateClosed() As Date
            Get
                Return oPipeInfo.DateClosed
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DateClosed = Value
            End Set
        End Property
        Public Property DateLastUsed() As Date
            Get
                Return oPipeInfo.DateLastUsed
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DateLastUsed = Value
                'CheckDateLastUsed(oPipeInfo.DateLastUsed)
            End Set
        End Property
        Public Property DateClosureRecd() As Date
            Get
                Return oPipeInfo.DateClosureRecd
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DateClosureRecd = Value
            End Set
        End Property
        Public Property DateRecd() As Date
            Get
                Return oPipeInfo.DateRecd
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DateRecd = Value
            End Set
        End Property
        Public Property DateSigned() As Date
            Get
                Return oPipeInfo.DateSigned
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DateSigned = Value
            End Set
        End Property
        Public Property ALLDType() As Integer
            Get
                Return oPipeInfo.ALLDType
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.ALLDType = Value
                'If oPipeInfo.ALLDType > 0 Then
                '    CheckALLDType(oPipeInfo.ALLDType)
                'End If
            End Set
        End Property
        Public ReadOnly Property ALLDType_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.ALLDType)
            End Get
        End Property
        Public Property InertMaterial() As Integer
            Get
                Return oPipeInfo.InertMaterial
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.InertMaterial = Value
            End Set
        End Property
        Public ReadOnly Property InertMaterial_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.InertMaterial)
            End Get
        End Property
        Public Property LCPInstallDate() As Date
            Get
                Return oPipeInfo.LCPInstallDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.LCPInstallDate = Value
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return oPipeInfo.LicenseeID
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.LicenseeID = Value
            End Set
        End Property
        Public Property ContractorID() As Integer
            Get
                Return oPipeInfo.ContractorID
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.ContractorID = Value
            End Set
        End Property
        Public Property LTTDate() As Date
            Get
                Return oPipeInfo.LTTDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.LTTDate = Value
                'CheckLTTDate(oPipeInfo.LTTDate)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property PipeCPTest() As Date
            Get
                Return oPipeInfo.PipeCPTest
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.PipeCPTest = Value
            End Set
        End Property
        Public Property DateShearTest() As Date
            Get
                Return oPipeInfo.DateShearTest
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DateShearTest = Value
            End Set
        End Property
        Public Property DatePipeSecInsp() As Date
            Get
                Return oPipeInfo.DatePipeSecInsp
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DatePipeSecInsp = Value
            End Set
        End Property
        Public Property DatePipeElecInsp() As Date
            Get
                Return oPipeInfo.DatePipeElecInsp
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.DatePipeElecInsp = Value
            End Set
        End Property
        Public Property PipeCPType() As Integer
            Get
                Return oPipeInfo.PipeCPType
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeCPType = Value
            End Set
        End Property
        Public ReadOnly Property PipeCPType_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeCPType)
            End Get
        End Property
        Public Property CheckCAPSTATE() As Boolean
            Get
                Return bolCheckCAPSTATE
            End Get
            Set(ByVal value As Boolean)
                bolCheckCAPSTATE = value
            End Set
        End Property
        Public Property PipeInstallDate() As Date
            Get
                Return oPipeInfo.PipeInstallDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.PipeInstallDate = Value
            End Set
        End Property
        Public Property PipeLD() As Integer
            Get
                Return oPipeInfo.PipeLD
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeLD = Value
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
                'If oPipeInfo.PipeLD = 245 Then  'line tightness testing
                '    RaiseEvent ePickPipeTightnessTest(True)
                '    'EnableDisablePipeOptions(oPipeInfo.PipeLD)
                'Else
                '    RaiseEvent ePickPipeTightnessTest(False)
                'End If
                'PopulatePipeReleaseDetection2(oPipeInfo.PipeLD)
            End Set
        End Property
        Public ReadOnly Property PipeLD_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeLD)
            End Get
        End Property
        Public Property PipeManufacturer() As Integer
            Get
                Return oPipeInfo.PipeManufacturer
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeManufacturer = Value
            End Set
        End Property
        Public ReadOnly Property PipeManufacturerOriginal() As Integer
            Get
                Return oPipeInfo.PipeManufacturerOriginal
            End Get
        End Property
        Public Property PipeMatDesc() As Integer
            Get
                Return oPipeInfo.PipeMatDesc
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeMatDesc = Value
                'CheckPipeMatDesc(oPipeInfo.PipeStatusDesc)
            End Set
        End Property
        Public ReadOnly Property PipeMatDescOriginal() As Integer
            Get
                Return oPipeInfo.PipeMatDescOriginal
            End Get
        End Property
        Public ReadOnly Property PipeMatDesc_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeMatDesc)
            End Get
        End Property
        Public Property PipeModDesc() As Integer
            Get
                Return oPipeInfo.PipeModDesc
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeModDesc = Value
                'CheckPipeModDesc(oPipeInfo.PipeModDesc)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property PipeModDesc_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeModDesc)
            End Get
        End Property
        Public Property PipeOtherMaterial() As String
            Get
                Return oPipeInfo.PipeOtherMaterial
            End Get
            Set(ByVal Value As String)
                oPipeInfo.PipeOtherMaterial = Value
            End Set
        End Property
        Public ReadOnly Property PipeOtherMaterial_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeOtherMaterial)
            End Get
        End Property
        Public Property PipeStatusDesc() As Integer
            Get
                Return oPipeInfo.PipeStatusDesc
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeStatusDesc = Value
                'Me.CheckPipeStatus(oPipeInfo.PipeStatusDesc)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property PipeStatusDesc_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeStatusDesc)
            End Get
        End Property
        Public ReadOnly Property PipeStatusDescOriginal() As Integer
            Get
                Return oPipeInfo.PipeStatusDescOriginal
            End Get
        End Property
        Public Property PipeTypeDesc() As Integer
            Get
                Return oPipeInfo.PipeTypeDesc
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.PipeTypeDesc = Value
                'CheckPipeTypeDesc(oPipeInfo.PipeTypeDesc)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property PipeTypeDesc_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.PipeTypeDesc)
            End Get
        End Property
        Public Property PipingComments() As String
            Get
                Return oPipeInfo.PipingComments
            End Get
            Set(ByVal Value As String)
                oPipeInfo.PipingComments = Value
            End Set
        End Property
        Public Property PipeInstallationPlannedFor() As Date
            Get
                Return oPipeInfo.PipeInstallationPlannedFor
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.PipeInstallationPlannedFor = Value
            End Set
        End Property
        Public Property PlacedInServiceDate() As Date
            Get
                Return oPipeInfo.PlacedInServiceDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.PlacedInServiceDate = Value
            End Set
        End Property
        Public Property SubstanceComments() As Integer
            Get
                Return oPipeInfo.SubstanceComments
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.SubstanceComments = Value
            End Set
        End Property
        Public Property SubstanceDesc() As Integer
            Get
                Return oPipeInfo.SubstanceDesc
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.SubstanceDesc = Value
            End Set
        End Property
        Public ReadOnly Property SubstanceDesc_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.SubstanceDesc)
            End Get
        End Property
        Public Property TermCPLastTested() As Date
            Get
                Return oPipeInfo.TermCPLastTested
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.TermCPLastTested = Value
                'CheckTermCPLastTested(oPipeInfo.TermCPLastTested)
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public Property TermCPTypeTank() As Integer
            Get
                Return oPipeInfo.TermCPTypeTank
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.TermCPTypeTank = Value
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property TermCPTypeTank_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.TermCPTypeTank)
            End Get
        End Property
        Public Property TermCPTypeDisp() As Integer
            Get
                Return oPipeInfo.TermCPTypeDisp
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.TermCPTypeDisp = Value
                'If bolCheckCAPSTATE Then
                '    CheckCAPStatus()
                'End If
            End Set
        End Property
        Public ReadOnly Property TermCPTypeDisp_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.TermCPTypeDisp)
            End Get
        End Property
        Public Property PipeCPInstalledDate() As Date
            Get
                Return oPipeInfo.PipeCPInstalledDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.PipeCPInstalledDate = Value
            End Set
        End Property
        Public Property TermCPInstalledDate() As Date
            Get
                Return oPipeInfo.TermCPInstalledDate
            End Get
            Set(ByVal Value As Date)
                oPipeInfo.TermCPInstalledDate = Value
            End Set
        End Property
        Public Property TermTypeDisp() As Integer
            Get
                Return oPipeInfo.TermTypeDisp
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.TermTypeDisp = Value
                'CheckPipeTermination()
            End Set
        End Property
        Public ReadOnly Property TermTypeDisp_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.TermTypeDisp)
            End Get
        End Property
        Public Property TermTypeTank() As Integer
            Get
                Return oPipeInfo.TermTypeTank
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.TermTypeTank = Value
                'CheckPipeTermination()
            End Set
        End Property
        Public ReadOnly Property TermTypeTank_Property() As MUSTER.Info.PropertyInfo
            Get
                Return Me.oProperty.Retrieve(Me.TermTypeTank)
            End Get
        End Property
        Public ReadOnly Property PipeCollection() As MUSTER.Info.PipesCollection
            Get
                Return oTankInfo.pipesCollection
            End Get
        End Property
        'Public ReadOnly Property EntityType() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xPipeInfo As MUSTER.Info.PipeInfo
                For Each xPipeInfo In oTankInfo.pipesCollection.Values
                    If xPipeInfo.IsDirty Then
                        'MsgBox("BLL F:" + oTankInfo.FacilityId.ToString + " TI:" + oTankInfo.TankIndex.ToString + " PID:" + xPipeInfo.PipeID.ToString)
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oPipeInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oPipeInfo.Deleted = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oPipeInfo.IsDirty
            End Get
            Set(ByVal value As Boolean)
                oPipeInfo.IsDirty = value
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
        Public Property CompartmentID() As String
            Get
                Return oPipeInfo.CompartmentID
            End Get
            Set(ByVal Value As String)
                oPipeInfo.CompartmentID = Value
            End Set
        End Property
        Public Property CompartmentNumber() As Integer
            Get
                Return oPipeInfo.CompartmentNumber
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CompartmentNumber = Value
            End Set
        End Property
        Public Property CompartmentSubstance() As Integer
            Get
                Return oPipeInfo.CompartmentSubstance
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CompartmentSubstance = Value
            End Set
        End Property
        Public Property CompartmentCERCLA() As Integer
            Get
                Return oPipeInfo.CompartmentCERCLA
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CompartmentCERCLA = Value
            End Set
        End Property
        Public Property CompartmentFuelType() As Integer
            Get
                Return oPipeInfo.CompartmentFuelType
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.CompartmentFuelType = Value
            End Set
        End Property
        Public Property TankSiteID() As Integer
            Get
                Return oPipeInfo.TankSiteID
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.TankSiteID = Value
            End Set
        End Property
        Public Property AttachedPipeID() As Integer
            Get
                Return oPipeInfo.AttachedPipeID
            End Get
            Set(ByVal Value As Integer)
                oPipeInfo.AttachedPipeID = Value
            End Set
        End Property
        Public ReadOnly Property CompartmentSubstanceDesc() As String
            Get
                Return oProperty.Retrieve(CompartmentSubstance).Name
            End Get
        End Property
        Public ReadOnly Property CompartmentCERCLADesc() As String
            Get
                ' TODO - have to put CERCLA values in property_master table
                'Return oProperty.Retrieve(CompartmentCERCLA).Name
                Dim dsLocal As DataSet
                Dim strSQL As String = "SELECT * FROM vCERLATYPE WHERE CASRN = '" + CompartmentCERCLA.ToString + "'"
                dsLocal = oPipeDB.DBGetDS(strSQL)
                If dsLocal.Tables(0).Rows.Count > 0 Then
                    Return dsLocal.Tables(0).Rows(0).Item("Substance").ToString
                Else
                    Return String.Empty
                End If
            End Get
        End Property
        Public ReadOnly Property CompartmentFuelTypeDesc() As String
            Get
                Return oProperty.Retrieve(CompartmentFuelType).Name
            End Get
        End Property
        Public Property Pipe() As MUSTER.Info.PipeInfo
            Get
                Return oPipeInfo
            End Get
            Set(ByVal Value As MUSTER.Info.PipeInfo)
                oPipeInfo = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oPipeInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oPipeInfo.CreatedBy = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oPipeInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oPipeInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oPipeInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oPipeInfo.ModifiedOn
            End Get
        End Property
        Public Property TankInfo() As MUSTER.Info.TankInfo
            Get
                Return oTankInfo
            End Get
            Set(ByVal Value As MUSTER.Info.TankInfo)
                oTankInfo = Value
            End Set
        End Property
        Public ReadOnly Property FacCapStatus() As Integer
            Get
                Return oTankInfo.FacCapStatus
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByRef TankInfo As MUSTER.Info.TankInfo, ByRef ds As DataSet, ByVal [Module] As String)
            Dim dr As DataRow
            oTankInfo = TankInfo
            Try
                If ds.Tables("Pipes").Rows.Count > 0 Then
                    For Each dr In ds.Tables("Pipes").Select("FACILITY_ID = " + oTankInfo.FacilityId.ToString + "AND COMPARTMENTS_PIPES_TANKID = " + oTankInfo.TankId.ToString)
                        oPipeInfo = New MUSTER.Info.PipeInfo(dr)
                        oPipeInfo.TankSiteID = oTankInfo.TankIndex
                        oPipeInfo.FacilityPowerOff = oTankInfo.FacilityPowerOff
                        oPipeInfo.FacCapStatus = oTankInfo.FacCapStatus
                        oTankInfo.pipesCollection.Add(oPipeInfo)
                        'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "ADD")
                        'oComments.Load(ds, [Module], nEntityTypeID, oPipeInfo.PipeID)
                        ds.Tables("Pipes").Rows.Remove(dr)
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        'Added By Elango on Jan 6 2005
        Public Function getCompartmentPipeRelationship(Optional ByVal nTank_ID As Integer = 0, Optional ByVal nCompartment_Number As Integer = 0, Optional ByVal nPipe_ID As Integer = 0) As DataSet
            Try
                getCompartmentPipeRelationship = Nothing
                Dim strSQl As String = String.Empty
                Dim strWhere As String = String.Empty
                Dim dsSet As New DataSet

                If nTank_ID > 0 Then
                    strWhere += " AND TANK_ID=" + nTank_ID.ToString
                End If
                If nCompartment_Number > 0 Then
                    strWhere += " AND COMPARTMENT_NUMBER=" + nCompartment_Number.ToString
                End If
                If nPipe_ID > 0 Then
                    strWhere += " AND PIPE_ID=" + nPipe_ID.ToString
                End If
                strSQl = "SELECT tblREG_COMPARTMENTS_PIPES.*, tblREG_TANK.TANK_INDEX FROM tblREG_COMPARTMENTS_PIPES, tblREG_TANK WHERE tblREG_TANK.TANK_ID=tblREG_COMPARTMENTS_PIPES.TANK_ID " + strWhere
                dsSet = oPipeDB.DBGetDS(strSQl)
                Return dsSet
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getCompartmentDetailsForPipe(Optional ByVal nPipe_ID As Integer = 0) As DataSet
            Try
                getCompartmentDetailsForPipe = Nothing
                Dim strSQL As String = String.Empty
                Dim dsSet As New DataSet
                strSQL = "SELECT A.TANK_ID, A.COMPARTMENT_NUMBER, B.TANK_INDEX, A.PIPE_ID, D.PROPERTY_NAME AS 'SUBSTANCE', E.PROPERTY_NAME AS 'FUEL_TYPE', F.CASRN, F.SUBSTANCE AS 'CERCLA_SUBSTANCE' FROM tblREG_COMPARTMENTS_PIPES AS A INNER JOIN tblREG_TANK AS B ON A.TANK_ID = B.TANK_ID INNER JOIN tblREG_COMPARTMENTS AS C ON A.TANK_ID = C.TANK_ID AND A.COMPARTMENT_NUMBER = C.COMPARTMENT_NUMBER LEFT OUTER JOIN tblSYS_PROPERTY_MASTER D ON C.SUBSTANCE = D.PROPERTY_ID LEFT OUTER JOIN tblSYS_PROPERTY_MASTER E ON C.FUEL_TYPE_ID = E.PROPERTY_ID LEFT OUTER JOIN tblREG_CERCLA F ON C.CERCLA# = F.CASRN WHERE A.PIPE_ID = " + nPipe_ID.ToString
                dsSet = oPipeDB.DBGetDS(strSQL)
                Return dsSet
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function Retrieve(ByRef TankInfo As MUSTER.Info.TankInfo, ByVal id As String, Optional ByVal compInfo As MUSTER.Info.CompartmentInfo = Nothing, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.PipeInfo
            Try
                oTankInfo = TankInfo
                If Not compInfo Is Nothing Then
                    oCompInfo = compInfo
                Else
                    Dim strArray() As String
                    strArray = id.Split("|")
                    oCompInfo = oTankInfo.CompartmentCollection.Item(strArray(0) + "|" + strArray(1))
                End If
                'Dim colPipesContained As MUSTER.Info.PipesCollection
                'RaiseEvent evtCompInfoPipeCol(colPipesContained)
                'oPipeInfo = colPipesContained.Item(id)

                oPipeInfo = oTankInfo.pipesCollection.Item(id)

                ' test for data age here...
                If IsNothing(oPipeInfo) Then
                    Add(id, showDeleted)
                Else
                    If oPipeInfo.IsDirty = False And oPipeInfo.IsAgedData = True Then
                        oTankInfo.pipesCollection.Remove(oPipeInfo)
                        'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                        Add(id, showDeleted)
                    End If
                End If

                Return oPipeInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'oComments.Clear()
            'oComments.GetByModule("", nEntityTypeID, oPipeInfo.PipeID)
            Return oPipeInfo
        End Function
        Public Function Retrieve(ByRef TankInfo As MUSTER.Info.TankInfo, ByVal id As Integer, Optional ByVal compInfo As MUSTER.Info.CompartmentInfo = Nothing, Optional ByVal showDeleted As Boolean = False, Optional ByVal strIDInfo As String = "TANK") As MUSTER.Info.PipeInfo
            Dim bolAgedData As Boolean = False
            Try
                oTankInfo = TankInfo
                If Not compInfo Is Nothing Then
                    oCompInfo = compInfo
                Else
                    'If UCase(strIDInfo) = "PIPE" Then
                    'Dim strArray() As String
                    'strArray = id.Split("|")
                    'oCompInfo = oTankInfo.CompartmentCollection.Item(strArray(0) + "|" + strArray(1))
                    'Else
                    oCompInfo = Nothing
                    'End If
                End If

                ' retrieve info only if validation succeeds no current instance
                'If Me.ValidateData() Then
                If id = 0 Then
                    Add(New MUSTER.Info.PipeInfo)
                    Exit Try
                End If
                'Dim colPipesContained As MUSTER.Info.PipesCollection
                'RaiseEvent evtCompInfoPipeCol(colPipesContained)
                Select Case UCase(strIDInfo).Trim
                    Case "PIPE"
                        For Each opipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                            If opipeinfolocal.PipeID = id Then
                                oPipeInfo = oPipeInfoLocal
                                Exit For
                            End If
                        Next
                        'oPipeInfo = oTankInfo.pipesCollection.Item(id)
                        ' test for data age here...
                        If Not oPipeInfo Is Nothing Then
                            If oPipeInfo.IsDirty = False And oPipeInfo.IsAgedData = True And IsNothing(compInfo) = False Then
                                oTankInfo.pipesCollection.Remove(oPipeInfo)
                                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                                bolAgedData = True
                            End If
                        End If

                        If oPipeInfo Is Nothing Or bolAgedData = True Then
                            Add(id, showDeleted)
                            If oPipeInfo.PipeID < 0 Then
                                oTankInfo.pipesCollection.Remove(oPipeInfo.ID)
                                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                                oPipeInfo = New MUSTER.Info.PipeInfo
                            End If
                        End If
                    Case "TANK"
                        If oTankInfo.pipesCollection.Count > 0 Then
                            For Each opipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                                oPipeInfo = opipeInfoLocal
                                Exit For
                            Next
                            If Not oPipeInfo Is Nothing Then
                                If oPipeInfo.TankID = id Then
                                    If oPipeInfo.IsDirty = False And oPipeInfo.IsAgedData = True Then
                                        Dim pipeID As String = oPipeInfo.ID
                                        oTankInfo.pipesCollection.Remove(oPipeInfo)
                                        'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                                        Add(pipeID, showDeleted)
                                    End If
                                    Exit Select
                                End If
                            End If
                        End If

                        ' From DB
                        'Dim colPipesLocal As New MUSTER.Info.PipesCollection
                        'If id = 0 Then id = -1
                        oTankInfo.pipesCollection = oPipeDB.DBGetByTankID(id, showDeleted)
                        If oTankInfo.pipesCollection.Count > 0 Then
                            'added by kiran
                            'raiseevent to tank to set compartmentCol to colComplocal
                            'RaiseEvent evtPipeColCompartment(id, colPipesLocal)
                            'end changes
                            Dim dtTempDate As Date = "12-22-1988"
                            For Each opipeInfoLocal As MUSTER.Info.PipeInfo In oTankInfo.pipesCollection.Values
                                oPipeInfo = opipeInfoLocal
                                If oPipeInfo.PipeStatusDesc = 426 Then
                                    oPipeInfo.POU = True
                                    If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                                        oPipeInfo.NonPre88 = True
                                    End If
                                End If
                            Next
                        End If
                End Select
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'oComments.Clear()
            'oComments.GetByModule("", nEntityTypeID, oPipeInfo.PipeID)
            Return oPipeInfo
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strUser As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal validateCapDates As Boolean = True, Optional ByVal strExcludeOCE As String = "", Optional ByVal bolSaveToInspectionMirror As Boolean = False, Optional ByVal passQuestion As Boolean = False) As Boolean
            Dim oldID As String
            Dim oldPipeID As Integer
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

                If Not bolValidated And Not oPipeInfo.Deleted Then
                    If Not Me.ValidateData(validateCapDates, strModule) Then
                        Return False
                    End If
                End If

                If Not (oPipeInfo.PipeID < 0 And oPipeInfo.Deleted) Then
                    oldID = oPipeInfo.ID
                    oldPipeID = oPipeInfo.PipeID
                    ' TODO - do i need to check for tos/tosi rules even when deleting the pipe
                    CheckTOSTOSIRules(oPipeInfo, moduleID, staffID, returnVal, strExcludeOCE, passquestion)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                    ' flags
                    If strModule.ToUpper = "REGISTRATION" Or strModule.ToUpper = "CLOSURE" Or strModule.ToUpper = "INSPECTION" Or strModule.ToUpper = "CAE" Then
                        ' if pipe status was changed to tos
                        Dim flags As New MUSTER.BusinessLogic.pFlag
                        If oPipeInfo.PipeStatusDesc = 425 Then ' TOS
                            If oPipeInfo.PipeStatusDescOriginal <> 425 Then ' NOT TOS
                                flags.RetrieveFlags(oPipeInfo.FacilityID, 6, , , , , "SYSTEM", "TOS Pipe")
                                If flags.FlagsCol.Count <= 0 Then
                                    flags.Add(New MUSTER.Info.FlagInfo(0, _
                                        oTankInfo.FacilityId, _
                                        6, _
                                        "TOS Pipes for Facility - " + oPipeInfo.FacilityID.ToString, _
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
                                flags.RetrieveFlags(oPipeInfo.FacilityID, 6, , , , , "SYSTEM", "TOS Pipe")
                                If flags.FlagsCol.Count <= 0 Then
                                    flags.Add(New MUSTER.Info.FlagInfo(0, _
                                        oTankInfo.FacilityId, _
                                        6, _
                                        "TOS Pipes for Facility - " + oPipeInfo.FacilityID.ToString, _
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
                            Dim bolNoTOSPipes As Boolean = True
                            Dim ds As DataSet = oPipeDB.DBGetDS("SELECT COUNT(PIPE_ID) AS TOSPIPES FROM tblREG_PIPE WHERE PIPE_ID IN (SELECT PIPE_ID FROM tblREG_COMPARTMENTS_PIPES WHERE ISNULL(DELETED,0) = 0 AND PIPE_STATUS_DESC = 425 AND TANK_ID = " + oPipeInfo.TankID.ToString + ")")
                            If ds.Tables(0).Rows(0)(0) > 0 Then
                                bolNoTOSPipes = False
                            Else
                                bolNoTOSPipes = True
                            End If
                            'For Each tnk As MUSTER.Info.TankInfo In FacilityInfo.TankCollection.Values
                            '    If tnk.TankStatus = 425 Then
                            '        bolNoTOSTanks = False
                            '        Exit For
                            '    End If
                            'Next
                            If bolNoTOSPipes Then
                                flags.RetrieveFlags(oPipeInfo.FacilityID, 6, , , , , "SYSTEM", "TOS Pipe")
                                For Each flagInfo As MUSTER.Info.FlagInfo In flags.FlagsCol.Values
                                    flagInfo.Deleted = True
                                Next
                                If flags.FlagsCol.Count > 0 Then
                                    flags.Flush()
                                End If
                            End If
                        End If
                    End If
                    ' save pipe
                    oPipeDB.Put(oPipeInfo, oTankInfo.FacCapStatus, moduleID, staffID, returnVal, strUser, bolSaveToInspectionMirror)

                    oPipeInfo.FacCapStatus = oTankInfo.FacCapStatus
                    'RaiseEvent evtFacCapStatus(oPipeInfo.FacilityID)
                    RaiseEvent evtPipeChanged(True)
                    If Not bolValidated Then
                        If oldPipeID <> oPipeInfo.PipeID And oPipeInfo.AttachedPipeID = 0 Then
                            oTankInfo.pipesCollection.ChangeKey(oldID, oPipeInfo.ID)
                            'RaiseEvent evtPipeChangeKey(oldID, oPipeInfo.ID)
                        End If
                        'oComments.Flush()
                    End If
                    oPipeInfo.Archive()
                    oPipeInfo.IsDirty = False
                End If
                If Not bolValidated Then
                    If oPipeInfo.Deleted Then
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oPipeInfo.ID Then
                            If strPrev = oPipeInfo.ID Then
                                RaiseEvent evtPipeErr("Pipe: " + oPipeInfo.Index.ToString + " of Tank: " + oPipeInfo.TankSiteID.ToString + " Compartment: " + oPipeInfo.CompartmentNumber.ToString + " deleted")
                                oTankInfo.pipesCollection.Remove(oPipeInfo)
                                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                                oPipeInfo = New MUSTER.Info.PipeInfo
                                RaiseEvent evtPipeChanged(False)
                            Else
                                RaiseEvent evtPipeErr("Pipe: " + oPipeInfo.Index.ToString + " of Tank: " + oPipeInfo.TankSiteID.ToString + " Compartment: " + oPipeInfo.CompartmentNumber.ToString + " deleted")
                                oTankInfo.pipesCollection.Remove(oPipeInfo)
                                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                                Dim compID As String = oTankInfo.pipesCollection.Item(strPrev).CompartmentID
                                oCompInfo = oTankInfo.CompartmentCollection.Item(compID)
                                oPipeInfo = Me.Retrieve(oTankInfo, strPrev, oCompInfo)
                            End If
                        Else
                            RaiseEvent evtPipeErr("Pipe: " + oPipeInfo.Index.ToString + " of Tank: " + oPipeInfo.TankSiteID.ToString + " Compartment: " + oPipeInfo.CompartmentNumber.ToString + " deleted")
                            oTankInfo.pipesCollection.Remove(oPipeInfo)
                            'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                            Dim compID As String = oTankInfo.pipesCollection.Item(strNext).CompartmentID
                            oCompInfo = oTankInfo.CompartmentCollection.Item(compID)
                            oPipeInfo = Me.Retrieve(oTankInfo, strNext, oCompInfo)
                        End If
                    End If
                End If
                'RaiseEvent evtSyncPipeInCol(oPipeInfo)
                RaiseEvent evtPipeChanged(Me.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub SaveCompartmentsPipe(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oPipeDB.PutCompartmentsPipe(oPipeInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                oPipeInfo.Archive()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(ByVal validateCapDates As Boolean, Optional ByVal [module] As String = "Registration", Optional ByRef strError As String = "", Optional ByVal returnString As Boolean = False) As Boolean
            'PipeModDesc - Cathodically Protected = 260
            'PipeStatusDesc - Currently in Use (CIU) = 424
            'PipeStatusDesc - Temporarily Out of Service Indefinitely (TOSI) = 429
            'Termination Type at Tank - Coated/Wrapped Cathodically Protected = 610
            'Termination Type at Disp - Coated/Wrapped Cathodically Protected = 61
            'PipeType - Pressurized = 266
            'PipeLD - Continuous Interstitial Monitoring = 243
            'PipeLD - Line Tightness Testing = 245
            Try
                Dim errStr As String = ""
                Dim msgStr As String = ""
                Dim dateLocal As Date = CDate("01/01/0001")
                Dim validateSuccess As Boolean = True
                Dim dtTempDate As Date = "12-22-1988"
                Select Case [module].Trim.ToUpper
                    Case "REGISTRATION", "INSPECTION"
                        If oPipeInfo.PipeStatusDesc = 426 Then 'POU
                            If Not (oPipeInfo.POU And oPipeInfo.NonPre88) Then
                                If Date.Compare(oPipeInfo.DateLastUsed, dateLocal) = 0 Then
                                    errStr += "Date Last Used cannot be empty" + vbCrLf
                                    validateSuccess = False
                                Else
                                    If Not oPipeInfo.NonPre88 And Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                                        errStr += "Date Last Used must be < '12-22-1988' for POU Pipes" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If
                            End If
                            'If Date.Compare(oPipeInfo.DateLastUsed, dateLocal) = 0 Then
                            '    errStr += "Date Last Used cannot be empty" + vbCrLf
                            '    validateSuccess = False
                            'Else
                            '    If Not oPipeInfo.NonPre88 And Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                            '        errStr += "Date Last Used must be < '12-22-88' for POU Pipes" + vbCrLf
                            '        validateSuccess = False
                            '    End If
                            'End If
                        ElseIf oPipeInfo.PipeStatusDesc = 425 Then 'TOS
                            If Date.Compare(oPipeInfo.DateLastUsed, dateLocal) = 0 Then
                                errStr += "Date Last Used cannot be empty" + vbCrLf
                                validateSuccess = False
                            End If
                        ElseIf oPipeInfo.PipeStatusDesc = 429 Then  ' TOSI
                            If oPipeInfo.PipeTypeDesc = 266 Then ' pressurized
                                ' release detection 1 and 2 are required
                                ' 1 - pipe ld, 2 - alld
                                ' #632
                                ' if release detection 1 is deferred, release detection 2 is not required
                                If oPipeInfo.PipeLD = 0 Then
                                    errStr += "Release Detection 1 is required" + vbCrLf
                                    validateSuccess = False
                                End If
                                If oPipeInfo.PipeLD <> 248 Then
                                    If oPipeInfo.ALLDType = 0 Then
                                        errStr += "Release Detection 2 is required" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If
                            ElseIf oPipeInfo.PipeTypeDesc = 268 Then ' u.s. suction
                                If oPipeInfo.PipeLD = 0 Then
                                    errStr += "Release Detection 1 is required" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If
                            If oPipeInfo.PipeMatDesc <> 255 And oPipeInfo.PipeMatDesc <> 257 And oPipeInfo.PipeMatDesc <> 0 Then
                                If oPipeInfo.PipeModDesc = 0 Then
                                    errStr += "Pipe Secondary Option cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If

                            If Date.Compare(oPipeInfo.DateLastUsed, CDate("01/01/0001")) = 0 Then
                                errStr += "Date Last Used cannot be empty" + vbCrLf
                                validateSuccess = False
                            End If

                            ValidateDates(errStr, validateSuccess, msgStr, validateCapDates)

                        ElseIf oPipeInfo.PipeStatusDesc = 424 Then 'CIU
                            If oPipeInfo.PipeTypeDesc = 266 Then ' pressurized
                                ' release detection 1 and 2 are required
                                ' 1 - pipe ld, 2 - alld
                                ' #632
                                ' if release detection 1 is deferred, release detection 2 is not required
                                If oPipeInfo.PipeLD = 0 Then
                                    errStr += "Release Detection 1 is required" + vbCrLf
                                    validateSuccess = False
                                End If
                                If oPipeInfo.PipeLD <> 248 Then
                                    If oPipeInfo.ALLDType = 0 Then
                                        errStr += "Release Detection 2 is required" + vbCrLf
                                        validateSuccess = False
                                    End If
                                End If
                            ElseIf oPipeInfo.PipeTypeDesc = 268 Then ' u.s. suction
                                If oPipeInfo.PipeLD = 0 Then
                                    errStr += "Release Detection 1 is required" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If
                            If oPipeInfo.PipeMatDesc <> 255 And oPipeInfo.PipeMatDesc <> 257 And oPipeInfo.PipeMatDesc <> 0 Then
                                If oPipeInfo.PipeModDesc = 0 Then
                                    errStr += "Pipe Secondary Option cannot be empty" + vbCrLf
                                    validateSuccess = False
                                End If
                            End If

                            ValidateDates(errStr, validateSuccess, msgStr, validateCapDates)

                        End If
                End Select
                ' if any validations failed
                If errStr.Length > 0 Or Not validateSuccess Then
                    If msgStr.Length > 0 Then errStr += "Optional:" + vbCrLf + msgStr
                    If returnString Then
                        strError = errStr
                    Else
                        RaiseEvent evtPipeErr(errStr)
                    End If
                ElseIf msgStr.Length > 0 Then
                    If returnString Then
                        strError = "Optional:" + vbCrLf + msgStr
                    Else
                        MsgBox("Optional:" + vbCrLf + msgStr, MsgBoxStyle.OKOnly, "Pipe Validation")
                    End If
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub ValidateDates(ByRef errStr As String, ByRef validateSuccess As Boolean, ByRef msgStr As String, ByVal validateCapDates As Boolean)
            Dim dateLocal As Date = CDate("01/01/0001")
            Dim dttemp, dtValidDate As Date
            Dim dtTodayPlus90Days As Date = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 4, CDate(Today.Month.ToString + "/1/" + Today.Year.ToString)))
            Try
                If oPipeInfo.PipeStatusDesc = 429 Or oPipeInfo.PipeStatusDesc = 424 Then  ' TOSI / CIU

                    ' Pipe Placed in Service On >= Pipe Installed On
                    If Date.Compare(oPipeInfo.PlacedInServiceDate, dateLocal) <> 0 And Date.Compare(oPipeInfo.PipeInstallDate, dateLocal) <> 0 Then
                        If Date.Compare(oPipeInfo.PlacedInServiceDate, oPipeInfo.PipeInstallDate) < 0 Then
                            errStr += "Pipe Placed in Service On must be greater or equal to Pipe Installed On" + vbCrLf
                            validateSuccess = False
                        End If
                    End If

                    ' PipeCPTest
                    ' if Pipe Mod Desc like Cathodically Protected
                    If validateCapDates Then
                        If oPipeInfo.PipeModDesc = 260 Or oPipeInfo.PipeModDesc = 263 Then
                            If Date.Compare(oPipeInfo.PipeCPInstalledDate, dateLocal) = 0 Then
                                msgStr += "Provide Pipe CP Install Date" + vbCrLf
                            End If
                            dttemp = oPipeInfo.PipeCPTest
                            dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                            dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                msgStr += "Pipe CP Last Tested must be greater than or equal to " + dtValidDate.ToShortDateString + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                            End If
                        End If
                    End If

                    ' Date Pipe CP Last Tested >= Date CP Installed
                    If Date.Compare(oPipeInfo.PipeCPTest, dateLocal) <> 0 And Date.Compare(oPipeInfo.PipeCPInstalledDate, dateLocal) <> 0 Then
                        If Date.Compare(oPipeInfo.PipeCPTest, oPipeInfo.PipeCPInstalledDate) < 0 Then
                            errStr += "Date Pipe CP Last Tested must be greater or equal to Date CP Installed" + vbCrLf
                            validateSuccess = False
                        End If
                    End If

                    ' Term type at tank or disp = coated wrapped/cathodically protected
                    If oPipeInfo.TermTypeDisp = 611 Or oPipeInfo.TermTypeTank = 610 Then
                        If validateCapDates Then
                            If Date.Compare(oPipeInfo.TermCPInstalledDate, dateLocal) = 0 Then
                                msgStr += "Provide Term CP Installed Date" + vbCrLf
                            End If
                            dttemp = oPipeInfo.TermCPLastTested
                            dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                            dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                msgStr += "Termination CP Last Tested must be greater than or equal to " + dtValidDate.ToShortDateString + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                            End If
                        End If
                    Else
                        oPipeInfo.TermCPInstalledDate = CDate("01/01/0001")
                    End If

                    ' Termination CP Last Tested On >= Termination CP Installed On
                    If Date.Compare(oPipeInfo.TermCPLastTested, dateLocal) <> 0 And Date.Compare(oPipeInfo.TermCPInstalledDate, dateLocal) <> 0 Then
                        If Date.Compare(oPipeInfo.TermCPLastTested, oPipeInfo.TermCPInstalledDate) < 0 Then
                            errStr += "Termination CP Last Tested On must be greater or equal to Termination CP Installed On" + vbCrLf
                            validateSuccess = False
                        End If
                    End If

                    '****************************************************************************************'
                    '*                                                                                      *'
                    '****************************************************************************************'
                    If oPipeInfo.PipeStatusDesc = 424 Then ' CIU

                        ' PipeLD = Line Tightness Testing and CAP
                        If oPipeInfo.PipeLD = 245 Then
                            If validateCapDates Then
                                If Date.Compare(oPipeInfo.LTTDate, dateLocal) = 0 Then
                                    msgStr += "Provide Last Pipe Tightness Test Date" + vbCrLf
                                Else
                                    dttemp = oPipeInfo.LTTDate
                                    If oPipeInfo.PipeTypeDesc = 268 Then
                                        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                                        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                    ElseIf oPipeInfo.PipeTypeDesc = 266 Then
                                        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
                                        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                    End If
                                    If oPipeInfo.PipeTypeDesc = 268 Or oPipeInfo.PipeTypeDesc = 266 Then
                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                            msgStr += "Last Pipe Tightness Tested must be greater than or equal to " + dtValidDate.ToShortDateString + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            oPipeInfo.LTTDate = CDate("01/01/0001")
                        End If

                        ' ALLDType = Mechanical
                        If oPipeInfo.ALLDType = 496 Then
                            If validateCapDates Then
                                If Date.Compare(oPipeInfo.ALLDTestDate, CDate("01/01/0001")) = 0 Then
                                    msgStr += "Provide Automatic Line Leak Detector Test" + vbCrLf
                                Else
                                    dttemp = oPipeInfo.ALLDTestDate
                                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                    dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
                                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                                        msgStr += "Automatic Line Leak Detector Test must be greater than or equal to " + dtValidDate.ToShortDateString + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                                    End If
                                End If
                            End If
                        End If

                    End If

                End If

                'If Date.Compare(oPipeInfo.DateLastUsed, CDate("01/01/0001")) = 0 Then
                '    errStr += "Date Last Used cannot be empty" + vbCrLf
                '    validateSuccess = False
                'End If

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub CopyPipeProfile(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal Comp_ID As Integer = -2)
            Dim oldID As String
            Dim newPipeID As Integer
            Dim oldPipeInfo As New Collection
            Try
                If MsgBox("Do you want to continue Pipe Profile Copy?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    oldPipeInfo.Add(oPipeInfo.TankID)
                    oldPipeInfo.Add(oPipeInfo.TankSiteID)
                    oldPipeInfo.Add(oPipeInfo.FacilityID)
                    oldPipeInfo.Add(oPipeInfo.CompartmentID)
                    oldPipeInfo.Add(oPipeInfo.CompartmentNumber)
                    oldPipeInfo.Add(oPipeInfo.CompartmentSubstance)
                    oldPipeInfo.Add(oPipeInfo.CompartmentCERCLA)
                    oldPipeInfo.Add(oPipeInfo.CompartmentFuelType)
                    newPipeID = oPipeDB.CopyPipeProfile(oPipeInfo.PipeID, moduleID, staffID, returnVal, UserID, Comp_ID)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    If newPipeID <> oPipeInfo.PipeID Then
                        Retrieve(oTankInfo, newPipeID, oCompInfo, , "PIPE")
                        oPipeInfo.TankSiteID = oldPipeInfo.Item(2)
                        oPipeInfo.FacilityID = oldPipeInfo.Item(3)
                        oPipeInfo.CompartmentID = oldPipeInfo.Item(4)
                        oPipeInfo.CompartmentSubstance = oldPipeInfo.Item(6)
                        oPipeInfo.CompartmentCERCLA = oldPipeInfo.Item(7)
                        oPipeInfo.CompartmentFuelType = oldPipeInfo.Item(8)
                        'ChangePipeTankCompartmentNumberKey(oldPipeInfo.Item(1), oldPipeInfo.Item(5), , oPipeInfo)
                        oTankInfo.pipesCollection.ChangeKey(CType(oldPipeInfo.Item(1), String) + "|0|" + oPipeInfo.PipeID.ToString, oPipeInfo.ID)
                        'RaiseEvent evtPipeChangeKey(CType(oldPipeInfo.Item(1), String) + "|0|" + oPipeInfo.PipeID.ToString, oPipeInfo.ID)
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub DeletePipe(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try
                If oTankInfo.TankId > 0 Then
                    Dim ds As DataSet
                    ds = oPipeDB.DBGetDS("EXEC spCheckDependancy NULL,NULL,NULL,0," + oPipeInfo.PipeID.ToString)
                    If ds.Tables(0).Rows(0)(0) Then
                        RaiseEvent evtPipeErr(IIf(ds.Tables(0).Rows(0)("MSG") Is DBNull.Value, "Pipe has dependants", ds.Tables(0).Rows(0)("MSG")))
                        Exit Sub
                    End If
                End If

                oPipeDB.DeletePipe(oPipeInfo.PipeID, moduleID, staffID, returnVal, UserID)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                Dim strNext As String = Me.GetNext()
                Dim strPrev As String = Me.GetPrevious()
                If strNext = oPipeInfo.ID Then
                    If strPrev = oPipeInfo.ID Then
                        RaiseEvent evtPipeErr("Pipe: " + oPipeInfo.Index.ToString + " of Tank: " + oPipeInfo.TankSiteID.ToString + " Compartment: " + oPipeInfo.CompartmentNumber.ToString + " deleted")
                        oTankInfo.pipesCollection.Remove(oPipeInfo)
                        'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                        oPipeInfo = New MUSTER.Info.PipeInfo
                        RaiseEvent evtPipeChanged(False)
                    Else
                        RaiseEvent evtPipeErr("Pipe: " + oPipeInfo.Index.ToString + " of Tank: " + oPipeInfo.TankSiteID.ToString + " Compartment: " + oPipeInfo.CompartmentNumber.ToString + " deleted")
                        oTankInfo.pipesCollection.Remove(oPipeInfo)
                        'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                        Dim compID As String = oTankInfo.pipesCollection.Item(strPrev).CompartmentID
                        oCompInfo = oTankInfo.CompartmentCollection.Item(compID)
                        oPipeInfo = Me.Retrieve(oTankInfo, strPrev, oCompInfo)
                    End If
                Else
                    RaiseEvent evtPipeErr("Pipe: " + oPipeInfo.Index.ToString + " of Tank: " + oPipeInfo.TankSiteID.ToString + " Compartment: " + oPipeInfo.CompartmentNumber.ToString + " deleted")
                    oTankInfo.pipesCollection.Remove(oPipeInfo)
                    'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                    Dim compID As String = oTankInfo.pipesCollection.Item(strPrev).CompartmentID
                    oCompInfo = oTankInfo.CompartmentCollection.Item(compID)
                    oPipeInfo = Me.Retrieve(oTankInfo, strNext, oCompInfo)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub CopyPipeInfo(ByVal strID As String)
            Try
                Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
                'Dim colPipesContained As MUSTER.Info.PipesCollection
                'RaiseEvent evtCompInfoPipeCol(colPipesContained)
                oPipeInfoLocal = oTankInfo.pipesCollection.Item(strID)
                If Not (oPipeInfoLocal Is Nothing) Then
                    oPipeInfo.ALLDTest = oPipeInfoLocal.ALLDTest
                    oPipeInfo.ALLDTestDate = oPipeInfoLocal.ALLDTestDate
                    oPipeInfo.CASNumber = oPipeInfoLocal.CASNumber
                    oPipeInfo.ClosureStatusDesc = oPipeInfoLocal.ClosureStatusDesc
                    oPipeInfo.ClosureType = oPipeInfoLocal.ClosureType
                    oPipeInfo.CompartmentCERCLA = oPipeInfoLocal.CompartmentCERCLA
                    oPipeInfo.CompartmentFuelType = oPipeInfoLocal.CompartmentFuelType
                    oPipeInfo.CompartmentID = oPipeInfo.TankID.ToString + "|" + oPipeInfo.CompartmentNumber.ToString
                    oPipeInfo.CompartmentSubstance = oPipeInfoLocal.CompartmentSubstance
                    oPipeInfo.CompPrimary = oPipeInfoLocal.CompPrimary
                    oPipeInfo.CompSecondary = oPipeInfoLocal.CompSecondary
                    oPipeInfo.ContainSumpDisp = oPipeInfoLocal.ContainSumpDisp
                    oPipeInfo.ContainSumpTank = oPipeInfoLocal.ContainSumpTank
                    oPipeInfo.ContractorID = oPipeInfoLocal.ContractorID
                    oPipeInfo.DateClosed = oPipeInfoLocal.DateClosed
                    oPipeInfo.DateClosureRecd = oPipeInfoLocal.DateClosureRecd
                    oPipeInfo.DateLastUsed = oPipeInfoLocal.DateLastUsed
                    oPipeInfo.DateRecd = oPipeInfoLocal.DateRecd
                    oPipeInfo.DateSigned = oPipeInfoLocal.DateSigned
                    oPipeInfo.Deleted = oPipeInfoLocal.Deleted
                    oPipeInfo.FacilityID = oPipeInfoLocal.FacilityID
                    oPipeInfo.Index = oPipeInfoLocal.Index
                    oPipeInfo.InertMaterial = oPipeInfoLocal.InertMaterial
                    oPipeInfo.LCPInstallDate = oPipeInfoLocal.LCPInstallDate
                    oPipeInfo.LicenseeID = oPipeInfoLocal.LicenseeID
                    oPipeInfo.LTTDate = oPipeInfoLocal.LTTDate
                    oPipeInfo.PipeCPInstalledDate = oPipeInfoLocal.PipeCPInstalledDate
                    oPipeInfo.PipeCPTest = oPipeInfoLocal.PipeCPTest
                    oPipeInfo.DateShearTest = oPipeInfoLocal.DateShearTest
                    oPipeInfo.DatePipeSecInsp = oPipeInfoLocal.DatePipeSecInsp
                    oPipeInfo.DatePipeElecInsp = oPipeInfoLocal.DatePipeElecInsp
                    oPipeInfo.PipeCPType = oPipeInfoLocal.PipeCPType
                    oPipeInfo.PipeInstallationPlannedFor = oPipeInfoLocal.PipeInstallationPlannedFor
                    oPipeInfo.PipeInstallDate = oPipeInfoLocal.PipeInstallDate
                    oPipeInfo.PipeLD = oPipeInfoLocal.PipeLD
                    oPipeInfo.PipeManufacturer = oPipeInfoLocal.PipeManufacturer
                    oPipeInfo.PipeMatDesc = oPipeInfoLocal.PipeMatDesc
                    oPipeInfo.PipeModDesc = oPipeInfoLocal.PipeModDesc
                    oPipeInfo.PipeOtherMaterial = oPipeInfoLocal.PipeOtherMaterial
                    oPipeInfo.PipeStatusDesc = oPipeInfoLocal.PipeStatusDesc
                    oPipeInfo.PipeTypeDesc = oPipeInfoLocal.PipeTypeDesc
                    oPipeInfo.PipingComments = oPipeInfoLocal.PipingComments
                    oPipeInfo.PlacedInServiceDate = oPipeInfoLocal.PlacedInServiceDate
                    oPipeInfo.SubstanceComments = oPipeInfoLocal.SubstanceComments
                    oPipeInfo.SubstanceDesc = oPipeInfoLocal.SubstanceDesc
                    oPipeInfo.TankSiteID = oPipeInfoLocal.TankSiteID
                    oPipeInfo.TermCPInstalledDate = oPipeInfoLocal.TermCPInstalledDate
                    oPipeInfo.TermCPLastTested = oPipeInfoLocal.TermCPLastTested
                    oPipeInfo.TermCPTypeDisp = oPipeInfoLocal.TermCPTypeDisp
                    oPipeInfo.TermCPTypeTank = oPipeInfoLocal.TermCPTypeTank
                    oPipeInfo.TermTypeDisp = oPipeInfoLocal.TermTypeDisp
                    oPipeInfo.TermTypeTank = oPipeInfoLocal.TermTypeTank
                    oPipeInfo.Archive()
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub CheckTOSTOSIRules(ByRef pipeInfo As MUSTER.Info.PipeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal strExcludeOCE As String, ByVal passQuestion As Boolean)
            Dim citation As New MUSTER.BusinessLogic.pInspectionCitation
            Dim citationExists As Boolean = False
            Try
                ' check for CIU <=> TOS/TOSI rules for existing pipe (not new pipe)
                If pipeInfo.PipeID > 0 Then
                    ' CIU TO TOS/TOSI
                    If pipeInfo.PipeStatusDescOriginal = 424 And _
                        (pipeInfo.PipeStatusDesc = 429 Or pipeInfo.PipeStatusDesc = 425) Then
                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), pipeInfo.FacilityID, False, 10, 1, 1, 0, strExcludeOCE)
                        ' if there is a citation for corrosion protection not maintained
                        If citationExists Then
                            ' notify user that selected status was changed to TOS if selected status was not TOS
                            If pipeInfo.PipeStatusDesc <> 425 AndAlso (Not passQuestion AndAlso MsgBox("Pipe: " + pipeInfo.Index.ToString + "s business rule states that the status should change to TOS. Do you wish to change to TOSI?", MsgBoxStyle.YesNo) = MsgBoxResult.No) Then

                                ' pipe status is TOS
                                pipeInfo.PipeStatusDesc = 425

                            End If
                        ElseIf pipeInfo.PipeStatusDesc = 425 Then ' TOS
                            ProcessSubstandard(pipeInfo, citationExists, passQuestion, "")
                        End If

                        ' TOSI TO CIU/TOS
                    ElseIf pipeInfo.PipeStatusDescOriginal = 429 And _
                        (pipeInfo.PipeStatusDesc = 424 Or pipeInfo.PipeStatusDesc = 425) Then
                        ' if status is tos, process substandard
                        ' if status is ciu, allow change - no need to check for any conditions
                        If pipeInfo.PipeStatusDesc = 425 Then ' TOS
                            citationExists = citation.CheckCitationExists(CDate("01/01/0001"), pipeInfo.FacilityID, False, 10, 1, 1, 0, strExcludeOCE)
                            ProcessSubstandard(pipeInfo, citationExists, passQuestion, "")
                            ' no need to create fce. need to check if citation exists
                            'If pipeInfo.PipeStatusDesc = 425 Then ' TOS
                            '    ' create a citation for corrosion protection not maintained
                            '    ' create a manual fce with citation 10
                            '    CreateCPNotMaintainedCitation(pipeInfo, moduleID, staffID, returnVal)
                            '    If Not returnVal = String.Empty Then
                            '        Exit Sub
                            '    End If
                            'End If
                        End If

                        ' TOS TO CIU/TOSI
                    ElseIf pipeInfo.PipeStatusDescOriginal = 425 And _
                            (pipeInfo.PipeStatusDesc = 424 Or pipeInfo.PipeStatusDesc = 429) Then
                        If pipeInfo.PipeStatusDesc = 429 Then ' TOSI
                            citationExists = citation.CheckCitationExists(CDate("01/01/0001"), pipeInfo.FacilityID, False, 10, 1, 1, 0, strExcludeOCE)
                            ' if status is tosi, process substandard
                            ProcessSubstandard(pipeInfo, citationExists, passQuestion, "")
                            'DeleteCPNotMaintainedCitation(pipeInfo, moduleID, staffID, returnVal)
                            'If Not returnVal = String.Empty Then
                            '    Exit Sub
                            'End If
                        ElseIf pipeInfo.PipeStatusDesc = 424 Then
                            citationExists = citation.CheckCitationExists(CDate("01/01/0001"), pipeInfo.FacilityID, False, 10, 1, 1, 0, strExcludeOCE)
                            ' if status is ciu
                            ' if there is a citation for corrosion protection not maintained
                            ' pipe status is tos
                            If citationExists AndAlso (Not passQuestion AndAlso MsgBox("Pipe: " + pipeInfo.Index.ToString + "s business rule states that the status should change to TOS. Do you wish to change to TOSI?", MsgBoxStyle.YesNo) = MsgBoxResult.No) Then
                                ' notify user that selected status was changed to TOS
                                pipeInfo.PipeStatusDesc = 425
                                'DeleteCPNotMaintainedCitation(pipeInfo, moduleID, staffID, returnVal)
                                'If Not returnVal = String.Empty Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    ElseIf pipeInfo.PipeStatusDescOriginal = pipeInfo.PipeStatusDesc And _
                            (pipeInfo.PipeStatusDesc = 425 Or pipeInfo.PipeStatusDesc = 429) Then ' TOS / TOSI

                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), TankInfo.FacilityId, False, 10, 1, 1, 0, strExcludeOCE)

                        If (pipeInfo.PipeStatusDesc = 429 And citationExists) AndAlso (Not passQuestion AndAlso MsgBox("Pipe: " + pipeInfo.Index.ToString + "s business rule states that the status should change to TOS. Do wish to change to TOSI?", MsgBoxStyle.YesNo) = MsgBoxResult.No) Then  ' TOSI
                            ' change to tos
                            pipeInfo.PipeStatusDesc = 425
                            'CreateCPNotMaintainedCitation(pipeInfo, moduleID, staffID, returnVal)
                            'If Not returnVal = String.Empty Then
                            '    Exit Sub
                            'End If
                        ElseIf pipeInfo.PipeStatusDesc = 425 And Not citationExists Then ' TOS
                            ProcessSubstandard(pipeInfo, citationExists, passQuestion, "")
                            'ElseIf pipeInfo.PipeStatusDesc = 425 And citationExists Then ' TOS
                            '    DeleteCPNotMaintainedCitation(pipeInfo, moduleID, staffID, returnVal)
                            '    If Not returnVal = String.Empty Then
                            '        Exit Sub
                            '    End If
                        End If

                    End If
                Else


                    ' if it is a new pipe
                    ' if pipe status is TOS, process substandard
                    If pipeInfo.PipeStatusDesc = 425 Then
                        citationExists = citation.CheckCitationExists(CDate("01/01/0001"), pipeInfo.FacilityID, False, 10, 1, 1, 0, strExcludeOCE)
                        ProcessSubstandard(pipeInfo, citationExists, passQuestion, "")
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Private Sub ProcessSubstandard(ByRef pipeInfo As MUSTER.Info.PipeInfo, ByVal citationExists As Boolean, ByVal passQuestion As Boolean, Optional ByVal msg2 As String = "")
            Dim p As BusinessLogic.pProperty
            Dim msg As String = String.Empty

            Try
                p = New BusinessLogic.pProperty
                ' If ((PipeMaterial = "Bare/Galvanized Steel" or "Copper" or "Coated/Wrapped Steel" or "Other") and
                '     (PipeSecondaryOption is not "Double-Walled" or "Cathodically Protected"))
                '                                    OR
                '    (there is a citation for corrosion protection not maintained and
                '    PipeCPType = "Impressed Current")
                '   PipeStatus = TOS
                ' Else
                '   PipeStatus = TOSI
                If ((pipeInfo.PipeMatDesc = 250 Or pipeInfo.PipeMatDesc = 253 Or pipeInfo.PipeMatDesc = 251 Or pipeInfo.PipeMatDesc = 256) And _
                    (pipeInfo.PipeModDesc <> 261 And pipeInfo.PipeModDesc <> 260)) Or _
                    (citationExists And pipeInfo.PipeCPType = 478) Then


                    ' notify user that selected status was changed to TOS if selected status was not TOS
                    If pipeInfo.PipeStatusDesc <> 425 AndAlso (Not passQuestion AndAlso MsgBox("Pipe: " + pipeInfo.Index.ToString + String.Format("s business rule states that the status should change to TOS{0}. Do you wish to change to TOSI?", msg), MsgBoxStyle.YesNo) = MsgBoxResult.No) Then
                        ' pipe status is TOS
                        pipeInfo.PipeStatusDesc = 425

                    End If
                Else

                    ' notify user that selected status was changed to TOSI if selected status was not TOSI
                    If pipeInfo.PipeStatusDesc <> 429 AndAlso (Not passQuestion AndAlso MsgBox("Pipe: " + pipeInfo.Index.ToString + String.Format("s business rule states that the status should change to TOSI{0}. Do you wish to change to TOS?", msg), MsgBoxStyle.YesNo) = MsgBoxResult.No) Then

                        ' pipe status is TOSI
                        pipeInfo.PipeStatusDesc = 429

                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                p = Nothing
            End Try
        End Sub
        'Private Sub CreateCPNotMaintainedCitation(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
        '    Dim oInspection As New MUSTER.BusinessLogic.pInspection
        '    Dim oFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
        '    Dim oInspectionCitation As New MUSTER.BusinessLogic.pInspectionCitation
        '    Dim nOwnerID As Integer = 0
        '    Try
        '        ' check rights
        '        ' no need to check rights as the fce needs to be created irrespective of user rights
        '        'CheckRightsToSaveCPNotMaintainedCitation(moduleID, staffID, returnVal)
        '        'If Not returnVal = String.Empty Then
        '        '    Exit Sub
        '        'End If
        '        ' create inspection
        '        oInspection.Retrieve(0)
        '        oInspection.FacilityID = pipeInfo.FacilityID
        '        Dim ds As DataSet = oPipeDB.DBGetDS("SELECT OWNER_ID FROM TBLREG_FACILITY WHERE FACILITY_ID = " + pipeInfo.FacilityID.ToString)
        '        If ds.Tables(0).Rows(0)("OWNER_ID") Is DBNull.Value Then
        '            returnVal = "Invalid Owner ID for facility " + pipeInfo.FacilityID.ToString
        '            Exit Sub
        '        Else
        '            nOwnerID = ds.Tables(0).Rows(0)("OWNER_ID")
        '        End If
        '        oInspection.OwnerID = nOwnerID
        '        oInspection.InspectionType = 1132
        '        oInspection.LetterGenerated = False
        '        oInspection.CreatedBy = IIf(pipeInfo.ModifiedBy = String.Empty, pipeInfo.CreatedBy, pipeInfo.ModifiedBy)
        '        oInspection.Save(moduleID, staffID, returnVal, , , , True)
        '        If Not returnVal = String.Empty Then
        '            Exit Sub
        '        End If
        '        ' create manual fce
        '        oFCE.Retrieve(0)
        '        oFCE.InspectionID = oInspection.ID
        '        oFCE.OwnerID = oInspection.OwnerID
        '        oFCE.FacilityID = oInspection.FacilityID
        '        oFCE.Source = "ADMIN"
        '        oFCE.FCEDate = Today.Date
        '        oFCE.CreatedBy = oInspection.CreatedBy
        '        oFCE.Save(moduleID, staffID, returnVal, , , True)
        '        If Not returnVal = String.Empty Then
        '            Exit Sub
        '        End If
        '        ' create citation
        '        oInspectionCitation.Retrieve(oInspection.InspectionInfo, 0)
        '        oInspectionCitation.FacilityID = oFCE.FacilityID
        '        oInspectionCitation.FCEID = oFCE.ID
        '        oInspectionCitation.InspectionID = oInspection.ID
        '        oInspectionCitation.CitationID = 10
        '        oInspectionCitation.QuestionID = oInspection.CheckListMaster.RetrieveByCheckListItemNum("99998").ID
        '        oInspectionCitation.CreatedBy = oFCE.CreatedBy
        '        oInspectionCitation.Save(moduleID, staffID, returnVal, , , True)
        '        If Not returnVal = String.Empty Then
        '            Exit Sub
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub DeleteCPNotMaintainedCitation(ByVal pipeInfo As MUSTER.Info.PipeInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
        '    Dim nFCEID, nInspectionID, nCitationID As Integer
        '    Try
        '        ' check rights
        '        ' no need to check rights as the fce needs to be created irrespective of user rights
        '        'CheckRightsToSaveCPNotMaintainedCitation(moduleID, staffID, returnVal)
        '        'If Not returnVal = String.Empty Then
        '        '    Exit Sub
        '        'End If
        '        ' delete only if the citation is linked to manual fce
        '        Dim strSQL As String = "SELECT FCE.FCE_ID, FCE.INSPECTION_ID, C.INS_CIT_ID " + _
        '                                "FROM tblCAE_FACILITIY_COMPLIANCE_EVENT FCE " + _
        '                                "INNER JOIN tblINS_INSPECTION_CITATION C ON C.INSPECTION_ID = FCE.INSPECTION_ID AND C.FCE_ID = FCE.FCE_ID " + _
        '                                "WHERE FCE.FACILITY_ID = " + pipeInfo.FacilityID.ToString + _
        '                                "AND FCE.SOURCE = 'ADMIN' " + _
        '                                "AND FCE.DELETED = 0 " + _
        '                                "AND FCE.OCE_GENERATED = 0 " + _
        '                                "AND C.RESCINDED = 0 " + _
        '                                "AND C.NFA_DATE IS NULL"
        '        Dim ds As DataSet = oPipeDB.DBGetDS(strSQL)
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
        '        If Not oPipeDB.SqlHelperProperty.HasWriteAccess(moduleID, staffID, oPipeDB.SqlHelperProperty.EntityTypes.Inspection) Then
        '            returnVal = "Inspection, "
        '        End If
        '        If Not oPipeDB.SqlHelperProperty.HasWriteAccess(moduleID, staffID, oPipeDB.SqlHelperProperty.EntityTypes.CAEFacilityCompliantEvent) Then
        '            returnVal += "FCE, "
        '        End If
        '        If Not oPipeDB.SqlHelperProperty.HasWriteAccess(moduleID, staffID, oPipeDB.SqlHelperProperty.EntityTypes.Citation) Then
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
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.PipesCollection
            Try
                oTankInfo.pipesCollection.Clear()
                oTankInfo.pipesCollection = oPipeDB.GetAllInfo
                Return oTankInfo.pipesCollection
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Add(ByVal ID As String, Optional ByVal ShowDeleted As Boolean = False)
            Try
                Dim dtTempDate As Date = "12-22-1988"
                Dim strID() As String = ID.Split("|")
                oPipeInfo = oPipeDB.DBGetByID(strID(2), ShowDeleted)
                If oPipeInfo.PipeID = 0 Then
                    oPipeInfo.PipeID = nID
                    oPipeInfo.FacilityID = oTankInfo.FacilityId
                    nID -= 1

                    oPipeInfo.DateLastUsed = oTankInfo.DateLastUsed
                    oPipeInfo.PlacedInServiceDate = oTankInfo.PlacedInServiceDate
                End If
                If oPipeInfo.PipeStatusDesc = 426 Then
                    oPipeInfo.POU = True
                    If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                        oPipeInfo.NonPre88 = True
                    End If
                End If

                oPipeInfo.TankID = CType(strID(0), Integer)
                oPipeInfo.CompartmentNumber = CType(strID(1), Integer)
                oPipeInfo.TankSiteID = oTankInfo.TankIndex
                If Not oCompInfo Is Nothing Then
                    oPipeInfo.CompartmentCERCLA = oCompInfo.CCERCLA
                    oPipeInfo.CompartmentFuelType = oCompInfo.FuelTypeId
                    oPipeInfo.CompartmentSubstance = oCompInfo.Substance
                End If
                If Date.Compare(oPipeInfo.DateSigned, CDate("01/01/0001")) = 0 Then
                    oPipeInfo.DateSigned = oTankInfo.DateSigned
                End If
                If oPipeInfo.LicenseeID = 0 Then
                    oPipeInfo.LicenseeID = oTankInfo.LicenseeID
                End If
                If oPipeInfo.ContractorID = 0 Then
                    oPipeInfo.ContractorID = oTankInfo.ContractorID
                End If
                oPipeInfo.FacCapStatus = oTankInfo.FacCapStatus

                oTankInfo.pipesCollection.Add(oPipeInfo)
                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "ADD")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer, Optional ByVal ShowDeleted As Boolean = False)
            Try
                Dim dtTempDate As Date = "12-22-1988"
                oPipeInfo = oPipeDB.DBGetByID(ID, ShowDeleted)
                If oPipeInfo.PipeID = 0 Then
                    oPipeInfo.PipeID = nID
                    oPipeInfo.FacilityID = oTankInfo.FacilityId
                    nID -= 1

                    oPipeInfo.DateLastUsed = oTankInfo.DateLastUsed
                    oPipeInfo.PlacedInServiceDate = oTankInfo.PlacedInServiceDate
                End If
                If oPipeInfo.PipeStatusDesc = 426 Then
                    oPipeInfo.POU = True
                    If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                        oPipeInfo.NonPre88 = True
                    End If
                End If

                oPipeInfo.TankID = oTankInfo.TankId
                oPipeInfo.TankSiteID = oTankInfo.TankIndex
                If Not oCompInfo Is Nothing Then
                    oPipeInfo.CompartmentCERCLA = oCompInfo.CCERCLA
                    oPipeInfo.CompartmentFuelType = oCompInfo.FuelTypeId
                    oPipeInfo.CompartmentNumber = oCompInfo.COMPARTMENTNumber
                    oPipeInfo.CompartmentSubstance = oCompInfo.Substance
                End If
                If Date.Compare(oPipeInfo.DateSigned, CDate("01/01/0001")) = 0 Then
                    oPipeInfo.DateSigned = oTankInfo.DateSigned
                End If
                If oPipeInfo.LicenseeID = 0 Then
                    oPipeInfo.LicenseeID = oTankInfo.LicenseeID
                End If
                If oPipeInfo.ContractorID = 0 Then
                    oPipeInfo.ContractorID = oTankInfo.ContractorID
                End If
                oPipeInfo.FacCapStatus = oTankInfo.FacCapStatus

                oTankInfo.pipesCollection.Add(oPipeInfo)
                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "ADD")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oPipe As MUSTER.Info.PipeInfo)
            Try
                Dim dtTempDate As Date = "12-22-1988"
                oPipeInfo = oPipe
                If oPipeInfo.PipeID = 0 Then
                    oPipeInfo.PipeID = nID
                    oPipeInfo.FacilityID = oTankInfo.FacilityId
                    nID -= 1

                    oPipeInfo.DateLastUsed = oTankInfo.DateLastUsed
                    oPipeInfo.PlacedInServiceDate = oTankInfo.PlacedInServiceDate
                End If
                If oPipeInfo.PipeStatusDesc = 426 Then
                    oPipeInfo.POU = True
                    If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                        oPipeInfo.NonPre88 = True
                    End If
                End If

                oPipeInfo.TankID = oTankInfo.TankId
                oPipeInfo.TankSiteID = oTankInfo.TankIndex
                If Not oCompInfo Is Nothing Then
                    oPipeInfo.CompartmentCERCLA = oCompInfo.CCERCLA
                    oPipeInfo.CompartmentFuelType = oCompInfo.FuelTypeId
                    oPipeInfo.CompartmentNumber = oCompInfo.COMPARTMENTNumber
                    oPipeInfo.CompartmentSubstance = oCompInfo.Substance
                End If

                If Date.Compare(oPipeInfo.DateSigned, CDate("01/01/0001")) = 0 Then
                    oPipeInfo.DateSigned = oTankInfo.DateSigned
                End If
                If oPipeInfo.LicenseeID = 0 Then
                    oPipeInfo.LicenseeID = oTankInfo.LicenseeID
                End If
                If oPipeInfo.ContractorID = 0 Then
                    oPipeInfo.ContractorID = oTankInfo.ContractorID
                End If
                oPipeInfo.FacCapStatus = oTankInfo.FacCapStatus

                oTankInfo.pipesCollection.Add(oPipeInfo)
                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "ADD")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As String)
            Dim myIndex As Int16 = 1
            Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Try
                oPipeInfoLocal = oTankInfo.pipesCollection.Item(ID)
                If Not (oPipeInfoLocal Is Nothing) Then
                    oTankInfo.pipesCollection.Remove(oPipeInfoLocal)
                    'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Pipe " & ID & " is not in the collection of Pipes.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oPipe As MUSTER.Info.PipeInfo)
            Try
                If oTankInfo.pipesCollection.Contains(oPipe.ID) Then
                    oTankInfo.pipesCollection.Remove(oPipe)
                End If
                'RaiseEvent evtPipeInfoCompartment(oPipeInfo, "REMOVE")
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            'Throw New Exception("Pipe " & oPipe.ID & " is not in the collection of Pipes.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal bolSaveAsInspection As Boolean = False)
            Dim IDs As New Collection
            Dim delIDs As New Collection
            Dim index As Integer
            Dim xPipeInfo As MUSTER.Info.PipeInfo
            Try
                'Dim colPipesContained As MUSTER.Info.PipesCollection
                'RaiseEvent evtCompInfoPipeCol(colPipesContained)
                For Each xPipeInfo In oTankInfo.pipesCollection.Values
                    If xPipeInfo.IsDirty Then
                        oPipeInfo = xPipeInfo
                        If oPipeInfo.Deleted Then
                            If oPipeInfo.PipeID < 0 Then
                                delIDs.Add(oPipeInfo.PipeID)
                            Else
                                Me.Save(moduleID, staffID, returnVal, UserID, True, False, , bolSaveAsInspection)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Me.ValidateData(False) Then
                                If oPipeInfo.PipeID < 0 Then
                                    IDs.Add(oPipeInfo.PipeID)
                                End If
                                Me.Save(moduleID, staffID, returnVal, UserID, True, False, , bolSaveAsInspection)
                                If Not returnVal = String.Empty Then
                                    Exit Sub
                                End If
                            Else : Exit For
                            End If
                        End If
                        'If oPipeInfo.PipeID < 0 Then
                        '    If oPipeInfo.Deleted Then
                        '        delIDs.Add(oPipeInfo.ID)
                        '    Else
                        '        If Me.ValidateData() Then
                        '            IDs.Add(oPipeInfo.ID)
                        '            oPipeInfo.CreatedBy = UserID
                        '            Me.Save(moduleID, staffID, returnVal, True)
                        '        Else : Exit For
                        '        End If
                        '    End If
                        'Else
                        '    If oPipeInfo.Deleted Then
                        '        delIDs.Add(oPipeInfo.ID)
                        '    End If
                        '    oPipeInfo.ModifiedBy = UserID
                        '    Me.Save(moduleID, staffID, returnVal, True)
                        'End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        xPipeInfo = oTankInfo.pipesCollection.Item(CType(delIDs.Item(index), String))
                        oTankInfo.pipesCollection.Remove(xPipeInfo)
                        'RaiseEvent evtPipeInfoCompartment(xPipeInfo, "REMOVE")
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        oPipeInfo = oTankInfo.pipesCollection.Item(colKey)
                        oTankInfo.pipesCollection.ChangeKey(colKey, oPipeInfo.ID.ToString)
                        'RaiseEvent evtPipeChangeKey(colKey, oPipeInfo.ID)
                    Next
                End If
                'oComments.Flush()
                RaiseEvent evtPipesChanged(Me.IsDirty)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Dim colstrArr As New Collection
            Dim i As Integer
            For Each oPipeInfoLocal In oTankInfo.pipesCollection.Values
                'If oPipeInfoLocal.TankID = oPipeInfo.TankID Then
                colstrArr.Add(oPipeInfoLocal.ID)
                'End If
            Next
            Dim strArr(colstrArr.Count - 1) As String ' = colPipes.GetKeys()
            For i = 0 To colstrArr.Count - 1
                strArr(i) = CType(colstrArr(i + 1), String)
            Next
            strArr.Sort(strArr)
            colIndex = Array.BinarySearch(strArr, Me.ID)
            If colIndex + direction > -1 Then
                If colIndex + direction <= strArr.GetUpperBound(0) Then
                    Return oTankInfo.pipesCollection.Item(strArr.GetValue(colIndex + direction)).ID
                Else
                    Return oTankInfo.pipesCollection.Item(strArr.GetValue(0)).ID
                End If
            Else
                Return oTankInfo.pipesCollection.Item(strArr.GetValue(strArr.GetUpperBound(0))).ID
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oPipeInfo = New MUSTER.Info.PipeInfo
        End Sub
        Public Sub Reset()
            oPipeInfo.Reset()
            'If Date.Compare(oPipeInfo.DateLastUsed, CDate("01/01/0001")) <> 0 Then
            '    CheckDateLastUsed(oPipeInfo.DateLastUsed)
            'End If
            'If oPipeInfo.PipeModDesc <> 0 Then
            '    CheckPipeModDesc(oPipeInfo.PipeModDesc)
            'End If
            'If oPipeInfo.PipeStatusDesc <> 0 Then
            '    CheckPipeMatDesc(oPipeInfo.PipeStatusDesc)
            'End If
            'If Date.Compare(oPipeInfo.LTTDate, CDate("01/01/0001")) <> 0 Then
            'CheckLTTDate(oPipeInfo.LTTDate)
            'End If
            'If Date.Compare(oPipeInfo.ALLDTestDate, CDate("01/01/0001")) <> 0 Then
            'CheckALLDTestDate(oPipeInfo.ALLDTestDate)
            'End If
            'If Date.Compare(oPipeInfo.TermCPLastTested, CDate("01/01/0001")) <> 0 Then
            'CheckTermCPLastTested(oPipeInfo.TermCPLastTested)
            'End If
            'If Date.Compare(oPipeInfo.PipeCPTest, CDate("01/01/0001")) <> 0 Then
            '    CheckPipeCPTest(oPipeInfo.PipeCPTest)
            'End If
            'CheckPipeTermination()
            'If oPipeInfo.PipeTypeDesc <> 0 Then
            '    CheckPipeTypeDesc(oPipeInfo.PipeTypeDesc)
            'End If
            'If oPipeInfo.PipeStatusDesc <> 0 Then
            '    CheckPipeStatus(oPipeInfo.PipeStatusDesc)
            'End If
            'If oPipeInfo.ALLDType <> 0 Then
            '    CheckALLDType(oPipeInfo.ALLDType)
            'End If
            'If oPipeInfo.PipeLD = 245 Then  'line tightness testing
            '    EnableDisablePipeOptions(oPipeInfo.PipeLD)
            'End If
        End Sub
        Public Sub ResetCollection()
            Dim xPipeInf As MUSTER.Info.PipeInfo
            'Dim colPipesContained As MUSTER.Info.PipesCollection
            'RaiseEvent evtCompInfoPipeCol(colPipesContained)
            If oTankInfo.pipesCollection.Count > 0 Then
                For Each xPipeInf In oTankInfo.pipesCollection.Values
                    If xPipeInf.IsDirty Then
                        xPipeInf.Reset()
                    End If
                Next
            End If
            'Need to check with J/ADAM
            'oComments.Reset()
        End Sub
#End Region
#Region "Lookup Operations"
        Public Function PopulatePipeMaterialOfConstruction() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPEMATERIALTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeCPType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPECATHODICPROTECTIONTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Public Function PopulatePipeManufacturer() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPEMANUFACTURER")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeTerminationDispenserType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPEDISPENSERTERMINATIONTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeTerminationDispenserCPType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPECATHODICPROTECTIONTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeTerminationTankType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPIETANKTERMINATIONTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeTerminationTankCPType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPECATHODICPROTECTIONTYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VPIPETYPE")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeReleaseDetection1(ByVal pipeLD As String, Optional ByVal nPipeModDesc As Integer = 0, Optional ByVal After_10_1_08 As Boolean = False) As DataTable
            Try
                Dim dtReturn As DataTable
                Dim str As String = String.Empty

                If After_10_1_08 Then
                    str = " WHERE (Property_ID = 242 or Property_ID = 243)  "
                End If

                If nPipeModDesc = 0 Then
                    dtReturn = GetDataTable(String.Format("VPIPERELEASEDETECTIONTYPE {0}", str), , True, True)
                    'Return dtReturn
                Else
                    If oTankInfo.TankEmergen AndAlso Not After_10_1_08 Then
                        If oPipeInfo.PipeTypeDesc <> 267 Then
                            dtReturn = GetDataTable("VPIPERELEASEDETECTIONTYPE WHERE PROPERTY_ID = 248 ", , True, True)
                        End If
                    End If



                    If dtReturn Is Nothing Then
                        dtReturn = GetDataTable(String.Format("VPIPERELEASEDETECTIONTYPE {0}", str), nPipeModDesc, False, True)
                        Dim dr As DataRow
                        If Not dtReturn Is Nothing AndAlso dtReturn.Rows.Count > 0 Then
                            For Each dr In dtReturn.Rows
                                If dr.Item("Property_id_parent") = nPipeModDesc Then
                                    dtReturn.DefaultView.RowFilter = "Property_Id <> 341"
                                    Exit For
                                End If
                            Next
                        End If

                    End If

                End If

                'RaiseEvent ePipeReleaseDetection1(dtReturn)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeReleaseDetection2(Optional ByVal nPipeLD As Integer = 0, Optional ByVal nPipeModDesc As Integer = 0, Optional ByVal after_10_1_08 As Boolean = False) As DataTable
            Try
                Dim dtReturn As DataTable
                If nPipeLD = 0 Then
                    dtReturn = GetDataTable("vPIPEAUTOMATICLINELEAKDECTIONTYPE", , True)
                    'dtReturn = Me.GetDistinctDataTableListItems(dtReturn)
                    ' Return dtReturn
                Else
                    If oTankInfo.TankEmergen And Pipe.PipeLD = 248 AndAlso Not after_10_1_08 Then
                        'RaiseEvent ecmbPipeReleaseDetection2(False)
                        'RaiseEvent ePickPipeLeakDetectorTeststatus(False)
                    Else
                        dtReturn = GetDataTable("vPIPEAUTOMATICLINELEAKDECTIONTYPE", nPipeLD)
                        Dim dataRow As dataRow
                        For Each dataRow In dtReturn.Rows
                            If dataRow.Item("Property_id_parent") = nPipeLD Then
                                dtReturn.DefaultView.RowFilter = "Property_Id <> 341"
                            End If
                        Next
                    End If
                End If
                If nPipeModDesc <> 261 Then
                    If dtReturn.Select("Property_Id = 498").Length > 0 Then
                        '        Dim dataRow As DataRow = dtReturn.Select("Property_Id = 498")(0)
                        '       dtReturn.Rows.Remove(datarow)
                    End If
                End If
                'RaiseEvent ePipeReleaseDetection2(dtReturn)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeSecondaryOption(Optional ByVal nVal As Int64 = 0, Optional ByVal nPipeMatDesc As Integer = 0) As DataTable
            Dim dtReturn As DataTable
            Try
                If nVal = 0 Then
                    If nPipeMatDesc <> 0 Then
                        dtReturn = GetDataTable_Status_Based("vPIPESECONDARYOPTIONTYPE", oPipeInfo.PipeStatusDesc)
                    Else
                        If oPipeInfo.PipeStatusDesc = 424 Or oPipeInfo.PipeStatusDesc = 429 Then ' CIU / TOSI
                            dtReturn = Nothing
                            'RaiseEvent eSecondaryOption(dtReturn)
                            Return dtReturn
                        Else
                            dtReturn = GetDataTable("vPIPESECONDARYOPTIONTYPE", , True)
                            'Return dtReturn
                            If oPipeInfo.PipeStatusDesc = 425 Or oPipeInfo.PipeStatusDesc = 426 Then ' TOS / POU
                                'If Not (dtReturn Is Nothing) Then
                                'dtReturn = Me.GetDistinctDataTableListItems(dtReturn)
                                'End If
                                'RaiseEvent eSecondaryOption(dtReturn)
                                Return dtReturn
                            End If
                        End If
                    End If
                Else
                    If oPipeInfo.PipeStatusDesc = 424 Or oPipeInfo.PipeStatusDesc = 429 Then ' CIU / TOSI
                        dtReturn = GetDataTable("vPIPESECONDARYOPTIONTYPE", nVal)
                        'RaiseEvent eSecondaryOption(dtReturn)
                        Return dtReturn
                    Else
                        Return PopulatePipeSecondaryOption()
                    End If
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            'Try
            '    If nPipeMatDesc <> 0 Then
            '        dtReturn = GetDataTable_Status_Based("vPIPESECONDARYOPTIONTYPE", nPipeMatDesc)
            '    Else
            '        If nVal = 0 Then
            '            dtReturn = GetDataTable("vPIPESECONDARYOPTIONTYPE")
            '            Return dtReturn
            '        Else
            '            dtReturn = GetDataTable("vPIPESECONDARYOPTIONTYPE", nVal)
            '        End If
            '    End If
            '    RaiseEvent eSecondaryOption(dtReturn)
            'Catch Ex As Exception
            '    If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
            '    Throw Ex
            'End Try
        End Function
        Public Function PopulatePipeSecondaryOptionNew(Optional ByVal nPipeStatus As Integer = 0, Optional ByVal nPipeMatDesc As Integer = 0) As DataTable
            Try
                Dim dt As DataTable
                ' Constraints are only if the pipe status = ciu / tosi. all other statuses contain 
                ' no material of construction to pipe secondary description relationships
                If nPipeStatus = 424 Or nPipeStatus = 429 Then ' CIU OR TOSI
                    dt = GetDataTable("vPIPESECONDARYOPTIONTYPE", nPipeMatDesc)
                Else
                    dt = GetDataTable("vPIPESECONDARYOPTIONTYPE", , True)
                End If
                Return dt
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulatePipeStatus(Optional ByVal nMode As String = "") As DataTable
            Try
                Dim dtReturn As DataTable
                Dim str As String = ""
                Dim dtTempDate As Date = "12-22-1988"
                If UCase(nMode).Trim = "ADD" Then
                    'Remove Unregulated "430" 
                    str = " WHERE PROPERTY_ID <> 430"
                    'If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                    '    str += " AND PROPERTY_ID <> 426"
                    'End If
                    dtReturn = GetDataTable("vPIPESTATUSTYPE" + str)
                ElseIf UCase(nMode).Trim = "EDIT" Then
                    If oPipeInfo.POU Then
                        dtReturn = GetDataTable("vPIPESTATUSTYPE WHERE PROPERTY_ID = 426")
                    Else
                        str = ""
                        ' str = " WHERE PROPERTY_ID <> 430"
                        If oPipeInfo.PipeStatusDesc = 424 Or oPipeInfo.PipeStatusDesc = 425 Or oPipeInfo.PipeStatusDesc = 429 Then
                            dtReturn = GetDataTable("vPIPESTATUSTYPE" + str)
                        Else
                            dtReturn = GetDataTable("vPIPESTATUSTYPE WHERE PROPERTY_ID = 426")
                        End If
                        'If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                        '    str += " AND PROPERTY_ID <> 426"
                        'End If
                        'dtReturn = GetDataTable("vPIPESTATUSTYPE" + str)
                    End If
                Else
                    dtReturn = GetDataTable("vPIPESTATUSTYPE")
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulatePipeStatusShort(Optional ByVal nMode As String = "") As DataTable
            Try
                Dim dtReturn As DataTable
                Dim str As String = ""
                Dim dtTempDate As Date = "12-22-1988"
                If UCase(nMode).Trim = "ADD" Then
                    'Remove Unregulated "430" 
                    str = " WHERE PROPERTY_ID <> 430"
                    'If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                    '    str += " AND PROPERTY_ID <> 426"
                    'End If
                    dtReturn = GetDataTable("vPIPESTATUSTYPESHORT" + str)
                ElseIf UCase(nMode).Trim = "EDIT" Then
                    If oPipeInfo.POU Then
                        dtReturn = GetDataTable("vPIPESTATUSTYPESHORT WHERE PROPERTY_ID = 426")
                    Else
                        str = ""
                        ' str = " WHERE PROPERTY_ID <> 430"
                        If oPipeInfo.PipeStatusDesc = 424 Or oPipeInfo.PipeStatusDesc = 425 Or oPipeInfo.PipeStatusDesc = 429 Then
                            dtReturn = GetDataTable("vPIPESTATUSTYPESHORT" + str)
                        Else
                            dtReturn = GetDataTable("vPIPESTATUSTYPESHORT WHERE PROPERTY_ID = 426")
                        End If
                        'If Date.Compare(oPipeInfo.DateLastUsed, dtTempDate) >= 0 Then
                        '    str += " AND PROPERTY_ID <> 426"
                        'End If
                        'dtReturn = GetDataTable("vPIPESTATUSTYPE" + str)
                    End If
                Else
                    dtReturn = GetDataTable("vPIPESTATUSTYPESHORT")
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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
        Public Function PopulateTankPipeClosureStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTANKPIPECLOSURESTATUS")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateTankPipeInertFill() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vINNERTMATERIAL")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal nVal As Int64 = 0, Optional ByVal bolDistinct As Boolean = False, Optional ByRef GroupByParentID As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Dim groupBy As String = String.Empty
            If bolDistinct Then
                strSQL = "SELECT DISTINCT PROPERTY_ID, PROPERTY_NAME FROM " + strProperty
            Else
                If nVal <> 0 Or Not GroupByParentID Then
                    strSQL = "SELECT * FROM " & strProperty
                Else
                    strSQL = String.Format("SELECT PROPERTY_ID, PROPERTY_NAME, max(PROPERTY_ID_PARENT) as PROPERTY_ID_PARENT  FROM {0}", strProperty)

                    groupBy = " GROUP BY PROPERTY_ID, PROPERTY_NAME"

                End If

            End If

            If nVal <> 0 AndAlso strSQL.IndexOf("PROPERTY_ID_PARENT = ") <= -1 Then
                'Release Detection Group 1 has like ‘%Interstitial Monitoring%’ available
                'only if Secondary Pipe Option is like ‘%Double-Walled%’.
                If (nVal <> 263) And (nVal <> 261) Then ' 263 and 261 is like "Double-Walled"
                    strSQL = strSQL + " WHERE (property_id <> 242 and property_id <> 243) and PROPERTY_ID_PARENT = ".Replace(" WHERE ", IIf(strSQL.IndexOf(" WHERE") > -1, " AND ", " WHERE ")) + nVal.ToString()
                Else
                    strSQL = strSQL + " WHERE PROPERTY_ID_PARENT = ".Replace(" WHERE ", IIf(strSQL.IndexOf(" WHERE") > -1, " AND ", " WHERE ")) + nVal.ToString()
                End If
            End If
            Try

                strSQL = String.Format("{0}{1}", strSQL, groupBy)


                dsReturn = oPipeDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If

                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDataTable_Status_Based(ByVal strProperty As String, Optional ByVal nVal As Int64 = 0) As DataTable
            'Currently In Use	424
            'Temporarily Out of Service	425
            'Permanently Out of Use	426
            'Registration Pending 428
            'Temporarily Out of Service Indefinitely	429
            'Unregulated(430)
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try
                If nVal = 424 Or nVal = 429 Then 'Pipe Status is CIU or TOSI
                    strSQL = "SELECT * FROM " & strProperty
                    If nVal <> 0 Then
                        strSQL += " WHERE PROPERTY_ID_PARENT = '" + Me.PipeMatDesc.ToString + "'"
                    End If
                Else
                    '
                    ' Return ALL Secondary Option values
                    '
                    strSQL = "SELECT * FROM tblSYS_PROPERTY_MASTER WHERE PROPERTY_TYPE_ID = 31"
                End If
                dsReturn = oPipeDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function PopulatePipeInstaller() As DataTable
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
                dsReturn = oPipeDB.DBGetCompanyDetails(LicenseeID)
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

                dsReturn = oPipeDB.DBGetDS(strSQL)
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

        Public Function GetPipeExtensions(ByVal parentPipeID As Integer) As DataTable

            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = String.Format("select * from tblREG_PIPE_EXTENSION where Parent_Pipe_ID = {0}", parentPipeID)

                dsReturn = oPipeDB.DBGetDS(strSQL)
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
#Region "External Event Handlers"
        'Private Sub PipesChanged(ByVal strSrc As String) Handles colPipes.InfoChanged
        '    RaiseEvent evtPipesChanged(Me.colIsDirty)
        'End Sub
        Private Sub PipeChanged(ByVal bolValue As Boolean) Handles oPipeInfo.evtPipeInfoChanged
            RaiseEvent evtPipeChanged(bolValue)
        End Sub
        'Private Sub PipeCommentsChanged(ByVal bolValue As Boolean) Handles oComments.InfoBecameDirty
        '    RaiseEvent evtPipeCommentsChanged(bolValue)
        'End Sub
#End Region
#Region "RaisingEvents for Enable/Disabling controls in the Form"
        'Private Sub CheckDateLastUsed(ByVal dtPickPipeLastUsed As Date)
        '    Try
        '        If Me.PipeStatusDesc = 426 Then   'permanently out of use
        '            Dim dtTempDate As Date = "12-22-88"
        '            If Date.Compare(dtPickPipeLastUsed, dtTempDate) < 0 Then
        '                RaiseEvent ecmbPipeClosureType(True)
        '                RaiseEvent ecmbPipeInertFill(True)
        '            Else
        '                RaiseEvent ecmbPipeClosureType(False)
        '                RaiseEvent ecmbPipeInertFill(False)
        '            End If
        '        Else
        '            RaiseEvent ecmbPipeClosureType(False)
        '            RaiseEvent ecmbPipeInertFill(False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckPipeModDesc(ByVal PipeModDesc As Integer)
        '    Try
        '        EnableDisablePipeOptions(PipeModDesc)
        '        PopulatePipeReleaseDetection1(PipeModDesc)
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try

        'End Sub
        'Private Sub CheckPipeMatDesc(ByVal PipeStatusDesc As Integer)
        '    Try
        '        If PipeStatusDesc = 424 Or PipeStatusDesc = 429 Then   'Currently In Use or 'Temporarily Out of Service Indefinitely
        '            If oPipeInfo.PipeMatDesc <> -1 Then
        '                PopulatePipeSecondaryOption(PipeMatDesc)
        '            End If
        '        End If
        '        If PipeMatDesc = 255 Then ' unknown
        '            RaiseEvent ePipeCPTypeEnable(False, -1)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try

        'End Sub
        'Private Sub CheckLTTDate(ByVal dtPickPipeTightnessTest As Date)
        '    Try
        '        If Me.PipeTypeDesc = 268 Then 'u.s. suction
        '            If (Date.Compare(dtPickPipeTightnessTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, -3, Today()), dtPickPipeTightnessTest) < 0) Then
        '                RaiseEvent ePickPipeTightnessTestMessage("Pipe Tightness Test Date cannot be greater than 3 years old")
        '                'If Date.Compare(olddtPickPipeTightnessTest, CDate("01/01/0001")) = 0 Then
        '                '    oPipeInfo.LTTDate = System.DateTime.Now
        '                'Else
        '                '    oPipeInfo.LTTDate = olddtPickPipeTightnessTest
        '                'End If
        '                'RaiseEvent ePickPipeTightnessTest(False)
        '                Exit Sub
        '            ElseIf (Date.Compare(dtPickPipeTightnessTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, Today(), dtPickPipeTightnessTest) > 0) Then
        '                RaiseEvent ePickPipeTightnessTestMessage("Pipe Tightness Test Date cannot be greater than Today")
        '                'If Date.Compare(olddtPickPipeTightnessTest, CDate("01/01/0001")) = 0 Then
        '                '    oPipeInfo.LTTDate = System.DateTime.Now
        '                'Else
        '                '    oPipeInfo.LTTDate = olddtPickPipeTightnessTest
        '                'End If
        '                'RaiseEvent ePickPipeTightnessTest(False)
        '                Exit Sub
        '            End If
        '        Else
        '            If (Date.Compare(dtPickPipeTightnessTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, -1, Today()), dtPickPipeTightnessTest) < 0) Then
        '                RaiseEvent ePickPipeTightnessTestMessage("Pipe Tightness Test Date cannot be greater than 1 year old")
        '                'If Date.Compare(olddtPickPipeTightnessTest, CDate("01/01/0001")) = 0 Then
        '                '    oPipeInfo.LTTDate = System.DateTime.Now
        '                'Else
        '                '    oPipeInfo.LTTDate = olddtPickPipeTightnessTest
        '                'End If
        '                'RaiseEvent ePickPipeTightnessTest(False)
        '                Exit Sub
        '            ElseIf (Date.Compare(dtPickPipeTightnessTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, Today(), dtPickPipeTightnessTest) > 0) Then
        '                RaiseEvent ePickPipeTightnessTestMessage("Pipe Tightness Test Date cannot be greater than Today")
        '                'If Date.Compare(olddtPickPipeTightnessTest, CDate("01/01/0001")) = 0 Then
        '                '    oPipeInfo.LTTDate = System.DateTime.Now
        '                'Else
        '                '    oPipeInfo.LTTDate = olddtPickPipeTightnessTest
        '                'End If
        '                'RaiseEvent ePickPipeTightnessTest(False)
        '                Exit Sub
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckALLDTestDate(ByVal dtPickPipeLeakDetectorTest As Date)
        '    Try
        '        If (Date.Compare(dtPickPipeLeakDetectorTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, -1, Today()), dtPickPipeLeakDetectorTest) < 0) Then
        '            RaiseEvent edtPickPipeLeakDetectorTestMessage("Pipe Leak Detector Test Date cannot be more than 1 year old")
        '            'RaiseEvent edtPickPipeLeakDetectorTest(False)
        '            Exit Sub
        '            '' elseif commented by manju since it is handled in the ui
        '            ''ElseIf (Date.Compare(dtPickPipeLeakDetectorTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, Today(), dtPickPipeLeakDetectorTest) > 0) Then
        '            ''RaiseEvent edtPickPipeLeakDetectorTestMessage("Pipe Leak Detector Test Date cannot be greater than Today")
        '            '' end comment by manju
        '            'RaiseEvent edtPickPipeLeakDetectorTest(False)
        '            Exit Sub
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckTermCPLastTested(ByVal dtPickPipeTerminationCPLastTested As Date)
        '    Try
        '        If (Date.Compare(dtPickPipeTerminationCPLastTested, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, -3, Today()), dtPickPipeTerminationCPLastTested) < 0) Then
        '            RaiseEvent edtPickPipeTerminationCPLastTestedMessage("Pipe Termination CP Last Tested Date cannot be more than 3 years old")
        '            'RaiseEvent edtPickPipeTerminationCPLastTested(False)
        '            Exit Sub
        '        ElseIf (Date.Compare(dtPickPipeTerminationCPLastTested, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, Today(), dtPickPipeTerminationCPLastTested) > 0) Then
        '            RaiseEvent edtPickPipeTerminationCPLastTestedMessage("Pipe Termination CP Last Tested Date cannot be greater than Today")
        '            'RaiseEvent edtPickPipeTerminationCPLastTested(False)
        '            Exit Sub
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckPipeCPTest(ByVal dtPickPipeCPLastTest As Date)
        '    Try
        '        If (Date.Compare(dtPickPipeCPLastTest, CDate("01/01/0001")) <> 0) And (DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, -3, Today()), dtPickPipeCPLastTest) < 0) Then
        '            RaiseEvent ePipeCPTestMessage("Pipe CP Last Tested Date cannot be more than 3 years old")
        '            'RaiseEvent edtPickPipeCPLastTest(False)
        '            Exit Sub
        '            'ElseIf DateDiff(DateInterval.Day, Today(), dtPickPipeCPLastTest) > 0 Then
        '            'RaiseEvent ePipeCPTestMessage("Pipe CP Last Tested Date cannot be greater than Today")
        '            'RaiseEvent edtPickPipeCPLastTest(False)
        '            'Exit Sub
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckPipeTermination()
        '    Try
        '        If Me.TermTypeDisp = 611 Or Me.TermTypeTank = 610 Then
        '            If Me.TermTypeDisp = 611 Then
        '                RaiseEvent ecmbPipeTerminationDispenserCPType(True)
        '            Else
        '                RaiseEvent ecmbPipeTerminationDispenserCPType(False)
        '            End If
        '            If Me.TermTypeTank = 610 Then
        '                RaiseEvent ecmbPipeTerminationTankCPType(True)
        '            Else
        '                RaiseEvent ecmbPipeTerminationTankCPType(False)
        '            End If
        '            RaiseEvent edtPickPipeTerminationCPInstalled(True)
        '            RaiseEvent edtPickPipeTerminationCPLastTested(True)
        '        Else
        '            RaiseEvent ecmbPipeTerminationDispenserCPType(False)
        '            RaiseEvent ecmbPipeTerminationTankCPType(False)
        '            RaiseEvent edtPickPipeTerminationCPInstalled(False)
        '            RaiseEvent edtPickPipeTerminationCPLastTested(False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try

        'End Sub
        'Private Sub CheckPipeTypeDesc(ByVal nPipeTypeDesc As Integer)
        '    Try
        '        If nPipeTypeDesc = 266 Then 'pressurized
        '            RaiseEvent ePipeReleaseDetectionstatus1(-1, True)
        '            RaiseEvent ePipeReleaseDetectionstatus2(-1, True)
        '        ElseIf nPipeTypeDesc = 267 Then 'safe suction
        '            RaiseEvent ePipeReleaseDetectionstatus1(-1, False)
        '            RaiseEvent ePipeReleaseDetectionstatus2(-1, False)
        '        ElseIf nPipeTypeDesc = 268 Then 'u.s. suction
        '            RaiseEvent ePipeReleaseDetectionstatus1(-1, True)
        '            RaiseEvent ePipeReleaseDetectionstatus2(-1, False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckPipeStatus(ByVal nPipeStatus As Integer)
        '    Dim dtTempDate As Date = "12-22-88"
        '    Try
        '        RaiseEvent enabledisablepipecontrols(False)
        '        If nPipeStatus = 424 Or nPipeStatus = 429 Then ' CIU OR TOSI
        '            PopulatePipeSecondaryOption(, oPipeInfo.PipeMatDesc)
        '            If nPipeStatus = 429 Then
        '                RaiseEvent edtPickPipeLastUsed(True)
        '                RaiseEvent edtPickPipeLastUsedFocus(Me.ToString)
        '            End If
        '            RaiseEvent ecmbPipeClosureType(False)
        '            RaiseEvent ecmbPipeInertFill(False)
        '        Else
        '            PopulatePipeSecondaryOption()
        '            If nPipeStatus = 425 Then  ' TOS
        '                RaiseEvent edtPickPipeLastUsed(True)
        '                RaiseEvent edtPickPipeLastUsedFocus(Me.ToString)
        '                RaiseEvent ecmbPipeClosureType(False)
        '                RaiseEvent ecmbPipeInertFill(False)
        '            End If
        '            If nPipeStatus = 426 Then ' POU
        '                If oPipeInfo.POU And oPipeInfo.NonPre88 Then
        '                    RaiseEvent edtPickPipeLastUsed(False)
        '                Else
        '                    RaiseEvent edtPickPipeLastUsed(True)
        '                End If
        '                RaiseEvent edtPickPipeLastUsedFocus(Me.ToString)
        '                If Date.Compare(DateLastUsed, dtTempDate) < 0 Then
        '                    RaiseEvent ecmbPipeClosureType(True)
        '                    RaiseEvent ecmbPipeInertFill(True)
        '                Else
        '                    RaiseEvent ecmbPipeClosureType(False)
        '                    RaiseEvent ecmbPipeInertFill(False)
        '                End If
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Private Sub CheckALLDType(ByVal nVal As Integer)
        '    Try
        '        If nVal = 497 Then 'Electronic
        '            RaiseEvent ePipeReleaseDetectiontext1("Electronic ALLD with 0.2 Test")
        '        End If
        '        If nVal = 496 Then '"mechanical"
        '            RaiseEvent ePickPipeLeakDetectorTeststatus(True)
        '        Else
        '            RaiseEvent ePickPipeLeakDetectorTeststatus(False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        'Public Sub EnableDisablePipeOptions(ByVal nVal As Integer)
        '    Try
        '        If nVal = 260 Then ' cathodically protected
        '            RaiseEvent ePipeCPTypeEnable(True, -1)
        '            RaiseEvent edtPickPipeCPLastTest(True)
        '            RaiseEvent ePickPipeCPInstalled(True)
        '        Else
        '            RaiseEvent ePipeCPTypeEnable(False, -1)
        '            RaiseEvent edtPickPipeCPLastTest(False)
        '            RaiseEvent ePickPipeCPInstalled(False)
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try

        'End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oPipeInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Pipe ID")
                tbEntityTable.Columns.Add("Pipe Index")
                tbEntityTable.Columns.Add("Facility_ID")
                tbEntityTable.Columns.Add("Tank ID")
                tbEntityTable.Columns.Add("ALLD Test")
                tbEntityTable.Columns.Add("ALLD Test Date")
                tbEntityTable.Columns.Add("Cas Number")
                tbEntityTable.Columns.Add("Closure Status Desc")
                tbEntityTable.Columns.Add("Closure Type")
                tbEntityTable.Columns.Add("Composite Primary")
                tbEntityTable.Columns.Add("Composite Secondary")
                tbEntityTable.Columns.Add("Contain SumpDisp")
                tbEntityTable.Columns.Add("Contain SumpTank")
                tbEntityTable.Columns.Add("Date Closed")
                tbEntityTable.Columns.Add("Date Last Used")
                tbEntityTable.Columns.Add("Date Closure Recd")
                tbEntityTable.Columns.Add("Date Recd")
                tbEntityTable.Columns.Add("Date Signed")
                tbEntityTable.Columns.Add("ALLD Type")
                tbEntityTable.Columns.Add("Inert Material")
                tbEntityTable.Columns.Add("LCP Install Date")
                tbEntityTable.Columns.Add("Licensee ID")
                tbEntityTable.Columns.Add("Contractor ID")
                tbEntityTable.Columns.Add("LTT Date")
                tbEntityTable.Columns.Add("Pipe CP Test")
                tbEntityTable.Columns.Add("Pipe CP Type")
                tbEntityTable.Columns.Add("Pipe Install Date")
                tbEntityTable.Columns.Add("Pipe LD")
                tbEntityTable.Columns.Add("Pipe Manufacturer")
                tbEntityTable.Columns.Add("Pipe Mat Desc")
                tbEntityTable.Columns.Add("Pipe Mod Desc")
                tbEntityTable.Columns.Add("Pipe Other Material")
                tbEntityTable.Columns.Add("Pipe Status Desc")
                tbEntityTable.Columns.Add("Pipe Type Desc")
                tbEntityTable.Columns.Add("Piping Comments")
                tbEntityTable.Columns.Add("Pipe Installation Planned For")
                tbEntityTable.Columns.Add("Placed In Service Date")
                tbEntityTable.Columns.Add("Substance Comments")
                tbEntityTable.Columns.Add("Substance Desc")
                tbEntityTable.Columns.Add("Term CP Last Tested")
                tbEntityTable.Columns.Add("Term CP Type Tank")
                tbEntityTable.Columns.Add("Term CP Type Disp")
                tbEntityTable.Columns.Add("Pipe CP Installed Date")
                tbEntityTable.Columns.Add("Termination CP Installed Date")
                tbEntityTable.Columns.Add("Termination Type Disp")
                tbEntityTable.Columns.Add("Termination Type Tank")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")
                tbEntityTable.Columns.Add("Parent Pipe ID")

                For Each oPipeInfoLocal In oTankInfo.pipesCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oPipeInfoLocal.ID
                    dr("Pipe Index") = oPipeInfoLocal.Index
                    dr("Facility_ID") = oPipeInfoLocal.FacilityID
                    dr("Tank ID") = oPipeInfoLocal.TankID
                    dr("ALLD Test") = oPipeInfoLocal.ALLDTestDate
                    dr("ALLD Test Date") = oPipeInfoLocal.ALLDTestDate
                    dr("Cas Number") = oPipeInfoLocal.CASNumber
                    dr("Closure Status Desc") = oPipeInfoLocal.ClosureStatusDesc
                    dr("Closure Type") = oPipeInfoLocal.ClosureType
                    dr("Composite Primary") = oPipeInfoLocal.CompPrimary
                    dr("Composite Secondary") = oPipeInfoLocal.CompSecondary
                    dr("Contain SumpDisp") = oPipeInfoLocal.ContainSumpDisp
                    dr("Contain SumpTank") = oPipeInfoLocal.ContainSumpTank
                    dr("Date Closed") = oPipeInfoLocal.DateClosed
                    dr("Date Last Used") = oPipeInfoLocal.DateLastUsed
                    dr("Date Closure Recd") = oPipeInfoLocal.DateClosureRecd
                    dr("Date Recd") = oPipeInfoLocal.DateRecd
                    dr("Date Signed") = oPipeInfoLocal.DateSigned
                    dr("ALLD Type") = oPipeInfoLocal.ALLDType
                    dr("Inert Material") = oPipeInfoLocal.InertMaterial
                    dr("LCP Install Date") = oPipeInfoLocal.LCPInstallDate
                    dr("Licensee ID") = oPipeInfoLocal.LicenseeID
                    dr("Contractor ID") = oPipeInfoLocal.ContractorID
                    dr("LTT Date") = oPipeInfoLocal.LTTDate
                    dr("Pipe CP Test") = oPipeInfoLocal.PipeCPTest
                    dr("Shear Test") = oPipeInfoLocal.DateShearTest
                    dr("Pipe Sec Insp") = oPipeInfoLocal.DatePipeSecInsp
                    dr("Pipe Elec Insp") = oPipeInfoLocal.DatePipeElecInsp
                    dr("Pipe CP Type") = oPipeInfoLocal.PipeCPType
                    dr("Pipe Install Date") = oPipeInfoLocal.PipeInstallDate
                    dr("Pipe LD") = oPipeInfoLocal.PipeLD
                    dr("Pipe Manufacturer") = oPipeInfoLocal.PipeManufacturer
                    dr("Pipe Mat Desc") = oPipeInfoLocal.PipeMatDesc
                    dr("Pipe Mod Desc") = oPipeInfoLocal.PipeModDesc
                    dr("Pipe Other Material") = oPipeInfoLocal.PipeOtherMaterial
                    dr("Pipe Status Desc") = oPipeInfoLocal.PipeStatusDesc
                    dr("Pipe Type Desc") = oPipeInfoLocal.PipeTypeDesc
                    dr("Piping Comments") = oPipeInfoLocal.PipingComments
                    dr("Pipe Installation Planned For") = oPipeInfoLocal.PipeInstallationPlannedFor
                    dr("Placed In Service Date") = oPipeInfoLocal.PlacedInServiceDate
                    dr("Substance Comments") = oPipeInfoLocal.SubstanceComments
                    dr("Substance Desc") = oPipeInfoLocal.SubstanceDesc
                    dr("Term CP Last Tested") = oPipeInfoLocal.TermCPLastTested
                    dr("Term CP Type Tank") = oPipeInfoLocal.TermCPTypeTank
                    dr("Term CP Type Disp") = oPipeInfoLocal.TermCPTypeDisp
                    dr("Pipe CP Installed Date") = oPipeInfoLocal.PipeCPInstalledDate
                    dr("Termination CP Installed Date") = oPipeInfoLocal.TermCPInstalledDate
                    dr("Termination Type Disp") = oPipeInfoLocal.TermTypeDisp
                    dr("Termination Type Tank") = oPipeInfoLocal.TermTypeTank
                    dr("Deleted") = oPipeInfoLocal.Deleted
                    dr("Created By") = oPipeInfoLocal.CreatedBy
                    dr("Date Created") = oPipeInfoLocal.CreatedOn
                    dr("Last Edited By") = oPipeInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oPipeInfoLocal.ModifiedOn
                    dr("Parent Pipe ID") = oPipeInfoLocal.ParentPipeID
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PipeCAPTable(ByVal nFacilityId As Integer) As DataTable
            Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try

                tbEntityTable.Columns.Add("Facility ID", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("TANK ID", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("Pipe ID", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("Pipe Index", Type.GetType("System.Int64"))
                tbEntityTable.Columns.Add("STATUS", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("INSTALL DATE", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("CP DATE", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("TT DATE", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("ALLD Test Date", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("ALLD Test", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("TERM CP TEST", Type.GetType("System.DateTime"))
                tbEntityTable.Columns.Add("DISP CP TYPE", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("TANK CP TYPE", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Pipe Mod Desc", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Pipe LD", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Pipe Type Desc", Type.GetType("System.String"))
                tbEntityTable.Columns.Add("Parent Pipe ID", Type.GetType("System.Int64"))

                For Each oPipeInfoLocal In oTankInfo.pipesCollection.Values
                    If nFacilityId = oPipeInfoLocal.FacilityID And Not (oPipeInfoLocal.Deleted) Then
                        dr = tbEntityTable.NewRow()
                        dr("Facility ID") = oPipeInfoLocal.FacilityID
                        dr("TANK ID") = oPipeInfoLocal.TankID
                        dr("Pipe ID") = oPipeInfoLocal.ID
                        dr("Pipe Index") = oPipeInfoLocal.Index
                        dr("STATUS") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeStatusDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeStatusDesc).Name)
                        dr("INSTALL DATE") = IIf(Date.Compare(oPipeInfoLocal.PipeInstallDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.PipeInstallDate)
                        dr("CP DATE") = IIf(Date.Compare(oPipeInfoLocal.PipeCPTest, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.PipeCPTest)
                        dr("SHEAR TEST") = IIf(Date.Compare(oPipeInfoLocal.DateShearTest, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.DateShearTest)
                        dr("PIPE SEC INSP") = IIf(Date.Compare(oPipeInfoLocal.DatePipeSecInsp, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.DatePipeSecInsp)
                        dr("PIPE ELEC INSP") = IIf(Date.Compare(oPipeInfoLocal.DatePipeElecInsp, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.DatePipeElecInsp)
                        dr("TT DATE") = IIf(Date.Compare(oPipeInfoLocal.LTTDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.LTTDate)
                        dr("ALLD Test Date") = IIf(Date.Compare(oPipeInfoLocal.ALLDTestDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.ALLDTestDate)
                        dr("ALLD Test") = IIf(oProperty.Retrieve(oPipeInfoLocal.ALLDTest).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.ALLDTest).Name)
                        dr("TERM CP TEST") = IIf(Date.Compare(oPipeInfoLocal.TermCPLastTested, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.TermCPLastTested)
                        dr("DISP CP TYPE") = IIf(oProperty.Retrieve(oPipeInfoLocal.TermCPTypeDisp).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.TermCPTypeDisp).Name)
                        dr("TANK CP TYPE") = IIf(oProperty.Retrieve(oPipeInfoLocal.TermCPTypeTank).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.TermCPTypeTank).Name)
                        dr("Pipe Mod Desc") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeModDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeModDesc).Name)
                        dr("Pipe LD") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeLD).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeLD).Name)
                        dr("Pipe Type Desc") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeTypeDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeTypeDesc).Name)
                        dr("Parent Pipe ID") = oPipeInfoLocal.ParentPipeID
                        tbEntityTable.Rows.Add(dr)
                    End If
                Next
                Return tbEntityTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function CheckCAPStatus()
        '    RaiseEvent evtPipeCAPChanged(Me.FacilityID)
        'End Function
        Public Function PipesTable(ByVal facID As Integer) As DataTable
            Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim dtPipeTable As New DataTable
            Try
                'dtPipeTable.Columns.Add("Facility_ID", Type.GetType("System.Int64"))
                dtPipeTable.Columns.Add("Tank_ID", Type.GetType("System.Int64"))
                dtPipeTable.Columns.Add("Pipe_ID", Type.GetType("System.Int64"))
                dtPipeTable.Columns.Add("Pipe Site ID", Type.GetType("System.Int64"))
                dtPipeTable.Columns.Add("Pipe Status", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("Install Date", Type.GetType("System.DateTime"))
                dtPipeTable.Columns.Add("Last Used", Type.GetType("System.DateTime"))
                dtPipeTable.Columns.Add("Type", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("Material", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("Sec Option", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("CP Type", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("LD Group 1", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("LD Group 2", Type.GetType("System.String"))
                dtPipeTable.Columns.Add("Parent Pipe ID", Type.GetType("System.int64"))

                For Each oPipeInfoLocal In oTankInfo.pipesCollection.Values
                    If oPipeInfoLocal.FacilityID = facID And Not (oPipeInfoLocal.Deleted) Then
                        dr = dtPipeTable.NewRow()
                        'dr("Facility_ID") = oPipeInfoLocal.FacilityID
                        dr("Tank_ID") = oPipeInfoLocal.TankID
                        dr("Pipe_ID") = oPipeInfoLocal.PipeID
                        'dr("Pipe_ID") = oPipeInfoLocal.ID
                        dr("Pipe Site ID") = oPipeInfoLocal.Index
                        dr("Pipe Status") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeStatusDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeStatusDesc).Name)
                        dr("Install Date") = IIf(Date.Compare(oPipeInfoLocal.PipeInstallDate, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.PipeInstallDate)
                        dr("Last Used") = IIf(Date.Compare(oPipeInfoLocal.DateLastUsed, CDate("01/01/0001")) = 0, System.DBNull.Value, oPipeInfoLocal.DateLastUsed)
                        dr("Type") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeTypeDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeTypeDesc).Name)
                        dr("Material") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeMatDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeMatDesc).Name)
                        dr("Sec Option") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeModDesc).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeModDesc).Name)
                        dr("CP Type") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeCPType).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeCPType).Name)
                        dr("LD Group 1") = IIf(oProperty.Retrieve(oPipeInfoLocal.PipeLD).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.PipeLD).Name)
                        dr("LD Group 2") = IIf(oProperty.Retrieve(oPipeInfoLocal.ALLDType).Name Is System.DBNull.Value, String.Empty, oProperty.Retrieve(oPipeInfoLocal.ALLDType).Name)
                        dr("Parent Pipe ID") = oPipeInfoLocal.ParentPipeID

                        dtPipeTable.Rows.Add(dr)
                    End If
                Next
                Return dtPipeTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function ExistingPipesTable(ByVal facID As Integer, ByVal tnkID As Integer, ByVal CompNum As Integer) As DataTable
            Try
                Return oPipeDB.DBGetExistingPipes(facID, tnkID, CompNum)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
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
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub ChangePipeTankCompartmentNumberKey(Optional ByVal tnkID As Integer = 0, Optional ByVal compNum As Integer = 0, Optional ByVal pipeID As Integer = 0, Optional ByRef pipeInfo As MUSTER.Info.PipeInfo = Nothing)
            Try
                Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
                If pipeInfo Is Nothing Then
                    pipeInfo = oPipeInfo
                End If
                Dim OldKey() As String = pipeInfo.ID.Split("|")
                If tnkID = 0 Then
                    tnkID = OldKey(0)
                End If
                pipeInfo.TankID = tnkID
                If compNum = 0 Then
                    compNum = OldKey(1)
                End If
                pipeInfo.CompartmentNumber = compNum
                If pipeID = 0 Then
                    pipeID = OldKey(2)
                End If
                pipeInfo.PipeID = pipeID
                oTankInfo.pipesCollection.ChangeKey(OldKey(0).ToString + "|" + OldKey(1).ToString + "|" + OldKey(2).ToString, pipeInfo.ID)
                'RaiseEvent evtPipeChangeKey(OldKey(0).ToString + "|" + OldKey(1).ToString + "|" + OldKey(2).ToString, pipeInfo.ID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#End Region
#Region "Event Handlers"
        'added by kiran
        'Private Sub CommentsCol(ByVal pipeID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection) Handles oComments.evtCommentColPipe
        '    'Dim oPipeInfoLocal As MUSTER.Info.PipeInfo
        '    'Try
        '    '    oPipeInfoLocal = colPipes.Item(pipeID)
        '    '    If Not (oPipeInfoLocal Is Nothing) Then
        '    '        oPipeInfoLocal.commentsCollection = commentsCol
        '    '    End If
        '    'Catch ex As Exception
        '    '    If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '    '    Throw ex
        '    'End Try
        '    RaiseEvent evtPipesCommentsCol(pipeID, oPipeInfo.CompartmentNumber, commentsCol)
        'End Sub
        'end changes
#End Region
    End Class
End Namespace